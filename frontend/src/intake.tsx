import { useEffect, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { useNavigate } from "react-router-dom";
import { api, label, terminal, type Entry, type JobData } from "./api";
import { useCanEdit } from "./accounts";

export type IntakeSummary = {
  total: number;
  matched: number;
  counts: Record<string, number>;
  items: {
    id: string;
    name: string;
    kind: string;
    status: string;
    detail: string;
  }[];
};

export function IntakePanel({
  id,
  onSummary,
}: {
  id: string;
  onSummary?: (value: IntakeSummary) => void;
}) {
  const navigate = useNavigate();
  const canEdit = useCanEdit();
  const [offset, setOffset] = useState(0),
    [status, setStatus] = useState(""),
    [error, setError] = useState<unknown>();
  const job = useQuery({
    queryKey: ["job", id],
    queryFn: () => api<Entry<JobData>>("/jobs/" + id),
    refetchInterval: (query) =>
      terminal.has(query.state.data?.data.status || "") ? false : 2000,
  });
  const plan = useQuery({
    queryKey: ["intake", id, offset, status],
    queryFn: () =>
      api<IntakeSummary>(
        `/jobs/${id}/intake?${new URLSearchParams({ offset: String(offset), status })}`,
      ),
    enabled: job.data?.data.status === "complete",
  });
  useEffect(() => {
    if (plan.data) onSummary?.(plan.data);
  }, [plan.data, onSummary]);
  async function importResult(itemId: string, name: string, kind: string) {
    try {
      const result = await api<Entry<JobData>>("/jobs", {
        mode: kind === "workbook" ? "workbook" : "render",
        name: "Imported " + name,
        input: { kind: "auto", intake_id: id, relative: itemId },
      });
      navigate("/runs/" + result.id);
    } catch (error) {
      setError(error);
    }
  }
  return (
    <section className="panel">
      <h2>Intake inventory</h2>
      <p>
        Recognised originals are processed together. Existing analyses and
        workbooks need an explicit import choice so old reports are not counted
        as new bills.
      </p>
      {!terminal.has(job.data?.data.status || "") && (
        <>
          <p role="status">
            Inspecting files and ZIP contents… You can leave and reconnect from
            Runs & evidence.
          </p>
          <button
            disabled={!canEdit}
            onClick={() =>
              void api(`/jobs/${id}/cancel`, {}).then(() => job.refetch())
            }
          >
            Cancel inspection
          </button>
        </>
      )}
      {Boolean(job.error || job.data?.data.error || plan.error || error) && (
        <p role="alert" className="error">
          {String(job.error || job.data?.data.error || plan.error || error)}
        </p>
      )}
      {plan.data && (
        <>
          <div className="summary-cards">
            {Object.entries(plan.data.counts).map(([key, value]) => (
              <div key={key}>
                <span>{label(key)}</span>
                <strong>{value.toLocaleString()}</strong>
              </div>
            ))}
          </div>
          <label>
            Show files
            <select
              aria-label="Intake file status"
              value={status}
              onChange={(event) => {
                setStatus(event.target.value);
                setOffset(0);
              }}
            >
              <option value="">All classifications</option>
              {Object.keys(plan.data.counts).map((key) => (
                <option key={key} value={key}>
                  {label(key)}
                </option>
              ))}
            </select>
          </label>
          <div className="evidence-table">
            <table aria-label="Intake files">
              <thead>
                <tr>
                  <th>File or archive member</th>
                  <th>Detected type</th>
                  <th>Status</th>
                  <th>Details</th>
                </tr>
              </thead>
              <tbody>
                {plan.data.items.map((item, index) => (
                  <tr key={index}>
                    <td title={item.name}>
                      <a
                        href={`/api/jobs/${id}/intake-source?${new URLSearchParams({ item_id: item.id })}`}
                      >
                        {item.name}
                      </a>
                    </td>
                    <td>{label(item.kind)}</td>
                    <td>{label(item.status)}</td>
                    <td>
                      {item.detail}
                      {item.status === "needs_choice" &&
                        ["bundle", "result", "workbook"].includes(
                          item.kind,
                        ) && (
                          <button
                            disabled={!canEdit}
                            onClick={() =>
                              void importResult(item.id, item.name, item.kind)
                            }
                          >
                            Import this{" "}
                            {item.kind === "workbook" ? "workbook" : "analysis"}
                          </button>
                        )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="toolbar">
            <button
              disabled={!offset}
              onClick={() => setOffset(Math.max(0, offset - 200))}
            >
              Previous files
            </button>
            <span>
              {offset + 1}–{Math.min(offset + 200, plan.data.matched)} of{" "}
              {plan.data.matched.toLocaleString()}
            </span>
            <button
              disabled={offset + 200 >= plan.data.matched}
              onClick={() => setOffset(offset + 200)}
            >
              Next files
            </button>
          </div>
        </>
      )}
    </section>
  );
}
