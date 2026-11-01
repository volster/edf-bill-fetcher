import { StrictMode, Suspense, lazy, useEffect, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  createHashRouter,
  RouterProvider,
  Link,
  NavLink,
  Route,
  Routes,
  useBlocker,
  useNavigate,
  useParams,
} from "react-router-dom";
import {
  QueryClient,
  QueryClientProvider,
  useQuery,
  useQueryClient,
} from "@tanstack/react-query";
import * as Tabs from "@radix-ui/react-tabs";
import {
  FileSearch,
  FolderPlus,
  History,
  LogOut,
  Settings2,
  ShieldCheck,
} from "lucide-react";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { z } from "zod";
import {
  api,
  label,
  setCsrf,
  terminal,
  uploadAsset,
  type Claim,
  type Context,
  type Credit,
  type Entry,
  type JobData,
  type Json,
  type Review,
  type Row,
  type Schema,
  type Source,
} from "./api";
import { EvidenceTable, Value } from "./values";
import "./style.css";
import { IntakePanel, type IntakeSummary } from "./intake";

const Pdf = lazy(() => import("./pdf"));
const queryClient = new QueryClient({
  defaultOptions: { queries: { retry: false, refetchOnWindowFocus: false } },
});
const kinds: Record<string, string> = {
  auto: "All evidence · detect automatically",
  pdf_dir: "PDF bills / SAP dumps",
  pst: "Outlook PST archive",
  ost: "Outlook OST archive",
  htm: "Account history HTM",
  eml_dir: "EML message folder",
  msg_dir: "MSG message folder",
  result: "Saved result JSON",
  workbook: "Amended review workbook",
  bundle: "Complete evidence pack ZIP",
};

function ErrorMessage({ error }: { error: unknown }) {
  return error ? (
    <p role="alert" className="error">
      {error instanceof Error ? error.message : String(error)}
    </p>
  ) : null;
}
function useSchema() {
  return useQuery({
    queryKey: ["schema"],
    queryFn: () => api<Schema>("/schema"),
  });
}
function usePending(dirty: boolean) {
  const blocker = useBlocker(dirty);
  useEffect(() => {
    if (blocker.state === "blocked") {
      if (window.confirm("Discard unapplied form edits and leave this screen?"))
        blocker.proceed();
      else blocker.reset();
    }
  }, [blocker]);
  useEffect(() => {
    const warn = (event: BeforeUnloadEvent) => {
      if (dirty) event.preventDefault();
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [dirty]);
}

const loginSchema = z.object({
  username: z.string().min(1, "Enter your username"),
  password: z.string().min(1, "Enter your password"),
});
function SignIn({
  onSuccess,
}: {
  onSuccess: (value: { username: string; csrf: string }) => void;
}) {
  const [error, setError] = useState<unknown>();
  const {
    register,
    handleSubmit,
    formState: { errors, isSubmitting },
  } = useForm<z.infer<typeof loginSchema>>({
    resolver: zodResolver(loginSchema),
  });
  return (
    <main className="signin">
      <div className="brand">
        <FileSearch size={32} /> EDF Evidence Collector
      </div>
      <h1>Your evidence, clearly connected.</h1>
      <p>
        Sign in to collect bills, review original documents and build your
        complaint pack.
      </p>
      <form
        onSubmit={handleSubmit(async (values) => {
          try {
            onSuccess(await api("/login", values));
          } catch (error) {
            setError(error);
          }
        })}
      >
        <label>
          Username
          <input autoFocus autoComplete="username" {...register("username")} />
        </label>
        {errors.username && <p role="alert">{errors.username.message}</p>}
        <label>
          Password
          <input
            type="password"
            autoComplete="current-password"
            {...register("password")}
          />
        </label>
        {errors.password && <p role="alert">{errors.password.message}</p>}
        <ErrorMessage error={error} />
        <button className="primary" disabled={isSubmitting}>
          Sign in
        </button>
      </form>
    </main>
  );
}

function App() {
  const [account, setAccount] = useState<{
      username: string;
      csrf: string;
    } | null>(null),
    [loaded, setLoaded] = useState(false),
    [expired, setExpired] = useState(false);
  useEffect(() => {
    const expire = () => setExpired(true);
    window.addEventListener("edf-session-expired", expire);
    return () => window.removeEventListener("edf-session-expired", expire);
  }, []);
  useEffect(() => {
    void api<{ username: string; csrf: string }>("/session")
      .then((value) => {
        setCsrf(value.csrf);
        setAccount(value);
      })
      .finally(() => setLoaded(true))
      .catch(() => {});
  }, []);
  if (!loaded)
    return (
      <p className="loading" role="status">
        Opening workspace…
      </p>
    );
  if (!account)
    return (
      <SignIn
        onSuccess={(value) => {
          setCsrf(value.csrf);
          setAccount(value);
          setExpired(false);
        }}
      />
    );
  return (
    <>
      <div className="app" inert={expired}>
        <aside className="sidebar">
          <Link to="/" className="brand">
            <FileSearch /> EDF Evidence<span>Collector</span>
          </Link>
          <nav aria-label="Main navigation">
            <NavLink to="/" end>
              <History /> Runs & evidence
            </NavLink>
            <NavLink to="/new">
              <FolderPlus /> Build a pack
            </NavLink>
            <NavLink to="/compare">
              <FileSearch /> Compare runs
            </NavLink>
            <NavLink to="/diagnostics">
              <ShieldCheck /> Readiness & build
            </NavLink>
          </nav>
          <div className="account">
            <span>{account.username}</span>
            <button
              onClick={async () => {
                if (
                  !window.confirm(
                    "Sign out? Apply pending form edits before leaving.",
                  )
                )
                  return;
                try {
                  await api("/logout", {});
                } catch {
                  /* Local sign-out also works after session expiry. */
                }
                setCsrf("");
                queryClient.clear();
                setAccount(null);
              }}
            >
              <LogOut size={16} /> Sign out
            </button>
          </div>
        </aside>
        <main className="content">
          <Routes>
            <Route path="/" element={<Runs />} />
            <Route path="/new" element={<Build />} />
            <Route path="/runs/:id" element={<Run />} />
            <Route path="/compare" element={<Compare />} />
            <Route path="/diagnostics" element={<Diagnostics />} />
          </Routes>
        </main>
      </div>
      {expired && (
        <div
          className="session-overlay"
          role="dialog"
          aria-modal="true"
          aria-label="Session expired"
        >
          <div>
            <p className="session-notice">
              Your session expired. Sign in to continue; pending form edits are
              kept.
            </p>
            <SignIn
              onSuccess={(value) => {
                setCsrf(value.csrf);
                setAccount(value);
                setExpired(false);
              }}
            />
          </div>
        </div>
      )}
    </>
  );
}

function useJobs() {
  return useQuery({
    queryKey: ["jobs"],
    queryFn: () => api<Entry<JobData>[]>("/jobs"),
    refetchInterval: 3000,
  });
}
function Runs() {
  const query = useJobs();
  return (
    <>
      <header className="page-heading">
        <div className="eyebrow">Shared workspace</div>
        <h1>Runs & evidence</h1>
        <p>
          Each run keeps its original documents, accepted analysis and reviewed
          outputs together.
        </p>
        <Link className="button primary" to="/new">
          Build evidence pack
        </Link>
      </header>
      <ErrorMessage error={query.error} />
      {query.isPending && <p role="status">Loading runs…</p>}
      {!query.isPending && !query.data?.length && (
        <div className="empty">
          <FolderPlus size={36} />
          <h2>Start with your original evidence</h2>
          <p>
            Add bills, account history or email archives. You can also import an
            existing pack.
          </p>
        </div>
      )}
      <div className="run-list">
        {query.data
          ?.slice()
          .reverse()
          .map((run) => (
            <Link key={run.id} to={"/runs/" + run.id} className="run-card">
              <div>
                <h2>{run.data.request.name}</h2>
                <p>
                  {label(run.data.request.mode)} ·{" "}
                  {new Date(run.data.created * 1000).toLocaleString()}
                </p>
              </div>
              <span className={"badge " + run.data.status}>
                {label(run.data.status)}
              </span>
            </Link>
          ))}
      </div>
    </>
  );
}

function Sources({
  onAdd,
  allowed,
}: {
  onAdd: (source: Source, name: string) => void;
  allowed?: string[];
}) {
  const [selection, setSelection] = useState("files");
  const [kind, setKind] = useState<Source["kind"]>(
      (allowed?.[0] || "pdf_dir") as Source["kind"],
    ),
    [mode, setMode] = useState("upload"),
    [files, setFiles] = useState<File[]>([]),
    [name, setName] = useState(""),
    [progress, setProgress] = useState(""),
    [busy, setBusy] = useState(false),
    [error, setError] = useState<unknown>(),
    [root, setRoot] = useState(""),
    [relative, setRelative] = useState("");
  const roots = useQuery({
    queryKey: ["roots"],
    queryFn: () => api<{ id: string; name: string }[]>("/roots"),
  });
  const assets = useQuery({
    queryKey: ["assets"],
    queryFn: () =>
      api<Entry<{ name: string; kind: string; sealed: boolean }>[]>("/assets"),
  });
  const browser = useQuery({
    queryKey: ["browse", root, relative],
    queryFn: () =>
      api<{ name: string; relative: string; directory: boolean }[]>(
        `/roots/${root}/browse?${new URLSearchParams({ relative })}`,
      ),
    enabled: mode === "server" && !!root,
  });
  const [existing, setExisting] = useState("");
  async function send(resume = false) {
    setBusy(true);
    setError(undefined);
    try {
      const saved = resume
        ? JSON.parse(sessionStorage.getItem("edf-upload") || "null")
        : null;
      if (resume && (!saved || saved.kind !== kind))
        throw new Error(
          "Select the same source type and files as the pending upload.",
        );
      const identity = await uploadAsset(
        kind,
        files,
        name || files[0]?.name || "Evidence",
        (done, total, stage) =>
          setProgress(
            stage ||
              `${Math.round((done / Math.max(total, 1)) * 100)}% uploaded`,
          ),
        saved?.id,
      );
      onAdd({ kind, asset_id: identity }, name || files[0]?.name || "Evidence");
      setFiles([]);
      setProgress("Source ready");
      await assets.refetch();
    } catch (error) {
      setError(error);
    } finally {
      setBusy(false);
    }
  }
  return (
    <section className="panel">
      <h2>Add evidence</h2>
      <div className="form-grid">
        <label>
          Source type
          <select
            aria-label="Source type"
            value={kind}
            onChange={(event) => {
              setKind(event.target.value as Source["kind"]);
              setFiles([]);
            }}
          >
            {Object.entries(kinds)
              .filter(([key]) => !allowed || allowed.includes(key))
              .map(([key, title]) => (
                <option key={key} value={key}>
                  {title}
                </option>
              ))}
          </select>
        </label>
        <label>
          Location
          <select
            aria-label="Location"
            value={mode}
            onChange={(event) => setMode(event.target.value)}
          >
            <option value="upload">Upload from this computer</option>
            <option value="existing">Previously uploaded source</option>
            <option value="server">Configured server folder</option>
          </select>
        </label>
      </div>
      {mode === "upload" && (
        <>
          {kind === "auto" && (
            <label>
              Upload selection
              <select
                aria-label="Upload selection"
                value={selection}
                onChange={(event) => {
                  setSelection(event.target.value);
                  setFiles([]);
                }}
              >
                <option value="files">Files or ZIP archives</option>
                <option value="folder">A folder and its subfolders</option>
              </select>
            </label>
          )}
          <label>
            Source label
            <input
              value={name}
              onChange={(event) => setName(event.target.value)}
              placeholder="e.g. Bills 2023–2025"
            />
          </label>
          <label>
            {kind === "auto"
              ? selection === "folder"
                ? "Choose a folder"
                : "Choose files or ZIPs"
              : kind.endsWith("_dir")
                ? "Choose a folder"
                : "Choose a file"}
            <input
              key={kind}
              type="file"
              multiple={kind === "auto" || kind.endsWith("_dir")}
              {...((kind === "auto" && selection === "folder") ||
              kind.endsWith("_dir")
                ? { webkitdirectory: "" }
                : {})}
              onChange={(event) =>
                setFiles(Array.from(event.target.files || []))
              }
            />
          </label>
          <p className="muted">
            {files.length} files selected. Large files upload in resumable
            chunks. Keep the source files unchanged when resuming.
          </p>
          <div className="actions">
            <button
              disabled={busy || !files.length}
              onClick={() => void send()}
            >
              Upload and add
            </button>
            {sessionStorage.getItem("edf-upload") && (
              <button
                disabled={busy || !files.length}
                onClick={() => void send(true)}
              >
                Resume pending upload
              </button>
            )}
            <span role="status">{progress}</span>
          </div>
        </>
      )}
      {mode === "existing" && (
        <>
          <label>
            Uploaded source
            <select
              aria-label="Uploaded source"
              value={existing}
              onChange={(event) => setExisting(event.target.value)}
            >
              <option value="">Choose…</option>
              {assets.data
                ?.filter(
                  (asset) => asset.data.sealed && asset.data.kind === kind,
                )
                .map((asset) => (
                  <option key={asset.id} value={asset.id}>
                    {asset.data.name}
                  </option>
                ))}
            </select>
          </label>
          <button
            disabled={!existing}
            onClick={() =>
              onAdd(
                { kind, asset_id: existing },
                assets.data?.find((asset) => asset.id === existing)?.data
                  .name || "Source",
              )
            }
          >
            Add source
          </button>
        </>
      )}
      {mode === "server" && (
        <>
          <label>
            Server source folder
            <select
              aria-label="Server source folder"
              value={root}
              onChange={(event) => {
                setRoot(event.target.value);
                setRelative("");
              }}
            >
              <option value="">Choose…</option>
              {roots.data?.map((item) => (
                <option key={item.id} value={item.id}>
                  {item.name}
                </option>
              ))}
            </select>
          </label>
          <label>
            Relative folder or filename
            <input
              value={relative}
              onChange={(event) => setRelative(event.target.value)}
            />
          </label>
          <div className="browse">
            {browser.data?.map((item) => (
              <button
                key={item.relative}
                onClick={() => setRelative(item.relative)}
              >
                {item.directory ? "📁" : "📄"} {item.name}
              </button>
            ))}
          </div>
          <ErrorMessage error={browser.error} />
          <button
            disabled={!root || !relative}
            onClick={() => onAdd({ kind, root_id: root, relative }, relative)}
          >
            Add selected path
          </button>
          <p className="muted">
            Server folders are configured at launch. Only those folders are
            browsable.
          </p>
        </>
      )}
      <ErrorMessage error={error} />
    </section>
  );
}

function Settings({
  config,
  setConfig,
}: {
  config: Record<string, Json>;
  setConfig: (value: Record<string, Json>) => void;
}) {
  const schema = useSchema();
  return (
    <details className="panel">
      <summary>
        <Settings2 size={16} /> Analysis settings
      </summary>
      <p>
        Blank settings use the shared application defaults. A saved analysis
        keeps its pinned settings unless you explicitly change them.
      </p>
      <div className="form-grid">
        {Object.entries(schema.data?.settings || {})
          .filter(([key]) => key !== "report_sections")
          .map(([key, type]) => (
            <label key={key}>
              {label(key)}
              {type.includes("bool") ? (
                <select
                  value={config[key] === undefined ? "" : String(config[key])}
                  onChange={(event) => {
                    const next = { ...config };
                    if (event.target.value === "") delete next[key];
                    else next[key] = event.target.value === "true";
                    setConfig(next);
                  }}
                >
                  <option value="">Application default</option>
                  <option value="true">Yes</option>
                  <option value="false">No</option>
                </select>
              ) : (
                <input
                  type={
                    key === "as_of"
                      ? "date"
                      : type.includes("float") || type.includes("int")
                        ? "number"
                        : "text"
                  }
                  step={type.includes("float") ? "any" : undefined}
                  min={
                    type.includes("float") || type.includes("int")
                      ? 0
                      : undefined
                  }
                  value={String(config[key] ?? "")}
                  onChange={(event) => {
                    const next = { ...config };
                    if (event.target.value === "") delete next[key];
                    else
                      next[key] =
                        type.includes("float") || type.includes("int")
                          ? Number(event.target.value)
                          : event.target.value;
                    setConfig(next);
                  }}
                />
              )}
            </label>
          ))}
      </div>
    </details>
  );
}

function ReportOptions({
  formats,
  setFormats,
  sections,
  setSections,
}: {
  formats: string[];
  setFormats: (value: string[]) => void;
  sections: string[] | null;
  setSections: (value: string[] | null) => void;
}) {
  const schema = useSchema();
  return (
    <details className="panel">
      <summary>Report formats & sections</summary>
      <p>Every published pack includes an editable Excel workbook.</p>
      <div className="checks">
        {["pdf", "docx", "html"].map((format) => (
          <label key={format}>
            <input
              type="checkbox"
              checked={formats.includes(format)}
              onChange={(event) =>
                setFormats(
                  event.target.checked
                    ? [...formats, format]
                    : formats.filter((item) => item !== format),
                )
              }
            />
            {format.toUpperCase()}
          </label>
        ))}
      </div>
      <label className="check">
        <input
          type="checkbox"
          checked={sections === null}
          onChange={(event) =>
            setSections(
              event.target.checked
                ? null
                : schema.data?.sections.map((section) => section.key) || [],
            )
          }
        />
        Include every report section
      </label>
      {sections !== null && (
        <div className="checks">
          {schema.data?.sections.map((section) => (
            <label key={section.key}>
              <input
                type="checkbox"
                checked={sections.includes(section.key)}
                onChange={(event) =>
                  setSections(
                    event.target.checked
                      ? [...sections, section.key]
                      : sections.filter((key) => key !== section.key),
                  )
                }
              />
              {section.title}
            </label>
          ))}
        </div>
      )}
    </details>
  );
}

function Build() {
  const navigate = useNavigate();
  const [mode, setMode] = useState("extract"),
    [name, setName] = useState("My evidence pack"),
    [sources, setSources] = useState<{ source: Source; name: string }[]>([]),
    [config, setConfig] = useState<Record<string, Json>>({}),
    [formats, setFormats] = useState(["xlsx", "pdf", "docx", "html"]),
    [sections, setSections] = useState<string[] | null>(null),
    [error, setError] = useState<unknown>(),
    [busy, setBusy] = useState(false);
  const [inspection, setInspection] = useState(""),
    [plan, setPlan] = useState<IntakeSummary | null>(null),
    [acceptSkipped, setAcceptSkipped] = useState(false);
  const automatic =
    mode === "extract" && sources.some((item) => item.source.kind === "auto");
  useEffect(() => {
    setInspection("");
    setPlan(null);
    setAcceptSkipped(false);
  }, [sources]);
  const inspectSources = async () => {
    setBusy(true);
    try {
      const value = await api<Entry<JobData>>("/jobs", {
        mode: "intake",
        name: "Inspect " + name,
        sources: sources.map((item) => item.source),
      });
      setInspection(value.id);
    } catch (error) {
      setError(error);
    } finally {
      setBusy(false);
    }
  };
  const submit = async () => {
    setBusy(true);
    try {
      const job = await api<Entry<JobData>>("/jobs", {
        name,
        mode,
        sources:
          mode === "extract"
            ? automatic && inspection
              ? [{ kind: "auto", intake_id: inspection, relative: "__all__" }]
              : sources.map((item) => item.source)
            : [],
        input: mode === "extract" ? null : sources[0]?.source,
        config,
        formats,
        sections,
      });
      navigate("/runs/" + job.id);
    } catch (error) {
      setError(error);
    } finally {
      setBusy(false);
    }
  };
  return (
    <>
      <header className="page-heading">
        <div className="eyebrow">Collect → check → correct → regenerate</div>
        <h1>Build an evidence pack</h1>
        <p>
          Preserve original files and produce the same accepted analysis as
          desktop and CLI.
        </p>
      </header>
      <section className="panel">
        <div className="form-grid">
          <label>
            Pack name
            <input
              value={name}
              onChange={(event) => setName(event.target.value)}
            />
          </label>
          <label>
            Workflow
            <select
              aria-label="Workflow"
              value={mode}
              onChange={(event) => {
                setMode(event.target.value);
                setSources([]);
              }}
            >
              <option value="extract">Extract original evidence</option>
              <option value="render">
                Import saved analysis / evidence pack
              </option>
              <option value="workbook">
                Regenerate amended Excel workbook
              </option>
            </select>
          </label>
        </div>
      </section>
      <Sources
        key={mode}
        allowed={
          mode === "extract"
            ? ["auto", "pdf_dir", "htm", "pst", "ost", "eml_dir", "msg_dir"]
            : mode === "workbook"
              ? ["workbook"]
              : ["bundle", "result"]
        }
        onAdd={(source, name) =>
          setSources((items) =>
            mode === "extract"
              ? [...items, { source, name }]
              : [{ source, name }],
          )
        }
      />
      <section className="panel">
        <h2>Selected sources</h2>
        {!sources.length && <p>No sources added yet.</p>}
        {sources.map((item, index) => (
          <div className="source-row" key={index}>
            <span>
              {kinds[item.source.kind]} · {item.name}
            </span>
            <button
              onClick={() =>
                setSources((items) => items.filter((_, n) => n !== index))
              }
            >
              Remove
            </button>
          </div>
        ))}
      </section>
      {automatic && (
        <>
          <button
            disabled={busy || !sources.length}
            onClick={() => void inspectSources()}
          >
            Inspect everything
          </button>
          {inspection && <IntakePanel id={inspection} onSummary={setPlan} />}{" "}
          {plan &&
          (plan.counts.unsupported ||
            plan.counts.needs_choice ||
            plan.counts.failed ||
            plan.counts.skipped) ? (
            <label className="check">
              <input
                type="checkbox"
                checked={acceptSkipped}
                onChange={(event) => setAcceptSkipped(event.target.checked)}
              />
              I have checked the unprocessed files. They remain in the intake
              inventory and are not automatically included as analysed evidence.
            </label>
          ) : null}
        </>
      )}
      <Settings config={config} setConfig={setConfig} />
      <ReportOptions
        formats={formats}
        setFormats={setFormats}
        sections={sections}
        setSections={setSections}
      />
      <ErrorMessage error={error} />
      <button
        className="primary"
        disabled={
          busy ||
          !sources.length ||
          !name.trim() ||
          (automatic &&
            (!plan?.counts.ready ||
              (!!(
                plan.counts.unsupported ||
                plan.counts.needs_choice ||
                plan.counts.failed ||
                plan.counts.skipped
              ) &&
                !acceptSkipped)))
        }
        onClick={() => void submit()}
      >
        {busy ? "Starting…" : "Create evidence pack"}
      </button>
    </>
  );
}

type SourceEntry = {
  category: string;
  row_id: string;
  fields: Record<string, Json>;
  relation_note: string;
  sources: {
    source_id: string;
    media_type: string;
    title: string;
    status: string;
    available: boolean;
    pages: [number, number] | null;
  }[];
};
function Run() {
  const id = useParams().id!;
  const [error, setError] = useState<unknown>(),
    [messages, setMessages] = useState<string[]>([]),
    [tab, setTab] = useState("evidence"),
    [pending, setPending] = useState(false);
  const job = useQuery({
    queryKey: ["job", id],
    queryFn: () => api<Entry<JobData>>("/jobs/" + id),
    refetchInterval: (query) =>
      terminal.has(query.state.data?.data.status || "") ? false : 2000,
  });
  const artifacts = useQuery({
    queryKey: ["artifacts", id, job.data?.data.status],
    queryFn: () =>
      api<{ path: string; size: number }[]>(`/jobs/${id}/artifacts`),
  });
  const hasResult = artifacts.data?.some((item) => item.path === "result.json");
  useEffect(() => {
    setMessages([]);
    const events = new EventSource(`/api/jobs/${id}/events`);
    events.onmessage = (event) => {
      const value = JSON.parse(event.data);
      setMessages((items) => [...items.slice(-499), value.message]);
    };
    events.addEventListener("finished", () => {
      events.close();
      void job.refetch();
    });
    return () => events.close();
  }, [id]);
  return (
    <>
      <header className="page-heading">
        <Link to="/">← All runs</Link>
        <h1>{job.data?.data.request.name || "Opening run…"}</h1>
        <span className={"badge " + job.data?.data.status}>
          {label(job.data?.data.status || "loading")}
        </span>
        <p>
          {job.data?.data.request.mode === "preview"
            ? "Preview only: this analysis has not published a reviewed evidence pack."
            : "Original inputs and earlier runs remain available."}
        </p>
      </header>
      <ErrorMessage error={job.error || error || job.data?.data.error} />
      {job.data && !terminal.has(job.data.data.status) && (
        <button
          onClick={async () => {
            try {
              await api(`/jobs/${id}/cancel`, {});
              await job.refetch();
            } catch (error) {
              setError(error);
            }
          }}
        >
          Cancel run
        </button>
      )}
      <Tabs.Root
        value={tab}
        onValueChange={(value) => {
          if (
            pending &&
            !window.confirm("Discard unapplied form edits and change screens?")
          )
            return;
          setPending(false);
          setTab(value);
        }}
      >
        <Tabs.List className="tabs" aria-label="Run workspace">
          <Tabs.Trigger value="evidence">Evidence & review</Tabs.Trigger>
          <Tabs.Trigger value="outputs">Outputs</Tabs.Trigger>
          <Tabs.Trigger value="activity">Activity</Tabs.Trigger>
        </Tabs.List>
        <Tabs.Content value="evidence">
          {job.data?.data.request.mode === "intake" ? (
            <IntakePanel id={id} />
          ) : hasResult ? (
            <Workspace key={id} id={id} onPending={setPending} />
          ) : (
            <div className="empty">
              <h2>
                {job.data?.data.status === "empty"
                  ? "No billing evidence found"
                  : "Analysis will appear here"}
              </h2>
              <p>
                Check activity for warnings and inspect available outputs if the
                run failed or was interrupted.
              </p>
            </div>
          )}
        </Tabs.Content>
        <Tabs.Content value="outputs">
          <section className="panel">
            <h2>Download outputs</h2>
            {job.data && terminal.has(job.data.data.status) && (
              <a className="button primary" href={`/api/jobs/${id}/pack`}>
                Download complete evidence pack
              </a>
            )}
            <p>Keep the pack together so source links remain portable.</p>
            <ul className="downloads">
              {artifacts.data?.map((item) => (
                <li key={item.path}>
                  <a
                    href={`/api/jobs/${id}/download?${new URLSearchParams({ path: item.path })}`}
                  >
                    {item.path}
                  </a>
                  <span>{Math.ceil(item.size / 1024).toLocaleString()} KB</span>
                </li>
              ))}
            </ul>
          </section>
        </Tabs.Content>
        <Tabs.Content value="activity">
          <section className="panel">
            <h2>Run activity</h2>
            <div className="log" role="log" aria-live="polite">
              {messages.map((message, index) => (
                <p key={index}>{message}</p>
              ))}
            </div>
          </section>
        </Tabs.Content>
      </Tabs.Root>
    </>
  );
}

function Workspace({
  id,
  onPending,
}: {
  id: string;
  onPending: (value: boolean) => void;
}) {
  const client = useQueryClient(),
    navigate = useNavigate();
  const review = useQuery({
    queryKey: ["review", id],
    queryFn: () => api<Review>(`/jobs/${id}/review`),
  });
  const sources = useQuery({
    queryKey: ["sources", id],
    queryFn: () => api<SourceEntry[]>(`/jobs/${id}/sources`),
  });
  const [selected, setSelected] = useState(""),
    [search, setSearch] = useState(""),
    [sourceId, setSourceId] = useState(""),
    [error, setError] = useState<unknown>(),
    [busy, setBusy] = useState(false),
    [dirty, setDirty] = useState(false),
    [filter, setFilter] = useState("all");
  const [formats, setFormats] = useState(["xlsx", "pdf", "docx", "html"]),
    [sections, setSections] = useState<string[] | null>(null);
  const [tool, setTool] = useState("records");
  usePending(dirty);
  useEffect(() => {
    onPending(dirty);
  }, [dirty, onPending]);
  const data = review.data;
  async function command(operation: string, arguments_: object) {
    if (!data) return;
    setBusy(true);
    setError(undefined);
    try {
      const value = await api<Review>(`/jobs/${id}/review`, {
        revision: data.revision,
        operation,
        arguments: arguments_,
      });
      client.setQueryData(["review", id], value);
    } catch (error) {
      setError(error);
      throw error;
    } finally {
      setBusy(false);
    }
  }
  async function publish(mode: "preview" | "review" | "render") {
    if (
      dirty &&
      !window.confirm(
        "Unapplied form edits are not included. Continue using your last saved draft?",
      )
    )
      return;
    setBusy(true);
    try {
      const job = await api<Entry<JobData>>("/jobs", {
        mode,
        parent_id: id,
        draft: data?.draft,
        name:
          mode === "review"
            ? "Reviewed evidence pack"
            : mode === "preview"
              ? "Review impact preview"
              : "Rendered evidence pack",
        formats,
        sections,
      });
      navigate("/runs/" + job.id);
    } catch (error) {
      setError(error);
    } finally {
      setBusy(false);
    }
  }
  const row = data?.rows.find((row) => row.id === selected);
  const entry = sources.data?.find((entry) => entry.row_id === selected);
  const available = entry?.sources || [];
  const source =
    available.find((source) => source.source_id === sourceId) || available[0];
  const filtered =
    data?.rows.filter(
      (row) =>
        filter === "all" ||
        (filter === "included"
          ? row.decision.include
          : filter === "excluded"
            ? !row.decision.include
            : row.status === filter),
    ) || [];
  const select = (value: string) => {
    if (
      dirty &&
      !window.confirm("Discard unapplied form edits and select another record?")
    )
      return;
    setSelected(value);
    setSourceId("");
    setDirty(false);
  };
  if (!data)
    return (
      <>
        <ErrorMessage error={review.error} />
        <p role="status">Opening accepted analysis…</p>
      </>
    );
  return (
    <>
      <div className="summary-cards">
        {Object.entries(data.summary).map(([key, value]) => (
          <div key={key}>
            <span>{label(key)}</span>
            <strong>{value === null ? "Unknown" : "£" + value}</strong>
          </div>
        ))}
      </div>
      <div className="toolbar">
        <button
          disabled={!data.can_undo || busy || dirty}
          onClick={() => void command("undo", {}).catch(() => {})}
        >
          Undo
        </button>
        <button
          disabled={!data.can_redo || busy || dirty}
          onClick={() => void command("redo", {}).catch(() => {})}
        >
          Redo
        </button>
        <button disabled={busy} onClick={() => void publish("preview")}>
          Preview changes
        </button>
        <button
          className="primary"
          disabled={busy}
          onClick={() => void publish("review")}
        >
          Create reviewed pack
        </button>
        <button disabled={busy} onClick={() => void publish("render")}>
          Render accepted analysis
        </button>
        <button disabled={dirty || busy} onClick={() => void review.refetch()}>
          Reload saved draft
        </button>
        {dirty && <span className="badge">Unapplied form edits</span>}
      </div>
      <ErrorMessage error={error} />
      <p className="muted">
        Accepted totals above remain unchanged until you preview or publish your
        corrections. Draft revision {data.revision} is saved on the server.
      </p>
      <Tabs.Root
        value={tool}
        onValueChange={(value) => {
          if (
            dirty &&
            !window.confirm("Discard unapplied form edits and change tools?")
          )
            return;
          setDirty(false);
          setTool(value);
        }}
      >
        <Tabs.List className="tabs" aria-label="Evidence tools">
          <Tabs.Trigger value="records">Records</Tabs.Trigger>
          <Tabs.Trigger value="revisions">Bill revisions</Tabs.Trigger>
          <Tabs.Trigger value="ledger">Dispute ledger</Tabs.Trigger>
          <Tabs.Trigger value="claims">Claims & complaint</Tabs.Trigger>
          <Tabs.Trigger value="analyses">Analysis</Tabs.Trigger>
          <Tabs.Trigger value="sources">Source files</Tabs.Trigger>
          <Tabs.Trigger value="reports">Report settings</Tabs.Trigger>
        </Tabs.List>
        <Tabs.Content value="records">
          <section className="panel">
            <div className="toolbar">
              <input
                aria-label="Search evidence"
                placeholder="Search invoice, date, amount, filename…"
                value={search}
                onChange={(event) => setSearch(event.target.value)}
              />
              <select
                aria-label="Record filter"
                value={filter}
                onChange={(event) => setFilter(event.target.value)}
              >
                {[
                  "all",
                  "included",
                  "excluded",
                  "verified",
                  "needs_information",
                  "unreviewed",
                ].map((value) => (
                  <option key={value} value={value}>
                    {label(value)}
                  </option>
                ))}
              </select>
            </div>
            <EvidenceTable
              rows={filtered}
              selected={selected}
              onSelect={select}
              search={search}
            />
          </section>
          {row ? (
            <div className="review-split">
              <section className="panel">
                <h2>{String(row.record["Invoice #"] || "Selected record")}</h2>
                <RecordEditor
                  key={selected + ":" + data.revision}
                  row={row}
                  busy={busy}
                  onDirty={setDirty}
                  onSave={(values) =>
                    command("update", { row_id: selected, ...values })
                  }
                />
                <details>
                  <summary>All record fields & extracted text</summary>
                  <Value value={row.record} />
                </details>
              </section>
              <section className="panel original">
                <h2>Original evidence</h2>
                {available.length > 0 ? (
                  <>
                    <label>
                      Source
                      <select
                        aria-label="Source"
                        value={source?.source_id || ""}
                        onChange={(event) => setSourceId(event.target.value)}
                      >
                        {available.map((source, index) => (
                          <option
                            key={source.source_id + index}
                            value={source.source_id}
                          >
                            {source.title} · {source.status}
                          </option>
                        ))}
                      </select>
                    </label>
                    <p>{source?.status}</p>
                    {source?.available && (
                      <>
                        <a
                          href={`/api/jobs/${id}/source?${new URLSearchParams({ source_id: source.source_id })}`}
                          target="_blank"
                          rel="noreferrer"
                        >
                          Open original file
                        </a>
                        {(source.media_type === "application/pdf" ||
                          source.title.toLowerCase().endsWith(".pdf")) && (
                          <Suspense fallback={<p>Loading PDF viewer…</p>}>
                            <Pdf
                              url={`/api/jobs/${id}/source?${new URLSearchParams({ source_id: source.source_id })}`}
                              initialPage={source.pages?.[0] || 1}
                            />
                          </Suspense>
                        )}
                      </>
                    )}
                  </>
                ) : (
                  <p>No verified source is available for this record.</p>
                )}
              </section>
            </div>
          ) : (
            <p className="empty">
              Select a record to review its fields beside the original.
            </p>
          )}
        </Tabs.Content>
        <Tabs.Content value="revisions">
          <RevisionEditor data={data} command={command} onDirty={setDirty} />
          <section className="panel">
            <h2>Revision assessment</h2>
            <Value value={data.revisions} />
          </section>
        </Tabs.Content>
        <Tabs.Content value="ledger">
          <PositionEditor data={data} command={command} onDirty={setDirty} />
          <section className="panel">
            <h2>Billing histories & comparisons</h2>
            <Value value={data.ledger} />
          </section>
        </Tabs.Content>
        <Tabs.Content value="claims">
          <ClaimEditor data={data} command={command} onDirty={setDirty} />
          <section className="panel">
            <h2>Reconciled proposals</h2>
            <Value value={data.claims} />
          </section>
        </Tabs.Content>
        <Tabs.Content value="analyses">
          <Analysis data={data} />
        </Tabs.Content>
        <Tabs.Content value="sources">
          <section className="panel">
            <h2>Source files & findings</h2>
            <ErrorMessage error={sources.error} />
            {sources.data?.map((entry, index) => (
              <details key={index}>
                <summary>
                  {entry.category} ·{" "}
                  {String(
                    entry.fields["Attachment Name"] ||
                      entry.fields["Invoice #"] ||
                      "Evidence",
                  )}
                </summary>
                <p>{entry.relation_note}</p>
                <Value value={entry.fields} />
                {entry.sources.map((source, index) => (
                  <p key={index}>
                    {source.available ? (
                      <a
                        href={`/api/jobs/${id}/source?${new URLSearchParams({ source_id: source.source_id })}`}
                        target="_blank"
                        rel="noreferrer"
                      >
                        {source.title}
                      </a>
                    ) : (
                      source.title
                    )}{" "}
                    · {source.status}
                  </p>
                ))}
              </details>
            ))}
          </section>
        </Tabs.Content>
        <Tabs.Content value="reports">
          <ReportOptions
            formats={formats}
            setFormats={setFormats}
            sections={sections}
            setSections={setSections}
          />
        </Tabs.Content>
      </Tabs.Root>
    </>
  );
}

type Command = (operation: string, args: object) => Promise<void>;
function RecordEditor({
  row,
  busy,
  onDirty,
  onSave,
}: {
  row: Row;
  busy: boolean;
  onDirty: (dirty: boolean) => void;
  onSave: (values: object) => Promise<void>;
}) {
  const schema = useSchema();
  const [changes, setChanges] = useState<Record<string, Json>>({}),
    [include, setInclude] = useState(row.decision.include),
    [reason, setReason] = useState(row.decision.reason),
    [status, setStatus] = useState(row.status),
    [error, setError] = useState<unknown>();
  const change = (key: string, value: Json) => {
    setChanges((values) => ({ ...values, [key]: value }));
    onDirty(true);
  };
  return (
    <form
      onSubmit={async (event) => {
        event.preventDefault();
        try {
          await onSave({ changes, include, reason, status });
          setChanges({});
          onDirty(false);
        } catch (error) {
          setError(error);
        }
      }}
    >
      <label className="check">
        <input
          type="checkbox"
          checked={include}
          onChange={(event) => {
            setInclude(event.target.checked);
            onDirty(true);
          }}
        />
        Include this record
      </label>
      <label>
        Review status
        <select
          aria-label="Review status"
          value={status}
          onChange={(event) => {
            setStatus(event.target.value);
            onDirty(true);
          }}
        >
          <option value="unreviewed">Unreviewed</option>
          <option value="verified">Verified against source</option>
          <option value="needs_information">Needs information</option>
        </select>
      </label>
      <div className="form-grid">
        {Object.entries(schema.data?.review_fields || {}).map(([key, kind]) => (
          <label key={key}>
            {key}
            {kind === "bool" ? (
              <select
                value={String(changes[key] ?? row.record[key] ?? false)}
                onChange={(event) => change(key, event.target.value === "true")}
              >
                <option value="false">No</option>
                <option value="true">Yes</option>
              </select>
            ) : (
              <input
                value={String(
                  key in changes
                    ? (changes[key] ?? "")
                    : (row.record[key] ?? ""),
                )}
                placeholder={kind === "date" ? "DD/MM/YYYY" : "Unknown"}
                onChange={(event) => change(key, event.target.value || null)}
              />
            )}
          </label>
        ))}
      </div>
      <label>
        Reason for change
        <textarea
          aria-label="Reason for change"
          value={reason}
          onChange={(event) => {
            setReason(event.target.value);
            onDirty(true);
          }}
        />
      </label>
      <ErrorMessage error={error} />
      <button disabled={busy} className="primary">
        Apply to draft
      </button>
      <button
        type="button"
        onClick={() => {
          setChanges({});
          setInclude(row.decision.include);
          setReason(row.decision.reason);
          setStatus(row.status);
          onDirty(false);
        }}
      >
        Discard form edits
      </button>
    </form>
  );
}

const bills = (data: Review) =>
  data.rows.filter((row) =>
    ["New Bill", "Ongoing Balance", "Charge", "Late Payment"].includes(
      String(row.record["Entry Type"]),
    ),
  );
const rowName = (row: Row) =>
  `${row.record["Invoice #"] || row.record["Entry Type"] || "Record"} · ${row.record["Date"] || "Undated"} · ${row.record["Account"] || "Unknown account"}`;
function RowSelect({
  title,
  value,
  setValue,
  rows,
}: {
  title: string;
  value: string;
  setValue: (value: string) => void;
  rows: Row[];
}) {
  return (
    <label>
      {title}
      <select value={value} onChange={(event) => setValue(event.target.value)}>
        <option value="">Choose…</option>
        {rows.map((row) => (
          <option key={row.id} value={row.id}>
            {rowName(row)}
          </option>
        ))}
      </select>
    </label>
  );
}

function RevisionEditor({
  data,
  command,
  onDirty,
}: {
  data: Review;
  command: Command;
  onDirty: (value: boolean) => void;
}) {
  const [older, setOlder] = useState(""),
    [newer, setNewer] = useState(""),
    [kind, setKind] = useState("replacement"),
    [status, setStatus] = useState("confirmed"),
    [reason, setReason] = useState(""),
    [error, setError] = useState<unknown>();
  return (
    <form
      className="panel"
      onChange={() => onDirty(true)}
      onSubmit={async (event) => {
        event.preventDefault();
        try {
          await command("set_revision", { older, newer, kind, status, reason });
          onDirty(false);
        } catch (error) {
          setError(error);
        }
      }}
    >
      <h2>Record a document relationship</h2>
      <p>
        Relationship decisions do not automatically exclude transactions. Verify
        the original documents.
      </p>
      <div className="form-grid">
        <RowSelect
          title="Earlier document"
          value={older}
          setValue={setOlder}
          rows={data.rows}
        />
        <RowSelect
          title="Later document"
          value={newer}
          setValue={setNewer}
          rows={data.rows}
        />
        <label>
          Relationship
          <select
            aria-label="Relationship"
            value={kind}
            onChange={(event) => setKind(event.target.value)}
          >
            {["replacement", "cancellation_credit"].map((value) => (
              <option key={value}>{value}</option>
            ))}
          </select>
        </label>
        <label>
          Decision
          <select
            aria-label="Decision"
            value={status}
            onChange={(event) => setStatus(event.target.value)}
          >
            {["confirmed", "rejected", "suggested"].map((value) => (
              <option key={value}>{value}</option>
            ))}
          </select>
        </label>
      </div>
      <label>
        Reason
        <textarea
          value={reason}
          onChange={(event) => setReason(event.target.value)}
        />
      </label>
      <ErrorMessage error={error} />
      <button>Save relationship</button>
    </form>
  );
}

function PositionEditor({
  data,
  command,
  onDirty,
}: {
  data: Review;
  command: Command;
  onDirty: (value: boolean) => void;
}) {
  const [selected, setSelected] = useState(""),
    [fields, setFields] = useState<Record<string, string>>({
      status: "unreviewed",
      challenged_amount: "",
      proposed_reduction: "",
      first_demand_date: "",
      basis: "",
      edf_request: "",
    }),
    [error, setError] = useState<unknown>();
  const changed = useRef(false);
  const reset = (value: string) => {
    changed.current = false;
    onDirty(false);
    setSelected(value);
    const schedule =
      (data.draft.schedule as unknown as Record<string, string>[]) || [];
    const position = schedule.find((item) => item.row_id === value);
    setFields({
      status: position?.status || "unreviewed",
      challenged_amount: position?.challenged_amount || "",
      proposed_reduction: position?.proposed_reduction || "",
      first_demand_date: position?.first_demand_date || "",
      basis: position?.basis || "",
      edf_request: position?.edf_request || "",
    });
  };
  const choose = (value: string) => {
    if (
      changed.current &&
      !window.confirm("Discard unapplied position edits and change bills?")
    )
      return;
    reset(value);
  };
  return (
    <form
      className="panel"
      onChange={(event) => {
        if (
          (event.target as HTMLElement).getAttribute("aria-label") !== "Bill"
        ) {
          changed.current = true;
          onDirty(true);
        }
      }}
      onSubmit={async (event) => {
        event.preventDefault();
        try {
          await command("set_position", {
            row_id: selected,
            ...fields,
            challenged_amount: fields.challenged_amount || null,
            proposed_reduction: fields.proposed_reduction || null,
            first_demand_date: fields.first_demand_date || null,
          });
          changed.current = false;
          onDirty(false);
        } catch (error) {
          setError(error);
        }
      }}
    >
      <h2>Dispute position</h2>
      <RowSelect
        title="Bill"
        value={selected}
        setValue={choose}
        rows={bills(data)}
      />
      <div className="form-grid">
        {Object.entries(fields).map(([key, value]) => (
          <label key={key}>
            {label(key)}
            {key === "status" ? (
              <select
                value={value}
                onChange={(event) =>
                  setFields((items) => ({
                    ...items,
                    [key]: event.target.value,
                  }))
                }
              >
                {[
                  "unreviewed",
                  "challenged",
                  "not_challenged",
                  "withdrawn",
                ].map((value) => (
                  <option key={value}>{value}</option>
                ))}
              </select>
            ) : key === "basis" || key === "edf_request" ? (
              <textarea
                aria-label={label(key)}
                value={value}
                onChange={(event) =>
                  setFields((items) => ({
                    ...items,
                    [key]: event.target.value,
                  }))
                }
              />
            ) : (
              <input
                value={value}
                placeholder={
                  key === "first_demand_date" ? "YYYY-MM-DD" : "Unknown"
                }
                onChange={(event) =>
                  setFields((items) => ({
                    ...items,
                    [key]: event.target.value,
                  }))
                }
              />
            )}
          </label>
        ))}
      </div>
      <ErrorMessage error={error} />
      <button disabled={!selected}>Save position</button>
      <button
        type="button"
        disabled={!selected}
        onClick={async () => {
          try {
            await command("clear_position", { row_id: selected });
            changed.current = false;
            onDirty(false);
            setFields({
              status: "unreviewed",
              challenged_amount: "",
              proposed_reduction: "",
              first_demand_date: "",
              basis: "",
              edf_request: "",
            });
          } catch (error) {
            setError(error);
          }
        }}
      >
        Clear position
      </button>
    </form>
  );
}

function ClaimEditor({
  data,
  command,
  onDirty,
}: {
  data: Review;
  command: Command;
  onDirty: (value: boolean) => void;
}) {
  const empty: Claim = {
    id: "",
    title: "",
    status: "draft",
    members: [],
    challenged: null,
    reduction: null,
    credits: [],
    reason: "",
    account_link_reason: "",
  };
  const [claim, setClaim] = useState<Claim>(empty),
    [context, setContext] = useState<Context>(data.draft.claims.context),
    [left, setLeft] = useState(""),
    [right, setRight] = useState(""),
    [reason, setReason] = useState(""),
    [error, setError] = useState<unknown>();
  const pending = useRef(new Set<string>());
  const mark = (scope: string) => {
    pending.current.add(scope);
    onDirty(true);
  };
  const update = (key: keyof Claim, value: Claim[keyof Claim]) => {
    setClaim((item) => ({ ...item, [key]: value }));
    mark("group");
  };
  const save = async (operation: string, args: object) => {
    try {
      await command(operation, args);
      const scope =
        operation === "set_complaint_context"
          ? "context"
          : operation.includes("separation") || operation === "separate_claims"
            ? "overlap"
            : "group";
      pending.current.delete(scope);
      onDirty(pending.current.size > 0);
    } catch (error) {
      setError(error);
    }
  };
  const groups = data.draft.claims.groups;
  useEffect(() => {
    if (!pending.current.has("group"))
      setClaim(
        (current) =>
          data.draft.claims.groups.find((group) => group.id === current.id) ||
          empty,
      );
    if (!pending.current.has("context")) setContext(data.draft.claims.context);
  }, [data.revision]);
  const chooseGroup = (id: string) => {
    if (
      pending.current.has("group") &&
      !window.confirm("Discard unapplied claim edits and change groups?")
    )
      return;
    pending.current.delete("group");
    onDirty(pending.current.size > 0);
    setClaim(groups.find((group) => group.id === id) || empty);
  };
  return (
    <>
      <section className="panel">
        <h2>Consolidate your claims</h2>
        <p>
          Keep proposed reductions separate from observed account balances.
          Repeated invoices must not multiply a single challenge.
        </p>
        <label>
          Existing claim
          <select
            aria-label="Existing claim"
            value={claim.id}
            onChange={(event) => chooseGroup(event.target.value)}
          >
            <option value="">New claim…</option>
            {groups.map((group) => (
              <option key={group.id} value={group.id}>
                {group.title}
              </option>
            ))}
          </select>
        </label>
        <form
          onSubmit={(event) => {
            event.preventDefault();
            const {
              id,
              title,
              status,
              members,
              challenged,
              reduction,
              credits,
              reason,
              account_link_reason,
            } = claim;
            void save("set_claim", {
              claim_id: id,
              title,
              status,
              members,
              challenged,
              reduction,
              credits,
              reason,
              account_link_reason,
            });
          }}
        >
          <div className="form-grid">
            {(
              [
                "id",
                "title",
                "challenged",
                "reduction",
                "reason",
                "account_link_reason",
              ] as const
            ).map((key) => (
              <label key={key}>
                {label(key)}
                <input
                  value={claim[key] || ""}
                  onChange={(event) =>
                    update(
                      key,
                      event.target.value ||
                        (key === "challenged" || key === "reduction"
                          ? null
                          : ""),
                    )
                  }
                />
              </label>
            ))}
            <label>
              Status
              <select
                aria-label="Status"
                value={claim.status}
                onChange={(event) => update("status", event.target.value)}
              >
                {["draft", "reviewed", "withdrawn"].map((value) => (
                  <option key={value}>{value}</option>
                ))}
              </select>
            </label>
          </div>
          <fieldset>
            <legend>Bills included in this consolidated challenge</legend>
            {bills(data).map((row) => (
              <label className="check" key={row.id}>
                <input
                  type="checkbox"
                  checked={claim.members.includes(row.id)}
                  onChange={(event) =>
                    update(
                      "members",
                      event.target.checked
                        ? [...claim.members, row.id]
                        : claim.members.filter((id) => id !== row.id),
                    )
                  }
                />
                {rowName(row)}
              </label>
            ))}
          </fieldset>
          <h3>Credit allocations</h3>
          {claim.credits.map((credit, index) => (
            <div className="credit" key={index}>
              <RowSelect
                title="Credit source"
                value={credit.row_id}
                setValue={(value) =>
                  update(
                    "credits",
                    claim.credits.map((item, n) =>
                      n === index ? { ...item, row_id: value } : item,
                    ),
                  )
                }
                rows={data.rows.filter((row) =>
                  [
                    "Credit",
                    "New Bill",
                    "Ongoing Balance",
                    "Charge",
                    "Late Payment",
                  ].includes(String(row.record["Entry Type"])),
                )}
              />
              {(["amount", "reason"] as const).map((key) => (
                <label key={key}>
                  {label(key)}
                  <input
                    value={credit[key]}
                    onChange={(event) =>
                      update(
                        "credits",
                        claim.credits.map((item, n) =>
                          n === index
                            ? { ...item, [key]: event.target.value }
                            : item,
                        ),
                      )
                    }
                  />
                </label>
              ))}
              <label>
                Treatment
                <select
                  aria-label="Treatment"
                  value={credit.treatment}
                  onChange={(event) =>
                    update(
                      "credits",
                      claim.credits.map((item, n) =>
                        n === index
                          ? { ...item, treatment: event.target.value }
                          : item,
                      ),
                    )
                  }
                >
                  <option value="already_applied_relief">
                    Already applied relief
                  </option>
                  <option value="cancellation_offset">
                    Cancellation offset
                  </option>
                </select>
              </label>
              <button
                type="button"
                onClick={() =>
                  update(
                    "credits",
                    claim.credits.filter((_, n) => n !== index),
                  )
                }
              >
                Remove allocation
              </button>
            </div>
          ))}
          <button
            type="button"
            onClick={() =>
              update("credits", [
                ...claim.credits,
                {
                  row_id: "",
                  amount: "",
                  treatment: "already_applied_relief",
                  reason: "",
                } satisfies Credit,
              ])
            }
          >
            Add credit allocation
          </button>
          <div className="actions">
            <button className="primary">Save claim</button>
            <button
              type="button"
              disabled={!groups.some((group) => group.id === claim.id)}
              onClick={() => void save("clear_claim", { claim_id: claim.id })}
            >
              Remove claim
            </button>
          </div>
        </form>
      </section>
      <section className="panel">
        <h2>Resolve potential overlaps</h2>
        <div className="form-grid">
          {[
            { title: "First claim", value: left, set: setLeft },
            { title: "Second claim", value: right, set: setRight },
          ].map((item) => (
            <label key={item.title}>
              {item.title}
              <select
                value={item.value}
                onChange={(event) => item.set(event.target.value)}
              >
                <option value="">Choose…</option>
                {groups.map((group) => (
                  <option key={group.id} value={group.id}>
                    {group.title}
                  </option>
                ))}
              </select>
            </label>
          ))}
        </div>
        <label>
          Evidence that these are separate charges
          <textarea
            aria-label="Reason for change"
            value={reason}
            onChange={(event) => {
              setReason(event.target.value);
              mark("overlap");
            }}
          />
        </label>
        <button
          onClick={() => void save("separate_claims", { left, right, reason })}
        >
          Record separation
        </button>
        <button onClick={() => void save("clear_separation", { left, right })}>
          Clear separation
        </button>
      </section>
      <form
        className="panel"
        onSubmit={(event) => {
          event.preventDefault();
          void save("set_complaint_context", {
            summary: context.complaint_summary,
            outcome: context.requested_outcome,
            include_pack: context.include_pack,
          });
        }}
      >
        <h2>Complaint wording</h2>
        <label>
          Complaint summary
          <textarea
            aria-label="Complaint summary"
            value={context.complaint_summary}
            onChange={(event) => {
              setContext((value) => ({
                ...value,
                complaint_summary: event.target.value,
              }));
              mark("context");
            }}
          />
        </label>
        <label>
          Requested outcome
          <textarea
            aria-label="Requested outcome"
            value={context.requested_outcome}
            onChange={(event) => {
              setContext((value) => ({
                ...value,
                requested_outcome: event.target.value,
              }));
              mark("context");
            }}
          />
        </label>
        <label className="check">
          <input
            type="checkbox"
            checked={context.include_pack}
            onChange={(event) => {
              setContext((value) => ({
                ...value,
                include_pack: event.target.checked,
              }));
              mark("context");
            }}
          />
          Include focused complaint pack when publishing
        </label>
        <button>Save complaint wording</button>
      </form>
      <ErrorMessage error={error} />
    </>
  );
}

function Analysis({ data }: { data: Review }) {
  const [selected, setSelected] = useState(Object.keys(data.models)[0] || "");
  const ledgerKey = (row: Row) =>
    JSON.stringify(
      ["Account", "Fuel", "Meter", "Register"].map(
        (key) => row.record[key] || null,
      ),
    );
  const ledgers = Array.from(new Set(data.rows.map(ledgerKey)));
  const [ledger, setLedger] = useState(ledgers[0] || "");
  const chart = useRef<HTMLDivElement>(null);
  useEffect(() => {
    let disposed = false;
    let instance: import("echarts").ECharts | undefined;
    let observer: ResizeObserver | undefined;
    void import("echarts").then((echarts) => {
      if (disposed || !chart.current) return;
      instance = echarts.init(chart.current);
      const points = data.rows.filter(
        (row) =>
          ledgerKey(row) === ledger &&
          row.record["Balance (£)"] !== null &&
          row.record["Balance (£)"] !== undefined,
      );
      instance.setOption({
        aria: { enabled: true },
        tooltip: { trigger: "axis" },
        xAxis: {
          type: "category",
          data: points.map((row) => String(row.record.Date || "Undated")),
        },
        yAxis: { type: "value", name: "Observed balance (£)" },
        dataZoom: [{ type: "inside" }, { type: "slider" }],
        series: [
          {
            type: "scatter",
            data: points.map((row) => Number(row.record["Balance (£)"])),
          },
        ],
      });
      observer = new ResizeObserver(() => instance?.resize());
      observer.observe(chart.current);
    });
    return () => {
      disposed = true;
      observer?.disconnect();
      instance?.dispose();
    };
  }, [data, ledger]);
  return (
    <>
      <section className="panel">
        <h2>Observed balances</h2>
        <label>
          Ledger
          <select
            aria-label="Balance chart ledger"
            value={ledger}
            onChange={(event) => setLedger(event.target.value)}
          >
            {ledgers.map((key) => (
              <option key={key} value={key}>
                {(JSON.parse(key) as (string | null)[])
                  .map((value) => value || "Unknown")
                  .join(" · ")}
              </option>
            ))}
          </select>
        </label>
        <p>
          Each point is a source observation. Balances are not transaction
          totals. This chart shows only the selected account/fuel/meter/register
          scope. Unknown identity fields remain unknown; check the stored ledger
          models before assuming continuity.
        </p>
        <div
          ref={chart}
          className="chart"
          role="img"
          aria-label="Observed balances; full values are available in the evidence table"
        />
      </section>
      <section className="panel">
        <label>
          Stored analytical model
          <select
            aria-label="Stored analytical model"
            value={selected}
            onChange={(event) => setSelected(event.target.value)}
          >
            {Object.keys(data.models).map((key) => (
              <option key={key} value={key}>
                {label(key)}
              </option>
            ))}
          </select>
        </label>
        <p>
          These are the accepted calculations. Unavailable results remain
          explicit.
        </p>
        <Value value={data.models[selected]} />
      </section>
    </>
  );
}

function Compare() {
  const jobs = useJobs(),
    navigate = useNavigate();
  const [old, setOld] = useState(""),
    [newer, setNewer] = useState(""),
    [error, setError] = useState<unknown>();
  return (
    <>
      <h1>Compare runs</h1>
      <p>
        Compare prepared records using the same added, removed and changed
        record rules as the CLI.
      </p>
      <form
        className="panel"
        onSubmit={async (event) => {
          event.preventDefault();
          try {
            const job = await api<Entry<JobData>>("/jobs", {
              mode: "diff",
              name: "Run comparison",
              parent_id: old,
              other_id: newer,
            });
            navigate("/runs/" + job.id);
          } catch (error) {
            setError(error);
          }
        }}
      >
        {[
          { title: "Earlier run", value: old, set: setOld },
          { title: "Later run", value: newer, set: setNewer },
        ].map((item) => (
          <label key={item.title}>
            {item.title}
            <select
              value={item.value}
              onChange={(event) => item.set(event.target.value)}
            >
              <option value="">Choose…</option>
              {jobs.data
                ?.filter((job) => terminal.has(job.data.status))
                .map((job) => (
                  <option key={job.id} value={job.id}>
                    {job.data.request.name} · {job.id.slice(0, 8)}
                  </option>
                ))}
            </select>
          </label>
        ))}
        <ErrorMessage error={error} />
        <button disabled={!old || !newer || old === newer}>
          Compare records
        </button>
      </form>
    </>
  );
}

function Diagnostics() {
  const query = useQuery({
    queryKey: ["diagnostics"],
    queryFn: () => api<Json>("/diagnostics"),
  });
  return (
    <>
      <h1>Readiness & build</h1>
      <p>
        Capabilities refer to the server processing your evidence. Your browser
        does not need Python or OCR installed.
      </p>
      <button onClick={() => void query.refetch()}>Refresh checks</button>
      <ErrorMessage error={query.error} />
      <section className="panel">
        <Value value={query.data} />
      </section>
    </>
  );
}

function WorkspaceError() {
  return (
    <main className="signin">
      <h1>We couldn’t open this screen</h1>
      <p>
        Reload the workspace to try again. Saved runs and applied changes are
        kept.
      </p>
      <button onClick={() => window.location.reload()}>Reload workspace</button>
    </main>
  );
}

const router = createHashRouter([
  { path: "*", element: <App />, errorElement: <WorkspaceError /> },
]);
createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <QueryClientProvider client={queryClient}>
      <RouterProvider router={router} />
    </QueryClientProvider>
  </StrictMode>,
);
