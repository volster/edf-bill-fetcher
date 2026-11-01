import { useMemo, useRef, useState } from "react";
import {
  flexRender,
  getCoreRowModel,
  getFilteredRowModel,
  getSortedRowModel,
  useReactTable,
  type ColumnDef,
} from "@tanstack/react-table";
import { useVirtualizer } from "@tanstack/react-virtual";
import { label, type Json, type Row } from "./api";

function ExpandedValue({ value }: { value: unknown }) {
  const [open, setOpen] = useState(false);
  return (
    <details onToggle={(event) => setOpen(event.currentTarget.open)}>
      <summary>Show details</summary>
      {open && <Value value={value} />}
    </details>
  );
}

function Frame({ columns, rows }: { columns: string[]; rows: unknown[][] }) {
  const [query, setQuery] = useState(""),
    [page, setPage] = useState(0);
  const filtered = useMemo(
    () =>
      query
        ? rows.filter((row) =>
            JSON.stringify(row)
              .toLocaleLowerCase()
              .includes(query.toLocaleLowerCase()),
          )
        : rows,
    [query, rows],
  );
  return (
    <div>
      <input
        aria-label="Search analysis table"
        placeholder="Search this table…"
        value={query}
        onChange={(event) => {
          setQuery(event.target.value);
          setPage(0);
        }}
      />
      <div className="evidence-table">
        <table>
          <thead>
            <tr>
              {columns.map((column, index) => (
                <th key={index}>{column}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.slice(page * 100, page * 100 + 100).map((row, index) => (
              <tr key={index}>
                {row.map((cell, index) => (
                  <td key={index}>
                    {cell && typeof cell === "object" ? (
                      <ExpandedValue value={cell} />
                    ) : (
                      <Value value={cell} />
                    )}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="toolbar">
        <button disabled={!page} onClick={() => setPage(page - 1)}>
          Previous rows
        </button>
        <span>
          {filtered.length.toLocaleString()} rows · page {page + 1}
        </span>
        <button
          disabled={(page + 1) * 100 >= filtered.length}
          onClick={() => setPage(page + 1)}
        >
          Next rows
        </button>
      </div>
    </div>
  );
}

export function Value({ value }: { value: unknown }) {
  if (value === null || value === undefined)
    return <span className="muted">Unknown</span>;
  if (typeof value === "boolean") return <span>{value ? "Yes" : "No"}</span>;
  if (typeof value !== "object")
    return (
      <span className={typeof value === "number" ? "numeric" : ""}>
        {String(value)}
      </span>
    );
  if (Array.isArray(value)) {
    if (!value.length) return <span className="muted">None recorded</span>;
    if (value.every((item) => typeof item !== "object" || item === null))
      return <span>{value.map(String).join(", ")}</span>;
    if (
      value.every(
        (item) => item && typeof item === "object" && !Array.isArray(item),
      )
    ) {
      const columns = Array.from(
        new Set(value.flatMap((item) => Object.keys(item))),
      );
      return (
        <Frame
          columns={columns.map(label)}
          rows={value.map((item) => columns.map((column) => item[column]))}
        />
      );
    }
    return <Frame columns={["Value"]} rows={value.map((item) => [item])} />;
  }
  const data = value as Record<string, unknown>;
  if (
    data.type === "frame" &&
    Array.isArray(data.columns) &&
    Array.isArray(data.rows)
  )
    return (
      <Frame
        columns={data.columns.map(String)}
        rows={data.rows as unknown[][]}
      />
    );
  if (data.type === "model" && data.fields)
    return <Value value={data.fields} />;
  if (
    ["timestamp", "date", "datetime"].includes(String(data.type)) &&
    data.value
  )
    return <span>{String(data.value)}</span>;
  return (
    <dl className="facts">
      {Object.entries(data).map(([key, item]) => (
        <div key={key}>
          <dt>{label(key)}</dt>
          <dd>
            {item && typeof item === "object" ? (
              <ExpandedValue value={item} />
            ) : (
              <Value value={item} />
            )}
          </dd>
        </div>
      ))}
    </dl>
  );
}

export function EvidenceTable({
  rows,
  selected,
  onSelect,
  search,
}: {
  rows: Row[];
  selected: string;
  onSelect: (id: string) => void;
  search: string;
}) {
  const container = useRef<HTMLDivElement>(null);
  const columns = useMemo<ColumnDef<Row>[]>(
    () => [
      {
        id: "invoice",
        header: "Invoice / reference",
        accessorFn: (row) =>
          row.record["Invoice #"] || row.record["Details"] || row.id,
      },
      ...[
        "Date",
        "Entry Type",
        "Period Charge (£)",
        "Transaction Amount (£)",
        "Balance (£)",
        "Account",
        "Fuel",
        "Attachment Name",
      ].map((key) => ({
        id: key,
        header: key,
        accessorFn: (row: Row) => row.record[key] as Json,
      })),
      {
        id: "included",
        header: "Included",
        accessorFn: (row) => (row.decision.include ? "Yes" : "No"),
      },
      { id: "status", header: "Review", accessorFn: (row) => row.status },
    ],
    [],
  );
  const table = useReactTable({
    data: rows,
    columns,
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
    globalFilterFn: (row, _columnId, query: string) =>
      JSON.stringify(row.original.record)
        .toLocaleLowerCase()
        .includes(query.toLocaleLowerCase()),
    state: { globalFilter: search },
  });
  const visible = table.getRowModel().rows;
  const virtual = useVirtualizer({
    count: visible.length,
    getScrollElement: () => container.current,
    estimateSize: () => 44,
    overscan: 10,
  });
  return (
    <>
      <p className="muted">
        {visible.length} of {rows.length} records · Select a record to compare
        it with the original.
      </p>
      <div className="evidence-table" ref={container}>
        <table aria-label="Evidence records">
          <thead>
            {table.getHeaderGroups().map((group) => (
              <tr key={group.id}>
                {group.headers.map((header) => (
                  <th
                    key={header.id}
                    aria-sort={
                      header.column.getIsSorted() === "asc"
                        ? "ascending"
                        : header.column.getIsSorted() === "desc"
                          ? "descending"
                          : "none"
                    }
                  >
                    <button onClick={header.column.getToggleSortingHandler()}>
                      {flexRender(
                        header.column.columnDef.header,
                        header.getContext(),
                      )}{" "}
                      {header.column.getIsSorted() === "asc"
                        ? "↑"
                        : header.column.getIsSorted() === "desc"
                          ? "↓"
                          : ""}
                    </button>
                  </th>
                ))}
              </tr>
            ))}
          </thead>
          <tbody>
            <tr aria-hidden="true">
              <td
                colSpan={columns.length}
                style={{
                  height: virtual.getVirtualItems()[0]?.start || 0,
                  padding: 0,
                }}
              />
            </tr>
            {virtual.getVirtualItems().map((item) => {
              const row = visible[item.index];
              return (
                <tr
                  key={row.id}
                  className={row.original.id === selected ? "selected" : ""}
                >
                  {row.getVisibleCells().map((cell, index) => (
                    <td key={cell.id}>
                      {index === 0 ? (
                        <button
                          className="link"
                          onClick={() => onSelect(row.original.id)}
                        >
                          {String(cell.getValue() || "No reference")}
                        </button>
                      ) : (
                        <Value value={cell.getValue()} />
                      )}
                    </td>
                  ))}
                </tr>
              );
            })}
            <tr aria-hidden="true">
              <td
                colSpan={columns.length}
                style={{
                  height: Math.max(
                    0,
                    virtual.getTotalSize() -
                      (virtual.getVirtualItems().at(-1)?.end || 0),
                  ),
                  padding: 0,
                }}
              />
            </tr>
          </tbody>
        </table>
      </div>
    </>
  );
}
