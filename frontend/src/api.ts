import { createSHA256 } from "hash-wasm";
import type { components } from "./generated";
export type Json =
  null | boolean | number | string | Json[] | { [key: string]: Json };
export type Entry<T> = { id: string; revision: number; kind: string; data: T };
export type Source = components["schemas"]["Source"];
export type JobData = {
  request: { name: string; mode: string; formats: string[] };
  status: string;
  created: number;
  error: string | null;
};
export type Row = {
  id: string;
  included: boolean;
  record: Record<string, Json>;
  decision: { changes: Record<string, Json>; reason: string; include: boolean };
  status: string;
};
export type Review = {
  revision: number;
  draft: { claims: { groups: Claim[]; context: Context } } & Record<
    string,
    Json
  >;
  rows: Row[];
  ledger: Record<string, Json>;
  claims: Record<string, Json>;
  revisions: Record<string, Json>;
  summary: Record<string, string | null>;
  models: Record<string, Json>;
  can_undo: boolean;
  can_redo: boolean;
};
export type Context = {
  complaint_summary: string;
  requested_outcome: string;
  include_pack: boolean;
};
export type Credit = {
  row_id: string;
  amount: string;
  treatment: string;
  reason: string;
};
export type Claim = {
  id: string;
  title: string;
  status: string;
  members: string[];
  challenged: string | null;
  reduction: string | null;
  credits: Credit[];
  reason: string;
  account_link_reason: string;
};
export type Schema = {
  review_fields: Record<string, string>;
  settings: Record<string, string>;
  sections: { key: string; title: string }[];
};
let csrf = "";
let workspace = "";
let account = "";
export function setWorkspace(value: string, accountId: string) {
  workspace = value;
  account = accountId;
}
export const uploadStorageKey = () => `edf-upload:${account}:${workspace}`;
export function setCsrf(value: string) {
  csrf = value;
}
export async function api<T>(
  path: string,
  body?: unknown,
  method?: string,
  signal?: AbortSignal,
): Promise<T> {
  const response = await fetch("/api" + path, {
    method: method || (body === undefined ? "GET" : "POST"),
    credentials: "same-origin",
    signal,
    headers: {
      "Content-Type": "application/json",
      "X-CSRF-Token": csrf,
      "X-Workspace-ID": workspace,
    },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  const value = await response.json();
  if (response.status === 401 && path !== "/login" && path !== "/session")
    window.dispatchEvent(new Event("edf-session-expired"));
  if (!response.ok)
    throw new Error(
      typeof value.detail === "string"
        ? value.detail
        : JSON.stringify(value.detail),
    );
  return value as T;
}
export async function uploadAsset(
  kind: Source["kind"],
  files: File[],
  name: string,
  progress: (done: number, total: number, stage?: string) => void,
  resumeId?: string,
  signal?: AbortSignal,
): Promise<string> {
  const storageKey = uploadStorageKey();
  if (!files.length || files.length > 100_000)
    throw new Error("Select between 1 and 100,000 files for one source.");
  if (files.some((file) => file.size > 64 * 1024 ** 3))
    throw new Error("Each uploaded file must be no larger than 64 GiB.");
  const paths = files.map((file) => file.webkitRelativePath || file.name);
  if (new Set(paths).size !== paths.length)
    throw new Error(
      "The selection contains duplicate paths. Choose a folder to preserve subfolders.",
    );
  const selectionDigest = await createSHA256();
  selectionDigest.init();
  files
    .map((file, index) => JSON.stringify([paths[index], file.size]))
    .sort()
    .forEach((value) => selectionDigest.update(value + "\n"));
  const selection = selectionDigest.digest("hex");
  const previous = JSON.parse(sessionStorage.getItem(storageKey) || "null");
  if (resumeId && previous?.selection && previous.selection !== selection)
    throw new Error(
      "Select all original files with the same paths and sizes as the pending upload.",
    );
  const asset = resumeId
    ? { id: resumeId }
    : await api<Entry<unknown>>("/assets", { kind, name }, undefined, signal);
  const batch =
    resumeId && previous?.batch
      ? previous.batch
      : Array.from(crypto.getRandomValues(new Uint8Array(16)), (value) =>
          value.toString(16).padStart(2, "0"),
        ).join("");
  sessionStorage.setItem(
    storageKey,
    JSON.stringify({ id: asset.id, kind, name, batch, selection }),
  );
  const total = files.reduce((sum, file) => sum + file.size, 0);
  let done = 0;
  let order = 0;
  for (const file of files) {
    signal?.throwIfAborted();
    const path = file.webkitRelativePath || file.name;
    const query = new URLSearchParams({ path });
    const pending = await api<{ offset: number; sha256: string | null }>(
      `/assets/${asset.id}/offset?${query}`,
      undefined,
      undefined,
      signal,
    );
    let offset = pending.offset;
    const digest = await createSHA256();
    digest.init();
    for (let position = 0; position < offset; position += 4 * 1024 * 1024) {
      signal?.throwIfAborted();
      digest.update(
        new Uint8Array(
          await file
            .slice(position, Math.min(offset, position + 4 * 1024 * 1024))
            .arrayBuffer(),
        ),
      );
    }
    if (offset) {
      const prefixState = digest.save();
      if (digest.digest("hex") !== pending.sha256)
        throw new Error(
          "The selected file does not match the pending upload. Choose the original file or create a fresh source.",
        );
      digest.load(prefixState);
    }
    if (offset > file.size)
      throw new Error(
        "The selected file is smaller than the pending upload. Create a new source.",
      );
    done += offset;
    if (offset < file.size || pending.sha256 === null)
      do {
        const chunk = file.slice(offset, offset + 4 * 1024 * 1024);
        const response = await fetch(
          `/api/assets/${asset.id}/content?${query}&offset=${offset}`,
          {
            method: "PUT",
            signal,
            headers: { "X-CSRF-Token": csrf, "X-Workspace-ID": workspace },
            body: chunk,
          },
        );
        const result = await response.json();
        if (response.status === 401)
          window.dispatchEvent(new Event("edf-session-expired"));
        if (!response.ok) throw new Error(result.detail);
        digest.update(new Uint8Array(await chunk.arrayBuffer()));
        offset = result.offset;
        done += chunk.size;
        progress(done, total);
      } while (offset < file.size);
    progress(done, total, `Verifying ${file.name}…`);
    await api(
      `/assets/${asset.id}/file-complete?${query}`,
      {
        size: file.size,
        sha256: digest.digest("hex"),
        batch_id: batch,
        order: order++,
      },
      undefined,
      signal,
    );
  }
  await api(
    `/assets/${asset.id}/seal`,
    {
      batch_id: batch,
      files_count: files.length,
    },
    undefined,
    signal,
  );
  sessionStorage.removeItem(storageKey);
  return asset.id;
}
export const terminal = new Set([
  "complete",
  "review_required",
  "empty",
  "failed",
  "cancelled",
  "interrupted",
]);
export const label = (key: string) =>
  key.replaceAll("_", " ").replace(/\b\w/g, (value) => value.toUpperCase());
