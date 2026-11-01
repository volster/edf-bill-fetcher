import { createContext, useContext, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { api, label } from "./api";

export type Account = {
  id: string;
  username: string;
  csrf: string;
  is_admin: boolean;
  setup_required: boolean;
  active: boolean;
  revision: number;
};
type UserAccount = Omit<Account, "csrf">;
export type WorkspaceInfo = {
  id: string;
  name: string;
  role: "owner" | "editor" | "viewer";
  revision: number;
};
export const Access = createContext<{
  account: Account;
  workspace: WorkspaceInfo | null;
}>({ account: null!, workspace: null });
export const useAccess = () => useContext(Access);
export const useCanEdit = () => {
  const value = useAccess().workspace;
  return !!value && value.role !== "viewer";
};

function ErrorMessage({ error }: { error: unknown }) {
  return error ? (
    <p role="alert" className="error">
      {error instanceof Error ? error.message : String(error)}
    </p>
  ) : null;
}

export function AccountForm({
  account,
  onSaved,
  setup = false,
}: {
  account: Account;
  onSaved: (value: Account) => void;
  setup?: boolean;
}) {
  const [username, setUsername] = useState(
    account.username === "admin" && setup ? "" : account.username,
  );
  const [password, setPassword] = useState("");
  const [confirmation, setConfirmation] = useState("");
  const [current, setCurrent] = useState("");
  const [error, setError] = useState<unknown>();
  const [busy, setBusy] = useState(false);
  const [saved, setSaved] = useState(false);
  return (
    <section className="panel account-form">
      <h1>{setup ? "Set up your account" : "Your account"}</h1>
      <p>
        {setup
          ? "Choose your username and password to continue. The initial credentials stop working when setup is complete."
          : "Changing your credentials signs out other sessions. Your workspace memberships and earlier reviews stay attached to your account."}
      </p>
      <form
        onSubmit={async (event) => {
          event.preventDefault();
          setError(undefined);
          setSaved(false);
          if (password !== confirmation) {
            setError(new Error("The passwords do not match"));
            return;
          }
          setBusy(true);
          try {
            const value = await api<Account>("/account/setup", {
              username,
              password,
              current_password: setup ? null : current,
            });
            onSaved(value);
            setPassword("");
            setConfirmation("");
            setCurrent("");
            setSaved(true);
          } catch (error) {
            setError(error);
          } finally {
            setBusy(false);
          }
        }}
      >
        <label>
          Username
          <input
            autoComplete="username"
            required
            maxLength={100}
            value={username}
            onChange={(e) => setUsername(e.target.value)}
          />
        </label>
        {!setup && (
          <label>
            Current password
            <input
              type="password"
              autoComplete="current-password"
              required
              value={current}
              onChange={(e) => setCurrent(e.target.value)}
            />
          </label>
        )}
        <label>
          New password
          <input
            type="password"
            autoComplete="new-password"
            required
            maxLength={1024}
            value={password}
            onChange={(e) => setPassword(e.target.value)}
          />
        </label>
        <label>
          Confirm password
          <input
            type="password"
            autoComplete="new-password"
            required
            maxLength={1024}
            value={confirmation}
            onChange={(e) => setConfirmation(e.target.value)}
          />
        </label>
        <p className="muted">
          No minimum length is imposed. Choose a password appropriate for who
          can reach this server.
        </p>
        <ErrorMessage error={error} />
        {saved && <p role="status">Account updated.</p>}
        <button className="primary" disabled={busy}>
          {busy ? "Saving…" : setup ? "Finish setup" : "Update account"}
        </button>
      </form>
    </section>
  );
}

type Member = {
  id: string;
  username: string;
  role: WorkspaceInfo["role"];
  active: boolean;
};
type Audit = {
  sequence: number;
  username: string;
  action: string;
  at: number;
  details: unknown;
};
export function WorkspaceSettings({
  onChanged,
  onCreated,
}: {
  onChanged: () => Promise<unknown>;
  onCreated: (value: WorkspaceInfo) => void;
}) {
  const { account, workspace: activeWorkspace } = useAccess();
  const details = useQuery({
    queryKey: ["workspace-settings", activeWorkspace?.id],
    queryFn: () =>
      api<WorkspaceInfo & { members: Member[]; roots: string[] }>(
        `/workspaces/${activeWorkspace!.id}`,
      ),
    enabled: !!activeWorkspace,
  });
  const workspace = details.data || activeWorkspace;
  const [name, setName] = useState("");
  const [rename, setRename] = useState(activeWorkspace?.name || "");
  const [member, setMember] = useState("");
  const [role, setRole] = useState<WorkspaceInfo["role"]>("viewer");
  const [error, setError] = useState<unknown>();
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState("");
  const owner = activeWorkspace?.role === "owner";
  const members = {
    data: details.data?.members,
    error: details.error,
    refetch: details.refetch,
  };
  const audit = useQuery({
    queryKey: ["audit", workspace?.id],
    queryFn: () => api<Audit[]>(`/workspaces/${workspace!.id}/audit`),
    enabled: owner,
  });
  const roots = { data: details.data?.roots };
  const availableRoots = useQuery({
    queryKey: ["server-roots"],
    queryFn: () => api<string[]>("/server-roots"),
    enabled: account.is_admin && owner,
  });
  async function save(action: () => Promise<unknown>) {
    setBusy(true);
    setError(undefined);
    setNotice("");
    try {
      await action();
      await onChanged();
      if (activeWorkspace) await members.refetch();
      if (owner) await audit.refetch();
      setNotice("Saved.");
    } catch (error) {
      setError(error);
    } finally {
      setBusy(false);
    }
  }
  return (
    <>
      <header className="page-heading">
        <div className="eyebrow">Workspace access</div>
        <h1>Workspaces & sharing</h1>
        <p>
          Each workspace has its own uploads, runs and review drafts. Sharing
          grants access to the whole workspace, including original documents and
          downloads.
        </p>
      </header>
      <ErrorMessage error={error || members.error} />
      {workspace && (
        <button
          disabled={busy}
          onClick={() => {
            if (
              !window.confirm(
                "Reload sharing and discard unapplied form fields?",
              )
            )
              return;
            void details.refetch().then((value) => {
              if (value.data) {
                setRename(value.data.name);
                setMember("");
                setRole("viewer");
                setError(undefined);
              }
            });
          }}
        >
          Reload sharing
        </button>
      )}
      {notice && <p role="status">{notice}</p>}
      <section className="panel">
        <h2>Create a workspace</h2>
        <form
          onSubmit={(event) => {
            event.preventDefault();
            void save(async () => {
              const value = await api<WorkspaceInfo>("/workspaces", { name });
              setName("");
              onCreated(value);
            });
          }}
        >
          <label>
            New workspace name
            <input
              required
              maxLength={120}
              value={name}
              onChange={(e) => setName(e.target.value)}
            />
          </label>
          <button disabled={busy}>Create workspace</button>
        </form>
      </section>
      {workspace && (
        <>
          <section className="panel">
            <h2>{workspace.name}</h2>
            <p>
              Your access: <strong>{label(workspace.role)}</strong>.
            </p>
            {owner && (
              <form
                onSubmit={(event) => {
                  event.preventDefault();
                  void save(() =>
                    api(
                      `/workspaces/${workspace.id}`,
                      { name: rename, revision: workspace.revision },
                      "PUT",
                    ),
                  );
                }}
              >
                <label>
                  Workspace name
                  <input
                    required
                    maxLength={120}
                    value={rename}
                    onChange={(e) => setRename(e.target.value)}
                  />
                </label>
                <button disabled={busy || !details.data}>
                  Rename workspace
                </button>
              </form>
            )}
            <dl className="role-guide">
              <dt>Owner</dt>
              <dd>
                Manage members and sharing, review evidence and publish packs.
              </dd>
              <dt>Editor</dt>
              <dd>Upload evidence, edit shared drafts and publish packs.</dd>
              <dt>Viewer</dt>
              <dd>
                Read evidence and download outputs. Cannot change drafts or run
                analysis.
              </dd>
            </dl>
          </section>
          <section className="panel">
            <h2>Members</h2>
            <p>
              Changes take effect on subsequent requests, including existing
              sign-ins.
            </p>
            <div className="member-list">
              {members.data?.map((item) => (
                <div className="member-row" key={item.id}>
                  <strong>
                    {item.username}
                    {!item.active && " · disabled"}
                  </strong>
                  {owner ? (
                    <>
                      <label>
                        Access for {item.username}
                        <select
                          aria-label={`Access for ${item.username}`}
                          value={item.role}
                          disabled={busy || !details.data}
                          onChange={(e) =>
                            void save(() =>
                              api(
                                `/workspaces/${workspace.id}/members`,
                                {
                                  username: item.username,
                                  role: e.target.value,
                                  revision: workspace.revision,
                                },
                                "PUT",
                              ),
                            )
                          }
                        >
                          <option value="viewer">Viewer</option>
                          <option value="editor">Editor</option>
                          <option value="owner">Owner</option>
                        </select>
                      </label>
                      <button
                        disabled={busy || !details.data}
                        onClick={() => {
                          if (
                            window.confirm(
                              `Remove ${item.username} from this workspace?`,
                            )
                          )
                            void save(() =>
                              api(
                                `/workspaces/${workspace.id}/members`,
                                {
                                  username: item.username,
                                  role: null,
                                  revision: workspace.revision,
                                },
                                "PUT",
                              ),
                            );
                        }}
                      >
                        Remove member
                      </button>
                    </>
                  ) : (
                    <span>{label(item.role)}</span>
                  )}
                </div>
              ))}
            </div>
            {owner && (
              <form
                onSubmit={(event) => {
                  event.preventDefault();
                  void save(async () => {
                    await api(
                      `/workspaces/${workspace.id}/members`,
                      { username: member, role, revision: workspace.revision },
                      "PUT",
                    );
                    setMember("");
                  });
                }}
              >
                <h3>Share with an existing account</h3>
                <label>
                  Member username
                  <input
                    required
                    autoComplete="off"
                    value={member}
                    onChange={(e) => setMember(e.target.value)}
                  />
                </label>
                <label>
                  Member access
                  <select
                    aria-label="Member access"
                    value={role}
                    onChange={(e) =>
                      setRole(e.target.value as WorkspaceInfo["role"])
                    }
                  >
                    <option value="viewer">Viewer</option>
                    <option value="editor">Editor</option>
                    <option value="owner">Owner</option>
                  </select>
                </label>
                <button disabled={busy || !details.data}>Add member</button>
                <p className="muted">
                  A server administrator can create an account from User
                  administration.
                </p>
              </form>
            )}
          </section>
          <section className="panel">
            <h2>Server source folders</h2>
            <p>
              Browser uploads are always available to editors. Server folders
              require an explicit grant from a server administrator who owns
              this workspace.
            </p>
            {account.is_admin && owner ? (
              <>
                {availableRoots.data?.map((value) => (
                  <label className="check" key={value}>
                    <input
                      type="checkbox"
                      checked={roots.data?.includes(value) || false}
                      disabled={busy || !details.data}
                      onChange={() =>
                        void save(async () => {
                          const labels = roots.data?.includes(value)
                            ? roots.data.filter((item) => item !== value)
                            : [...(roots.data || []), value];
                          await api(
                            `/workspaces/${workspace.id}/roots`,
                            { labels, revision: workspace.revision },
                            "PUT",
                          );
                        })
                      }
                    />
                    {value}
                  </label>
                ))}
                {!availableRoots.data?.length && (
                  <p>No server source folders are configured.</p>
                )}
              </>
            ) : (
              <p>{roots.data?.join(", ") || "No server folders granted."}</p>
            )}
          </section>
          {owner && (
            <section className="panel">
              <h2>Access & review history</h2>
              <ErrorMessage error={audit.error} />
              <p>
                Showing the first 200 events. Access changes retain their actor
                and time; passwords are never recorded.
              </p>
              <ul className="audit-list">
                {audit.data?.map((item) => (
                  <li key={item.sequence}>
                    <time>{new Date(item.at * 1000).toLocaleString()}</time> ·{" "}
                    {item.username} · {label(item.action.replaceAll(".", " "))}
                    <pre>{JSON.stringify(item.details)}</pre>
                  </li>
                ))}
              </ul>
            </section>
          )}
        </>
      )}
    </>
  );
}

export function Users() {
  const users = useQuery({
    queryKey: ["users"],
    queryFn: () => api<UserAccount[]>("/users"),
  });
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [admin, setAdmin] = useState(false);
  const [error, setError] = useState<unknown>();
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState("");
  async function save(action: () => Promise<unknown>) {
    setError(undefined);
    setBusy(true);
    setNotice("");
    try {
      await action();
      await users.refetch();
    } catch (error) {
      setError(error);
    } finally {
      setBusy(false);
    }
  }
  return (
    <>
      <header className="page-heading">
        <div className="eyebrow">Server administration</div>
        <h1>User administration</h1>
        <p>
          Create accounts and manage server access. Administrator status does
          not grant access to other users’ evidence; workspace membership is
          separate.
        </p>
      </header>
      <ErrorMessage error={error || users.error} />
      {notice && <p role="status">{notice}</p>}
      <section className="panel">
        <h2>Create an account</h2>
        <form
          onSubmit={(event) => {
            event.preventDefault();
            void save(async () => {
              await api("/users", { username, password, is_admin: admin });
              setUsername("");
              setPassword("");
              setAdmin(false);
              setNotice(
                "Account created. Share the initial credentials with its user; they will choose their own credentials at first sign-in.",
              );
            });
          }}
        >
          <label>
            New username
            <input
              required
              maxLength={100}
              autoComplete="off"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
            />
          </label>
          <label>
            Initial password
            <input
              required
              type="password"
              maxLength={1024}
              autoComplete="new-password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
            />
          </label>
          <label className="check">
            <input
              type="checkbox"
              checked={admin}
              onChange={(e) => setAdmin(e.target.checked)}
            />
            Server administrator
          </label>
          <button disabled={busy}>Create account</button>
        </form>
      </section>
      <section className="panel">
        <h2>Accounts</h2>
        {users.data?.map((user) => (
          <div className="member-row" key={user.id}>
            <div>
              <strong>{user.username}</strong>
              <p>
                {user.active ? "Active" : "Disabled"}
                {user.setup_required && " · setup required"}
              </p>
            </div>
            <label className="check">
              <input
                type="checkbox"
                checked={user.is_admin}
                disabled={busy}
                onChange={() =>
                  void save(() =>
                    api(
                      `/users/${user.id}`,
                      {
                        revision: user.revision,
                        active: user.active,
                        is_admin: !user.is_admin,
                      },
                      "PUT",
                    ),
                  )
                }
              />
              Administrator
            </label>
            <button
              disabled={busy}
              onClick={() =>
                void save(() =>
                  api(
                    `/users/${user.id}`,
                    {
                      revision: user.revision,
                      active: !user.active,
                      is_admin: user.is_admin,
                    },
                    "PUT",
                  ),
                )
              }
            >
              {user.active ? "Disable account" : "Enable account"}
            </button>
            <ResetPassword user={user} onSave={save} busy={busy} />
          </div>
        ))}
      </section>
    </>
  );
}

function ResetPassword({
  user,
  onSave,
  busy,
}: {
  user: UserAccount;
  onSave: (action: () => Promise<unknown>) => Promise<void>;
  busy: boolean;
}) {
  const [password, setPassword] = useState("");
  return (
    <details>
      <summary>Reset password for {user.username}</summary>
      <form
        onSubmit={(event) => {
          event.preventDefault();
          void onSave(async () => {
            await api(`/users/${user.id}/password`, { password });
            setPassword("");
          });
        }}
      >
        <label>
          Temporary password for {user.username}
          <input
            required
            type="password"
            autoComplete="new-password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
          />
        </label>
        <p>
          Revokes all sessions and requires account setup at the next sign-in.
        </p>
        <button disabled={busy}>Reset password</button>
      </form>
    </details>
  );
}
