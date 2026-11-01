# Accounts, workspaces and database administration

[Web quick start](../README.md) · [Browser guide](WEB_UI.md) · [Documentation index](README.md)

## First login and account recovery

A new server creates **admin / admin**. Sign in, choose your own username and
password, and confirm the password. Evidence and administrative operations remain
unavailable until setup is complete. Setup replaces the initial credentials;
restarting the app does not recreate the default account. There is no minimum
password length. Passwords are hashed with a unique salt, and raw session tokens
are not stored in the database.

Existing installations keep their accounts and passwords. Usernames are trimmed
and matched without case sensitivity; passwords retain their exact characters.
Changing credentials retains the account's identity, workspace memberships and
earlier review history, and signs out earlier sessions.

Use **Your account** to change your username or password later. This requires
your current password. A server administrator can reset a user's password in
**User administration**; the user then completes setup at the next sign-in.
Disabling an account revokes its sessions and prevents new sign-ins. Account
records and their evidence history remain available to authorized members.

If you cannot sign in, run the recovery command on the server using the same
data directory and database configuration:

```console
edf-web --data web-data --reset-password USERNAME
```

For Compose:

```console
docker compose run --rm app --reset-password USERNAME
```

The command prompts for the new password. `EDF_WEB_PASSWORD` supplies it for
automation instead of a prompt; remove that variable when it is no longer
needed. `--init-user USERNAME` remains available for creating an account from the
server terminal; browser account administration is the usual workflow.

## Accounts and workspace permissions

Server administration and evidence access are separate. An administrator can
create, reset, enable or disable accounts and change administrator status.
Administrator status alone does not grant access to another workspace.

Each new account receives a private workspace. Any account can create additional
workspaces. Select the active workspace in the sidebar, then use **Workspaces &
sharing** to rename it or manage its members.

| Role   | Read evidence and download packs | Upload, analyse, edit and publish | Rename and manage sharing |
| ------ | -------------------------------- | --------------------------------- | ------------------------- |
| Viewer | Yes                              | No                                | No                        |
| Editor | Yes                              | Yes                               | No                        |
| Owner  | Yes                              | Yes                               | Yes                       |

Owners share with an existing **username**, choosing a role. Sharing covers the
whole workspace: uploads, original documents, runs, review drafts and downloads.
It does not copy evidence into the recipient's personal workspace. Removing a
membership removes access to those objects, including from existing sign-ins.
Files already downloaded to another computer cannot be recalled.

Workspace owners cannot remove or demote the last active owner. The server must
also retain an active administrator. Access changes use database transactions;
stale sharing changes are rejected so an older tab cannot silently restore an
earlier grant. Applying edits to a shared review draft also rejects stale
revisions. An edit saves a draft; publishing creates a fresh reviewed pack.

Viewers see read-only controls and are also denied mutations by the server.
Object identifiers, imported results, intake references and parent/comparison
runs cannot be used to cross workspace boundaries. Queued jobs recheck their
creator's activation and editor access before processing.

Owners can inspect **Access & review history** for recorded membership, folder,
upload, run and draft changes. Events retain the actor's stable account identity
and time. Passwords and session tokens are excluded from audit details.

### Server source folders

Browser uploads need no server-folder grant. To browse existing server-side
evidence, first configure a source root with `--source-root NAME=PATH` and, for
Docker, mount that folder read-only. A server administrator who is also an owner
of the receiving workspace grants the folder in **Workspaces & sharing**.
Source folders must be separate from the app's managed data directory. Roots
inside it, or containing it, are rejected to keep shared-folder intake from
reading another workspace's managed evidence.

On first startup, configured folders are assigned once to the initial workspace.
On upgrade they are assigned to the imported shared workspace. New workspaces
receive no server-folder grants. Revoked grants stay revoked after restart, and
queued jobs recheck them before reading source files.

## Database configuration

**Docker Compose uses PostgreSQL 17 by default.** The database is on Compose's
internal network and does not publish a host port. Evidence stays in the
`evidence-data` volume; PostgreSQL keeps metadata in `database-data`.

`EDF_POSTGRES_PASSWORD` in the Compose `.env` file sets the database password on
first initialization. Use a URL-safe value with the supplied connection URL.
Changing this variable after the database volume exists does not rotate the
existing PostgreSQL role's password; update the role and connection configuration
together. This setting is separate from browser account passwords.

A source installation connects to PostgreSQL through `EDF_DATABASE_URL`, for
example in Bash:

```bash
export EDF_DATABASE_URL='postgresql+psycopg://edf:DATABASE_PASSWORD@localhost:5432/edf'
edf-web --data web-data
```

Or PowerShell:

```powershell
$env:EDF_DATABASE_URL = 'postgresql+psycopg://edf:DATABASE_PASSWORD@localhost:5432/edf'
edf-web --data web-data
```

Create the PostgreSQL database and its owning role first. The application role
needs permission to create and migrate its tables and indexes. Standard
`postgresql://` URLs are also accepted. The web extra installs the Psycopg driver.
`--database-url` can override the environment; environment configuration avoids
putting credentials into a shell command's argument list.

Without a URL, source installations use `web-data/catalogue.sqlite`. SQLite has
the same accounts, memberships, roles and review features, and remains supported
for local installations. Neither database is part of the desktop/CLI evidence
interchange format: `result.json` and complete packs stay portable.

Run one app instance per data directory. PostgreSQL makes account and catalogue
updates transactional; it does not replace the shared evidence filesystem or
turn the dispatcher into a multi-host worker service. Workers inherit the same
database URL as the server.
A database-wide lease also prevents a second PostgreSQL dispatcher from starting
with a different data directory against the same database.

## Upgrading an existing installation

Back up first and stop the old server. Database migrations run automatically.
The old catalogue's users and accepted evidence move into an **Imported
workspace**. Existing users retain access as owners; the first migrated account
becomes a server administrator. Existing browsers sign in again.
Jobs queued or running under the former shared-catalogue permissions become
interrupted during this upgrade; create a fresh run under a current account.
Ordinary restarts retain newly queued jobs.

When switching to PostgreSQL with the same data directory, an existing
`catalogue.sqlite` is imported **only when PostgreSQL has no accounts or evidence
entries**. This preserves account IDs, password hashes, workspace membership,
run IDs, draft history, progress events and references to existing files. The
copy commits as one transaction. It does not merge into or overwrite a populated
PostgreSQL catalogue, and does not delete the SQLite file.

Old evidence directories retain their paths so accepted source links keep
working. New uploads and runs use workspace-specific directories. After upgrade,
verify the account list, workspace membership and a known original document
before retiring your backup or old database. Do not run old and new servers
against the same data directory during migration.

## Backup and restore

For local SQLite, stop the app and copy the **whole data directory**, including
the catalogue, any SQLite sidecar files, uploads, runs and workspaces.

For PostgreSQL, back up **both the database and evidence files** while the app is
stopped. A database dump alone cannot restore uploaded originals; an evidence
folder alone cannot restore accounts, grants or saved drafts. Allow running jobs
to finish before taking a backup.

For Compose, create a `backups` directory in the checkout, then:

```console
docker compose stop app
docker compose exec database pg_dump -U edf -d edf -Fc -f /tmp/catalogue.dump
docker compose cp database:/tmp/catalogue.dump ./backups/catalogue.dump
docker compose run --name edf-evidence-backup --no-deps --entrypoint python app -c "import tarfile; archive = tarfile.open('/tmp/evidence-data.tar.gz', 'w:gz'); archive.add('/data', arcname='data'); archive.close()"
docker cp edf-evidence-backup:/tmp/evidence-data.tar.gz ./backups/evidence-data.tar.gz
docker rm edf-evidence-backup
docker compose start app
```

The dump is written inside the database container and copied as bytes, avoiding
shell text-redirection problems on Windows. These files contain private evidence
and credential hashes; keep the pair together in your chosen backup location.

To restore, keep the app stopped, restore the dump into an empty PostgreSQL
database using `pg_restore --no-owner --exit-on-error`, and restore the archive's
`data/` contents into the app's data volume. Preserve the original directory
structure and ensure UID 1000 can read and write it in Docker. Start the app with
the restored database and data volume together. Check sign-in, memberships,
drafts and original downloads before resuming work.

Do not use `docker compose down --volumes` as an update command: it deletes the
database and evidence volumes. A normal `docker compose up --build -d` retains
them. Backups remain necessary before schema upgrades.
