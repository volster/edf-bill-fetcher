# Upload and concurrency validation

[Browser guide](WEB_UI.md) · [Administration](WEB_ADMINISTRATION.md) · [Development](DEVELOPMENT.md)

Use synthetic evidence when testing hosting capacity. Upload limits bound
individual inputs; they do not establish how quickly a particular server can
process a full collection. Disk, database latency, OCR requirements and report
size all affect the result.

## Repeatable HTTP stress check

After installing the Python `[dev,msg,web]` extras and native prerequisites from
[source setup](BUILDING.md), run this from the repository root:

```console
python -m tests.web_upload_stress
```

The default check starts a disposable local server on an available loopback port
and uses four real accounts to upload 2,000 files into a shared workspace. It also
uploads and hashes a 64 MiB original and inspects a ZIP containing 2,000 nested
members. Unsupported ZIP contents must remain visible in the intake inventory;
they are deliberately synthetic binary files, not bills. Private-workspace
access must still reject another workspace's evidence.

The check creates temporary evidence and removes its server, files and database
when finished. It prints JSON containing elapsed time, input counts and sizes.
The ordinary test suite does not run this opt-in stress check.

For a smaller run or a different workload:

```console
python -m tests.web_upload_stress --files 100 --users 2 --large-mib 8
```

`--files` counts files across all uploaders and independently controls the ZIP
member count. Accepted values are 2–10 users, at least one file per user up to
100,000 files, and 1–1024 MiB for the large original. Leave space for the uploaded
ZIP, expanded members, original and catalogue together.

### PostgreSQL

By default the check uses isolated SQLite storage. To exercise PostgreSQL, set
`EDF_TEST_DATABASE_URL` to a **dedicated test database**, or supply
`--database-url`, using the `postgresql+psycopg://` scheme. The database role must
be allowed to create schemas. The check creates a unique `edf_stress_*` schema
and drops that schema afterward; it does not use the application's operational
tables. Keep real credentials out of committed commands and logs.

Do not interpret one elapsed time as a throughput guarantee. Compare runs using
the same machine, workload, database backend and competing load.

## Failure and concurrency regressions

The regular suites cover failures independently of the stress check:

```console
python -m pytest tests/test_web.py tests/test_web_accounts.py
npm test --prefix frontend
```

- Browser tests interrupt a multi-chunk upload before and after a server write,
  reject a changed original, recover a lost final confirmation, include empty
  files, reject an incomplete resumed selection before sending more bytes,
  pause/resume, and reset the file picker after completion.
- Storage regressions inject full-disk, quota and I/O errors after partial writes,
  including failure to roll back. Finalisation must reject incomplete evidence;
  a later resume verifies the actual prefix and completes the original bytes.
  Both batch confirmation and the older finalisation request reject an unfinished
  write. Folder-member registration and its collection count commit in one
  database transaction, so a rejected registration cannot leave an orphan entry.
- Concurrent editors writing the same offset produce one accepted chunk and a
  conflict. Uploads to different files retain both manifest entries.
- Concurrent review saves against one draft revision accept one change and
  reject the stale save, retaining the winning actor in the audit history.
- Account tests also cover live access revocation, stale membership changes and
  concurrent owner/admin changes. Set `EDF_TEST_DATABASE_URL` to exercise their
  real PostgreSQL cases; otherwise those cases are reported as skips.

Browser CI runs Chromium, Firefox and WebKit with retained failure traces.
The frontend TypeScript check includes the browser tests and Playwright
configuration as well as application code.
WebKit can finish a native folder selection while Playwright is still waiting
for its input notification. The folder test accepts that specific WebKit timeout
only after checking the complete native file list against every expected relative
path and size; it still uploads, inspects and extracts the selected evidence.
An incomplete or incorrect selection fails the check.

These checks establish upload and access behavior for the exercised workloads.
They do not establish maximum capacity, multi-host worker support, processing
checkpoints within an archive, or the validity of financial calculations. Those
calculations remain covered by the shared financial and evidence contracts.
