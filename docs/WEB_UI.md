# Browser workspace and network hosting

[Web quick start](../README.md) · [Documentation index](README.md) · [Source setup](BUILDING.md) · [Architecture](ARCHITECTURE.md)

The `feature/web-ui` branch adds a browser interface to the same extraction,
analysis, corrections and publication services used by desktop and CLI. It can
run from a Python installation or Docker. The web server processes evidence;
visiting computers need only a browser. No VNC is involved.

This is a shared household/team workspace. Named accounts can access the same
runs and evidence; there are no separate private workspaces or per-case roles.
Authentication is required even when the server listens only on localhost.

## Run from source

Use the supported Python environment and native prerequisites from [BUILDING.md](BUILDING.md).
Tesseract and English language data are native requirements; pip does not supply them.
With the environment activated, install the optional web dependencies:

```console
python -m pip install -c .github/ci-constraints.txt -e ".[web,msg]"
npm ci --prefix frontend
npm run build --prefix frontend
edf-web --data web-data --init-user dad
edf-web --data web-data
```

The account command prompts for a password of at least 12 characters. Open
`http://127.0.0.1:8000`. `python main.py --web` accepts the same server options.
Node 24 is needed to build/develop the frontend, not to run an already built
Python wheel or the container. Browser assets, PDF.js workers/fonts/CMaps and
catalogue migrations are included in the web package's built distribution.

To allow another computer on your network to connect:

```console
edf-web --data web-data --host 0.0.0.0 --port 8000 --source-root "bills=/path/to/evidence"
```

Dad opens `http://YOUR-COMPUTER:8000`. Allow the chosen port through your local
firewall if necessary. On Windows, a source root can be
`--source-root "bills=C:\Evidence"`. Roots are optional: browser uploads work
without them. Only configured roots are browsable, and uploads/runs have managed
directories. Root labels must be unique. Passwords can be supplied through
`EDF_WEB_PASSWORD` during account creation for automation; do not keep that
variable in a running server's environment unnecessarily.

For an HTTPS reverse proxy, preserve the public Host header and use
`--secure-cookie`. The application deliberately has no permissive credentialed
CORS configuration. Run one server process per data directory; a process lock
prevents two dispatchers from managing the same jobs.

## Run with Docker Compose

```console
docker compose build
docker compose run --rm app --init-user dad
docker compose up -d
```

The password is prompted. Open `http://127.0.0.1:8000`. The image includes the
Python/native dependencies, Tesseract English OCR and compiled browser interface.
It runs as an unprivileged user and keeps evidence/settings/catalogue in the
`evidence-data` volume. A container replacement does not discard the workspace.

Compose publishes only on localhost by default. Set `EDF_WEB_BIND_ADDRESS=0.0.0.0`
before starting Compose to permit LAN connections. `EDF_WEB_PORT` changes the
published port. Optional server-side evidence needs an explicit read-only bind
mount and a corresponding `--source-root` in the service command. Linux bind
mount ownership must permit the service's UID 1000 to write its data directory;
the named volume avoids host UID mismatch for the default setup.

Behind an organisation's HTTPS build proxy, Docker can accept a trusted CA
bundle as the optional `build_ca` BuildKit secret. It is mounted only during
dependency installation and is not copied into the image. TLS verification
remains enabled. Normal builds need no custom CA option.

## Add everything, then inspect

The default **All evidence · detect automatically** source accepts multiple
files, a directory tree, ZIPs, or a configured server folder. Dedicated source
types remain available when deliberately restricting an import.

1. Add/upload the original evidence. Directory uploads retain relative paths.
2. Choose **Inspect everything**. Inspection runs as a reconnectable background job.
3. Check the paginated inventory: ready originals, archive containers, unsupported
   files, failures and items needing an import choice are separate statuses.
4. Import a discovered saved analysis/workbook explicitly, or acknowledge the
   unprocessed files and create a pack from the recognised originals.

PDF signatures, Outlook archive signatures, MSG containers, email headers and
HTML content inform classification. Unknown content does not acquire invented
billing values. An existing output folder is skipped. ZIPs containing an earlier
`result.json` need an import choice so generated reports are not mistaken for new
bills. Existing results/workbooks are not silently merged into fresh extraction.

Inspection retains expanded ZIP members and extraction reuses that plan rather
than decompressing large archives twice. Member/container names and container
hashes survive as provenance. Files changed after inspection need inspection
again. An unavailable or unsupported source remains visible in the inventory and
review notes. Uploaded originals remain in the workspace; unsupported files are
not automatically published as analysed evidence. Mixed extractions also retain
`intake_inventory.json` in their output pack.

Uploads use 4 MiB browser chunks, with an 8 MiB server request ceiling. Resume
checks the existing prefix against the selected original. Files have independent
catalogue records, completion checks compare size/SHA256, and finalisation checks
that the whole selected batch is complete. If interrupted, select the same files
and choose **Resume pending upload** in that browser tab. A changed original or
different/incomplete selection is rejected instead of producing a mixed file.
Keep files unchanged until upload completion.

Current bounds are 64 GiB per uploaded file, 100,000 files per collection,
64 GiB aggregate ZIP expansion per discovery, and three archive nesting levels.
ZIP traversal, symbolic links, encrypted members and conflicting portable names
are rejected. Free space is checked before chunk writes and archive expansion.
Large ZIPs still require disk space for the compressed upload and expanded files.
These bounds are explicit limits, not claims of unlimited storage or processing.

## Review, regenerate and share

Open a completed run to search/sort records, inspect extracted text and original
PDF pages, correct/exclude/reinstate records, verify them, and preview changes.
Revision relationships, dispute positions, consolidated claims, credit allocations,
overlap decisions and complaint wording use the existing Python review contracts.
Report choices include cover, contents and all registered sections.

Applying an edit saves the draft on the server; unsubmitted form text is separate.
Draft revisions reject stale writes from another browser/tab. Undo/redo applies to
saved draft changes. A reviewed pack is a fresh run, leaving its baseline intact.
Opening a saved accepted analysis does not recalculate it. Explicit analytical
setting overrides create a new analysis; changing output sections alone does not.

Excel remains part of the workflow: download the editable workbook, amend its
designated tables, upload it through workbook regeneration and create a fresh
pack. Download the complete pack when sharing so the workbook, reports, evidence
manifest and original files remain together. Inspect completion/review/failure
status before relying on any outputs.

Source-backed analyses may retain configured benchmark paths. New custom cap
locations must be within configured roots; use `ROOT:relative/path.csv` in web
settings or a permitted absolute server path. An imported analysis's accepted
calculations remain viewable without its old benchmark file. Reanalysis or human
review needs an approved accessible resource location.

## Other entry points have the same intake/review operations

```console
python main.py --inspect "mixed-evidence" --output-dir "inspection"
python main.py --extract --intake-plan "inspection/intake.json" -o "output/evidence.xlsx"
python main.py --extract --auto "mixed-evidence" --auto "more.zip" -o "output/evidence.xlsx"
```

Desktop **Build evidence pack** includes mixed evidence, an inspection action and
an intake-plan input. Inspection produces an inventory workbook and JSON; keep
its `intake_files/` folder until extraction finishes. Desktop **Compare runs**
uses the ordinary managed worker and supports accepted results/records exports.

`--review-command` accepts a JSON file with `operation` and `arguments`, plus
`--result` and `--draft`. It exposes record, revision, dispute, claim and complaint
commands. `--preview FILE.json` saves a recalculated preview; `--publish DIRECTORY`
creates a fresh reviewed pack. `--formats html pdf docx xlsx` selects publications.
Undo/redo use the adjacent bounded `.history.json` journal, checked against the
current saved draft. GUI and web use equivalent bounded draft histories.

## Develop and verify

The frontend is React/TypeScript/Vite, TanStack Query/Table/Virtual, Radix controls,
PDF.js, ECharts, React Hook Form/Zod and Playwright. Python uses FastAPI/Pydantic,
SQLAlchemy/Alembic/SQLite and independent worker processes. Financial calculations
remain in the existing domain/application modules. Decimal summary/claim values
are transported as strings; original record observations retain their saved form.

```console
python -m pip install -c .github/ci-constraints.txt -e ".[dev,msg,web]"
python ci_pytest.py
npm ci --prefix frontend
npm run generate --prefix frontend
npm run format:check --prefix frontend
npm run build --prefix frontend
npm test --prefix frontend
```

Run frontend commands from an activated Python environment. `npm run generate`
regenerates committed OpenAPI/input types and requires that Python environment.
Playwright needs `npx playwright install chromium` from `frontend/`; choose
`EDF_BROWSER=firefox` or `webkit` and install that engine for the corresponding run.
Tests use an isolated synthetic server and never personal evidence. Container CI
also exercises an uploaded scanned bill through real OCR, HTTP corrections, all
report formats and byte-for-byte original preservation.

For live development, start `edf-web` and `npm run dev --prefix frontend`.
Vite proxies `/api` to port 8000; the production frontend is served by Python.
CI retains the 95% overall/100% protected-contract gates, checks every native
Python matrix leg, verifies browser workflows and builds amd64/arm64 images.

Back up the whole data directory/volume while the server is stopped, including
the SQLite catalogue and evidence directories. On restart, running jobs become
interrupted rather than successful. Queued jobs remain queued. Retry creates a
fresh run; browser reconnection resumes progress viewing, not an unimplemented
mid-PST processing checkpoint. Catalogue upgrades use Alembic; result.json's
portable evidence contract is separate from server metadata.
