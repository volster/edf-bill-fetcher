# EDF Evidence Collector — Web Workspace

Bring your EDF bills, emails and account history together, check the figures
against the originals, and build an evidence pack for a billing complaint.

This is the **`feature/web-ui` branch**. It adds a browser workspace you can run
on your own computer or host on your network for someone else to use. Docker,
source installations, the desktop app and the command line use the same Python
analysis and review services.

Evidence is processed and stored on the machine running the app. Other computers
connect through their browsers; they do not need Python or OCR installed. No EDF
login or cloud service is needed.

## What you can do

- **Add a whole collection:** upload files, folders and ZIPs together. Automatic
  inspection identifies supported evidence and shows what needs attention.
- **Handle large uploads:** uploads use resumable chunks and verify the original
  bytes before finalising a collection. Folder paths are preserved.
- **Check the evidence:** search records, read extracted text, and compare fields
  with original PDF pages. Scanned PDFs use local OCR.
- **Correct and review:** amend values, exclude or reinstate records, record your
  reasons, and undo or redo saved draft changes.
- **Build your case:** inspect bill revisions, review dispute positions,
  reconcile claims and credits, and prepare a curated complaint pack.
- **Keep using Excel:** download the editable workbook, make corrections in its
  review tables, and upload it to regenerate the pack.
- **Follow the account:** explore balances, payments, readings, timelines and
  evidence gaps, then compare saved runs.
- **Choose your outputs:** publish Excel, HTML, PDF and Word, with verified source
  copies and a complete pack download.

The app highlights possible billing problems; findings and proposed reductions
need checking against the original evidence. They are not confirmed entitlements.

## Get the app

### Docker quick start

Install **Git** and **Docker with Docker Compose**. Docker Desktop includes
Compose on Windows and macOS. Then run:

```console
git clone --branch feature/web-ui https://github.com/volster/edf-bill-fetcher.git
cd edf-bill-fetcher
docker compose build
docker compose run --rm app --init-user reviewer
docker compose up -d
```

Replace `reviewer` with your chosen username. The account command prompts for a
password of at least **12 characters**. Open
**<http://127.0.0.1:8000>** and sign in with the username and password you chose.
Create the account once; subsequent launches only need `docker compose up -d`.

The image includes the compiled browser interface, Python dependencies, Outlook
readers and Tesseract with English OCR data. The first build downloads and
installs these dependencies; the host does not need Python or Node.

Useful commands:

```console
docker compose logs -f app
docker compose stop
docker compose up -d
```

Your uploads, runs, accounts and saved drafts live in the **`evidence-data` named
volume**. Replacing the container preserves that volume. Back it up before
updates; `docker compose down --volumes` removes the workspace.

### Run from source

Use this branch's checkout from the clone command above. Install **Node 24** and
either **Conda** or **64-bit Python 3.11**; Python 3.10–3.12 are supported.

For a first setup, the existing bootstrapper prepares the Python environment and
native prerequisites without launching the desktop:

```console
python bootstrap.py develop --install-system
```

Run that in a Conda-enabled terminal if you use Conda. Otherwise use
`py -3.11 bootstrap.py develop --install-system` on Windows, or
`python3.11 bootstrap.py develop --install-system` on macOS/Debian/Ubuntu.
Native package installation may request administrator access.

The bootstrapper reuses or creates the named **`edf-bill-fetcher` Conda
environment** when Conda is available, and falls back to **`.venv`** otherwise.
Activate the environment it selected:

| Environment              | Activation command                |
| ------------------------ | --------------------------------- |
| Conda                    | `conda activate edf-bill-fetcher` |
| venv, Windows PowerShell | `.\.venv\Scripts\Activate.ps1`    |
| venv, macOS/Linux        | `source .venv/bin/activate`       |

From that activated environment, install the web extra and build the interface:

```console
python -m pip install -c .github/ci-constraints.txt -e ".[web,msg]"
npm ci --prefix frontend
npm run build --prefix frontend
python main.py --doctor
edf-web --data web-data --init-user reviewer
edf-web --data web-data
```

Open **<http://127.0.0.1:8000>**. Later launches only need
`edf-web --data web-data` from the same environment. `python main.py --web`
accepts the same server options.

Node is needed to build the frontend, rather than for subsequent launches.
Tesseract and its English data are native requirements: pip alone does not
install them. For manual environment setup, other Linux distributions or setup
problems, see [source prerequisites and troubleshooting](docs/BUILDING.md).

## Let someone else use it on your network

Only the host machine needs the app and processing dependencies. Another person
opens its address in a browser and signs in.

**With Docker:** create a `.env` file beside `compose.yaml` containing:

```dotenv
EDF_WEB_BIND_ADDRESS=0.0.0.0
EDF_WEB_PORT=8000
```

Run `docker compose up -d` again to apply the settings.

**From source:** start the server with:

```console
edf-web --data web-data --host 0.0.0.0 --port 8000
```

On the other computer, open the host's LAN address, for example
**`http://192.168.1.20:8000`**. Use your host's actual address and allow the chosen
port through its firewall. `127.0.0.1` always refers to the computer you are using.

Accounts share the same workspace and evidence. To add another named account,
repeat the account-creation command with a different username and the same data
location. The [hosting guide](docs/WEB_UI.md) covers configured server folders,
read-only Docker mounts, HTTPS proxies and secure cookies.

## Build your first evidence pack

1. Sign in and open **Build a pack**. Give the pack a name.
2. Leave **All evidence · detect automatically** selected. Choose files or ZIPs,
   or switch **Upload selection** to a folder and its subfolders. Click
   **Upload and add**; repeat for any other sources.
3. Choose **Inspect everything**. Review the inventory, including unsupported
   files and anything needing an explicit import choice. Acknowledge unprocessed
   files before continuing with the recognised originals.
4. Choose analysis settings and report formats as needed, then click
   **Create evidence pack**. Progress continues on the server if you leave the
   page; reopen the run from **Runs & evidence**.
5. Check the run status and review notes. Open a record beside its original PDF,
   correct or exclude it with a reason, and choose **Apply to draft**.
6. Use **Preview changes**, then **Create reviewed pack** to publish a new
   version. Download the complete pack from **Outputs**.

Applying an edit saves a draft. The accepted totals stay pinned until you preview
or publish your changes. Earlier runs remain available, and stale edits from
another tab are rejected.

For Excel review, download the workbook, edit the designated **Review &
Corrections** tables, and save it. In **Build a pack**, choose **Regenerate amended
Excel workbook** as the workflow and upload that file. Editing a summary sheet or
Word report alone does not update the saved analysis.

If an upload is interrupted, select the same originals and choose **Resume
pending upload** in the same browser tab. Changed or incomplete selections are
rejected. Keep the originals unchanged while uploading and processing.

Current limits include **100,000 files per upload collection** and **64 GiB per
uploaded file**. ZIP expansion has separate size and nesting limits. See the
[upload and archive guide](docs/WEB_UI.md#add-everything-then-inspect) for details.

## Supported evidence

| Input                                              | What it supplies                                                                                             |
| -------------------------------------------------- | ------------------------------------------------------------------------------------------------------------ |
| EDF PDF bills and credit notes                     | Older layouts, T-series and KI/KCR invoices, with charges, balances, periods and other fields where present. |
| Merged and scanned PDFs                            | Separate invoice records with original page references; English OCR for pages without useful text.           |
| Outlook PST/OST archives                           | Relevant email bodies and supported attachments, with archive occurrence information.                        |
| EML/MSG messages                                   | Individual emails and supported attachments.                                                                 |
| EDF MyAccount HTM/HTML and supported PDF histories | Account transactions, payments and balance observations.                                                     |
| Supported EDF SAP dump PDFs                        | Contract, meter-reading and financial history for reconciliation.                                            |
| ZIPs and folder trees                              | Automatic inspection of their contents, preserving source and archive-member provenance.                     |
| Saved analyses and review workbooks                | Explicit import or regeneration through their own workflows.                                                 |

Unknown files remain visible in the inventory. Previous output folders are
skipped, and ZIPs containing `result.json` need an import choice so their reports
are not counted as new bills. Sources, duplicate handling and missing information
use the same rules as the desktop and CLI.

## Keep and share the whole pack

A published pack includes an editable workbook, `result.json`, `run.json`,
requested document reports and available verified source copies. It can also
include a focused complaint draft and its supporting evidence.

Keep the complete download together so links to originals remain usable. Opening
an accepted analysis preserves its saved calculations; re-extract the originals
to apply newer parsers. Excluding a record does not redact a shared PDF: included
source files remain whole.

Stop the server before backing up the entire `web-data` directory or Docker
volume, including the catalogue and evidence files. Queued jobs and applied
drafts persist across restarts. A running job becomes interrupted and needs a
fresh run; reconnecting to progress does not resume midway through a PST scan.

See [output meanings](docs/OUTPUTS.md), [claim review](docs/REBILLING_LEDGER.md)
and the [hosting/backup guide](docs/WEB_UI.md) for the details.

## Desktop and command line

The Qt desktop and CLI remain available from this checkout's Python environment:

```console
python main.py
python main.py --extract --auto "evidence" --auto "more.zip" -o "evidence.xlsx"
```

The [desktop user guide](docs/USER_GUIDE.md) and [CLI reference](docs/CLI.md)
cover their controls and commands. Saved analyses, review workbooks and complete
evidence packs let you carry work between interfaces.

### Download a build

Prebuilt desktop downloads track **`how-its-going`**. Open
[Build Executables on GitHub Actions](https://github.com/volster/edf-bill-fetcher/actions/workflows/matrix_build.yml),
choose a successful run for that branch, and download its
`EDF_Evidence_Collector-…-ci-validated` artifact. GitHub may require sign-in.
Apple silicon users should follow the [macOS first-launch guide](docs/MACOS.md).

Use this README's Docker or source instructions for the browser workspace.

## Development and further reading

The frontend uses React, TypeScript and Vite; FastAPI serves it and calls the
shared Python application services. SQLite stores server metadata, files hold the
evidence, and separate worker processes handle queued jobs.

See [web development and checks](docs/WEB_UI.md#develop-and-verify) for setup,
generated API contracts, browser tests and the development server. CI checks
Python 3.10–3.12 across Windows, macOS and Linux, browser workflows in Chromium,
Firefox and WebKit, and Docker builds for amd64 and arm64. It retains the 95%
overall and 100% protected-contract coverage gates.

| Guide                                           | Use it for                                                                        |
| ----------------------------------------------- | --------------------------------------------------------------------------------- |
| [Browser workspace and hosting](docs/WEB_UI.md) | Deployment options, upload limits, source roots, backups and browser development. |
| [Documentation index](docs/README.md)           | Find the remaining guides by task.                                                |
| [Configuration](docs/CONFIGURATION.md)          | Filters, duplicates, valuation dates and report sections.                         |
| [Understanding your outputs](docs/OUTPUTS.md)   | Workbook sheets, report contents, source links and financial meanings.            |
| [Architecture](docs/ARCHITECTURE.md)            | Shared services, accepted results, review contracts and persistence.              |
| [Development](docs/DEVELOPMENT.md)              | Python development practices and regression checks.                               |

[Changes](CHANGELOG.md) · [Report a bug](https://github.com/volster/edf-bill-fetcher/issues) · [MIT licence](LICENSE)
