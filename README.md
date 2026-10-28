# EDF Energy Billing Evidence Collector

Turn a folder of EDF bills and emails into an organised evidence pack you can
check, correct and use to prepare a billing complaint.

The app brings together invoices, payments, meter readings and account history.
You can compare records with their original PDFs, trace revised bills and credits,
and review the figures either in the app or in Excel. When you make corrections,
it creates a new pack and keeps a record of what changed.

## What you can do

- **Gather the evidence:** import PDF bills, Outlook PST/OST archives, EML/MSG
  emails and EDF account-history exports. Scanned PDF pages use local OCR.
- **Understand the account:** follow balances, payments, readings and bill
  revisions, with missing information and possible billing problems highlighted.
- **Check and correct it:** inspect original pages, fix extracted values or exclude
  records, and carry your reasons into the next version of the pack.
- **Prepare a complaint:** record disputed charges, reconcile overlapping claims
  and credits, and produce a focused complaint draft with supporting originals.
- **Choose your outputs:** an Excel workbook for review, plus HTML, PDF and Word
  reports for reading and sharing.

Processing happens on your computer; no EDF login or cloud service is needed.
The app was built for a personal EDF dispute and understands the EDF document
layouts covered by its parsers. Check findings against the originals: a flagged
charge or proposed reduction is not a confirmed entitlement to a refund.

## Get the app

### Download a build

Open [Build Executables on GitHub Actions](https://github.com/volster/edf-bill-fetcher/actions/workflows/matrix_build.yml),
choose a successful run for `how-its-going`, and download the
`EDF_Evidence_Collector-…-ci-validated` artifact for your computer from the bottom
of the run page. GitHub may ask you to sign in to download artifacts.

| Your computer | What to open after extracting the download |
| --- | --- |
| Windows | `EDF_Evidence_Collector.exe` |
| Mac with Apple silicon | `EDF_Evidence_Collector.app` — follow the [first-launch instructions](docs/MACOS.md) |
| Linux | `EDF_Evidence_Collector` |

These builds include the OCR engine and English language data. You do not need
Python. The Mac build is ad-hoc signed and needs a one-time **Open Anyway**
approval; it is not Apple-notarized. Build artifacts can expire; see
[building from source](docs/BUILDING.md) if a download is no longer available.

### Run from source

Install Git and **64-bit Python 3.11**, then clone the working branch:

```console
git clone --branch how-its-going https://github.com/volster/edf-bill-fetcher.git
cd edf-bill-fetcher
```

On Windows:

```powershell
py -3.11 bootstrap.py run --install-system
```

On macOS or Debian/Ubuntu Linux:

```console
python3.11 bootstrap.py run --install-system
```

Setup creates a local Python environment, installs dependencies, checks OCR and
launches the app. Native package installation may request administrator access.
For other Linux distributions, development setup or executable builds, use the
[source and build guide](docs/BUILDING.md). Installing Python requirements alone
does not install Tesseract.

## Your first evidence pack

1. Open **Build evidence pack** and choose your PDF folder, email sources or
   account-history file. You can combine sources in one run.
2. Choose where to save it. Keep **Include original source files** selected so
   the pack can link records back to the evidence. Use **Report settings…** if
   you also want HTML, PDF or Word reports.
3. Click **Create evidence pack**. When it finishes, read any review notes and
   open the workbook's **Start here** sheet for the overview and priorities.
4. Check flagged records against their originals in **Review evidence**, or
   edit the workbook's **Review & Corrections** sheet. Create a reviewed pack
   to apply your changes and regenerate the outputs.

Keep the whole output folder together, including `result.json` and
`evidence_files/`. Each desktop run gets a new folder, so earlier packs remain
available. The [user guide](docs/USER_GUIDE.md) walks through both review routes.

## Where to go next

| I want to… | Guide |
| --- | --- |
| Correct records in the app or Excel | [Using and reviewing evidence](docs/USER_GUIDE.md) |
| Understand the workbook, reports and saved files | [Understanding your outputs](docs/OUTPUTS.md) |
| Review rebills, credits and a proposed complaint | [Rebilling and claim review](docs/REBILLING_LEDGER.md) |
| Fix a startup, extraction or import problem | [Troubleshooting](docs/TROUBLESHOOTING.md) |
| Automate extraction or report generation | [Command-line usage](docs/CLI.md) |
| Work on the code | [Development guide](docs/DEVELOPMENT.md) |

The [documentation index](docs/README.md) also covers architecture, testing,
builds and historical review notes. See the [changelog](CHANGELOG.md) for changes.

Found a bug? Include your OS, **About this build** information and the exact
error in a [GitHub issue](https://github.com/volster/edf-bill-fetcher/issues).
Use a synthetic or redacted example; bills, workbooks and saved analyses can
contain personal and financial information.

[MIT licence](LICENSE).
