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

Install Git and either Conda or **64-bit Python 3.11**, then clone the working branch:

```console
git clone --branch how-its-going https://github.com/volster/edf-bill-fetcher.git
cd edf-bill-fetcher
```

**With Conda:** open a Conda-enabled terminal (such as an Anaconda or Miniforge
prompt) and run:

```console
python bootstrap.py run --install-system
```

**Without Conda:** use `py -3.11 bootstrap.py run --install-system` on Windows,
or `python3.11 bootstrap.py run --install-system` on macOS/Debian/Ubuntu.

The bootstrapper first looks for Conda. If found, it reuses the named
`edf-bill-fetcher` environment or creates it with Python 3.11. Otherwise, it
reuses or creates `.venv` in the checkout. It then installs the app, checks native
capabilities including OCR, and launches it. It does not install the app into
Conda's `base` environment. Native package installation may request administrator
access.

Use `--conda-env NAME` for an existing project environment with another name, or
`--venv .venv` to explicitly select venv. **`--check` only checks** the selected
existing environment. See the [setup and build guide](docs/BUILDING.md) for
platform prerequisites and troubleshooting.

After setup, activate the selected environment and use `python main.py` or
`edf-collector` normally. With the default Conda environment:

```console
conda activate edf-bill-fetcher
python main.py
```

With venv, activation is optional: run `.\.venv\Scripts\python.exe main.py` on
Windows or `.venv/bin/python main.py` on macOS/Linux.

**Prefer a manual install?** In your chosen Python 3.10–3.12 environment:

```console
python -m pip install -c .github/ci-constraints.txt -e ".[msg]"
python main.py --doctor
python main.py
```

Install the [native prerequisites](docs/BUILDING.md#native-prerequisites) first.
Pip installs the Python libraries, including Qt and PDF readers, but does not
install Tesseract or its English data. The bootstrapper automates environment
setup and checks; it is not required to launch the app.

## Supported evidence

| Source | What the app looks for |
| --- | --- |
| PDF bills and credit notes | Older EDF layouts, T-series and newer KI/KCR layouts; invoice dates, periods, charges, balances, readings and tariff information where present. |
| Merged and scanned PDFs | Separate invoice records with original page references; local English OCR for pages lacking useful text. |
| PST/OST archives | Relevant email bodies and supported PDF attachments, with archive occurrence information. |
| EML/MSG files | Individual email messages and supported attachments; MSG support is included by the bootstrapper. |
| MyAccount history | HTM/HTML exports and supported printed PDF histories of charges, payments and balances. |
| EDF SAP exports | Supported contract-history, meter-reading and financial-transaction PDF dumps, with reconciliation against invoice evidence. |

Sources can be combined. Duplicate handling uses account/meter identity and
supporting document or transaction relationships; matching dates and amounts
alone do not prove a duplicate. Source occurrences remain available for review.
Unreadable messages, unsupported layouts and incomplete extraction are reported
as gaps rather than assumed to contain no relevant evidence.

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

## Review, correct and regenerate

| In the app | In Excel |
| --- | --- |
| Open `result.json` in **Review evidence** and compare a record with its original page. | Open the main workbook and follow **Start here** and **Workbook guide**. |
| Choose **Edit selected record**, correct fields or exclude the record, and record a reason. | Edit pale green inputs in **Review & Corrections**, keeping row IDs and reference fields intact. |
| Use **Preview changes**, then **Create reviewed pack**. | Save the amended workbook, select it in **Generate reports**, and create a new pack. |

Both routes retain original observations and record your changes. Editing an
Excel summary or Word report alone does not update the saved analysis. Opening
an old `result.json` preserves its accepted calculations; re-extract the originals
to apply parser improvements. See the [full review workflow](docs/USER_GUIDE.md).

For a disputed rebill, use **Bill revisions** to inspect replacements and credits,
then **Dispute ledger → Reconcile claims…** to consolidate your proposed challenge.
Individual revised invoices can repeat the same demand. The app keeps requested
reductions separate from observed account balances and withholds a combined claim
while overlaps or credit allocations remain unresolved. The
[claim guide](docs/REBILLING_LEDGER.md) explains the review fields and complaint draft.

## What you get

Each extraction pack includes a workbook and saved analysis. Selected reports and
source copies sit alongside them:

| Output | Use it for |
| --- | --- |
| Main Excel workbook | Overview, review priorities, editable corrections, evidence, payment/readings analysis and dispute schedules. |
| HTML / PDF / Word reports | Reading or sharing selected findings, timelines, methodology and evidence references. HTML works offline. |
| `result.json` | Reopening the accepted analysis with its settings, valuation date, source identities and calculations. |
| `run.json` | Checking completion, review requirements, failures and the artifacts produced. |
| `evidence_files/`, manifest and Word index | Following records to saved originals, including original filenames, page ranges and hashes where available. |
| Optional complaint folder and ZIP | A focused draft in three formats, reference claim schedule and relevant verified source files, generated during claim review. |

The workbook also includes conditional SAP reconciliation and analytical views
for billing anomalies, tariff comparisons, back-billing, data quality and
forecasts. Some require enough compatible observations; missing evidence stays
explicit. The [output guide](docs/OUTPUTS.md) explains the sheets and financial
meanings in detail.

Source copies are linked rather than embedded in Excel. Keep the whole pack and
any earlier packs it references. Excluding a record is **not PDF redaction**:
shared PDFs remain whole and saved analysis retains the original observations.

## Automation and development

The command line can extract evidence, render saved analyses, import corrected
workbooks and compare record exports. For example, from your active environment:

```console
python main.py --extract --pdf-dir "bills" --output "evidence.xlsx"
python main.py --html-report -i "path/to/run/result.json" -o "report.html"
python main.py --review-workbook -i "amended.xlsx" -o "reviewed-packs"
```

Use the actual saved-result path printed after extraction. See the
[CLI guide](docs/CLI.md) for source options, output locations and help commands.

`python bootstrap.py develop` adds tests, lint/type checks and packaging tools to
the selected environment. `python bootstrap.py build` creates an executable for
the current OS; `python build_all.py` requests all three native builds through
GitHub. See [building](docs/BUILDING.md) and [development](docs/DEVELOPMENT.md).
CI covers Python 3.10–3.12 on Windows, macOS and Linux; build/check results belong
to the revision shown in GitHub Actions and **About this build**.

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
