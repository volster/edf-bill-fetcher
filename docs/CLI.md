# Command-line usage

[Documentation index](README.md) · [Source setup](BUILDING.md)

Run commands from the checkout using the selected project environment. With
Conda, run `conda activate edf-bill-fetcher` first, or prefix commands with
`conda run --no-capture-output -n edf-bill-fetcher`. With venv, use
`.venv/bin/python` on macOS/Linux or `.\.venv\Scripts\python.exe` on Windows.
The examples use `python` for that interpreter. The installed `edf-collector`
command accepts the same arguments as `python main.py`.

Put the operation first, such as `--extract` or `--pdf-report`. For help, append
`--help` to that operation; running `main.py` without an operation opens the GUI.
Quote paths containing spaces.

## Extract evidence

```console
python main.py --extract --pdf-dir "bills" --output "output/evidence.xlsx"
```

Create the destination parent folder first. Extraction writes the requested
workbook and creates a unique adjacent source-bundle folder containing
`result.json` and `run.json`. Unlike desktop runs, the workbook destination is
explicit: choose a new name if you want to preserve an existing workbook.

Combine sources in one command when needed:

```console
python main.py --extract --pdf-dir "bills" --pst "archive.pst" --htm "history.htm" --eml-dir "emails" --msg-dir "messages" --output "output/combined.xlsx"
```

`--pst` accepts PST or OST. `--eml-dir` and `--msg-dir` scan their respective
message formats. Native OCR and archive-reading prerequisites still apply in
headless mode.

Useful options:

| Option | Purpose |
| --- | --- |
| `--config config.json` | Load extraction settings. |
| `--acc-filter ACCOUNT` | Restrict extraction to an account. |
| `--domain-filter edfenergy.com` | Set sender domains for PST filtering. |
| `--min-amount NUMBER` | Set the amount threshold. |
| `--no-dedup` | Keep duplicate records in the prepared output. |
| `--strict` | Return a nonzero status if no usable billing records are produced. |
| `--records-json records.json` | Also save a version 2 compatibility records export. |
| `--save-engine-json engine.json` | Also save version 2 compatibility engine state. |

Use `python main.py --extract --help` for all flags. Without `--strict`, a clean
empty extraction can exit successfully; inspect run status and review notes as
well as the process exit code.

## Render a saved analysis

Prefer the run's `result.json` to preserve its accepted calculations:

```console
python main.py --pdf-report -i "run/result.json" -o "run/report.pdf"
python main.py --docx-report -i "run/result.json" -o "run/report.docx"
python main.py --html-report -i "run/result.json" -o "run/report.html"
```

Each command also writes a result and run manifest beside its report, for example
`report.pdf.result.json` and `report.pdf.run.json`. Keep reports with their source
bundle for portable source links. If a render fails, inspect its run manifest;
completed analysis and earlier artifacts can remain available.

`-i` is also named `--records`. Legacy bare record lists and version 1/2 JSON
exports are accepted with migration limitations. `--engine-data engine.json`
can supply legacy metadata; omit it for a v3 result, which already contains it.
Pickles are not accepted as report interchange files.

Use `--config config.json` for report configuration. A changed analytical setting
creates a new result revision; the original accepted file remains unchanged.
The section keys are defined in
[`models/report_sections.py`](../edf_bill_fetcher/models/report_sections.py),
and configuration keys in [`models/config.py`](../edf_bill_fetcher/models/config.py).

## Import an amended workbook

```console
python main.py --review-workbook --workbook "amended.xlsx" --output-dir "reviewed-packs"
python main.py --review-workbook -i "amended.xlsx" -o "reviewed-packs" --formats xlsx,html
```

This validates the review inputs, recalculates using the preserved analysis
context and creates a new pack. An updated workbook is always included. Supported
format names are `xlsx`, `pdf`, `docx` and `html`; the default includes all four.
Follow the [Excel review rules](USER_GUIDE.md#correct-records-in-excel), including
keeping source copies accessible. `Claim_Schedule.xlsx` is reference-only and
cannot be imported here.

## Compare records exports

```console
python main.py --diff old_records.json new_records.json --diff-output diff.xlsx
```

This prints added, removed and changed records, with an optional three-sheet
workbook. It compares record exports; it is not a substitute for the saved review
history or claim reconciliation.

## Diagnose an installation

```console
python main.py --doctor --output diagnostics.json
```

The diagnostic checks installed capabilities and identifies the build. A nonzero
status means a capability is missing. See [troubleshooting](TROUBLESHOOTING.md)
for interpreting failures.
