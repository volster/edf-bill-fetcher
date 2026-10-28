# Troubleshooting

[Documentation index](README.md) · [Source setup](BUILDING.md) · [CI failures](CI.md)

Start with the exact error and **About this build** information. Check **Details**
after a run and the workbook's **Parse Errors** and **Start here** review priorities.
A completed run can still need review; a failed run can still contain useful
saved artifacts.

## The app will not open

On Apple silicon, follow the [Mac first-launch instructions](MACOS.md). The
initial Gatekeeper alert does not contain the Windows-style continue button;
approval is in **System Settings → Privacy & Security** after a blocked launch.

For a source installation, use the Python interpreter in the project's `.venv`,
not a different system Python. Run:

```console
python main.py --doctor --output diagnostics.json
```

Use `.\.venv\Scripts\python.exe` on Windows or `.venv/bin/python` on macOS/Linux
in place of `python`. The [bootstrap check](BUILDING.md#development-and-checks)
can also identify missing prerequisites. A desktop display is required for normal
GUI use; CLI commands work headlessly. On Debian/Ubuntu, Qt also needs
`libxcb-cursor0`, `libxkbcommon-x11-0` and `libegl1`.

## Scanned bills produce little or no text

Source installations need the native Tesseract executable **and English language
data**. Pip alone does not install them. Use the [setup guide](BUILDING.md);
packaged builds include both. Refresh the desktop's readiness checks after a repair.

Inspect the original page and any OCR notes. OCR cannot repair an unreadable PDF
or extract an attachment the archive reader could not access. It also does not
replace every substantial existing text layer, even if that layer is inaccurate.
Correct misread values through the review workflow; re-extract if a parser or OCR
installation has changed.

## A PST/OST message reports an attachment error

Errors such as `pypff_message_get_number_of_attachments` or
`invalid local descriptors node` come from reading the archive's structure.
Other readable messages may still be processed, but the failed attachment must
be treated as an evidence gap. OCR cannot recover bytes the reader did not obtain.

If the message opens in Outlook, save its relevant PDF attachments separately and
include them through **PDF bills**. Keep the original archive and the error log.
Do not assume a repeated parser error establishes that every message is corrupt.

## Records or attachments seem to be missing

Check account/sender filters, minimum amounts, **Filtered (Below Min)**,
**Duplicate Entries** and **Parse Errors**. Duplicate source occurrences can be
retained under one prepared record. The app skips recognised earlier output
folders during recursive scans to avoid ingesting its own exports.

Unsupported PNG or TXT attachments can be signatures or incidental email content;
the message alone does not mean a bill failed. If one contains relevant evidence,
inspect it separately. For SAP short/long-row warnings, inspect the source row
and reconciliation gaps rather than assuming all fields were parsed correctly.

## Excel changes do not appear in the report

Edit the designated inputs in **Review & Corrections**, or the dispute/claim
input tables. Save and close the amended workbook, then select that `.xlsx` in
**Generate reports**. Opening the old `result.json` reuses the old accepted
analysis. Editing summary/reference tabs or a Word report does not update the
underlying records.

If import is rejected, use the named row/field in the error. Common causes are
formulas, missing reasons, invalid dates, altered reference cells or deleted IDs.
Restore reference cells from an untouched copy; do not remove metadata to bypass
validation. See the [review instructions](USER_GUIDE.md#correct-records-in-excel).

A workbook open in Excel can also prevent the app from writing to its path on
Windows. Close it or choose a new destination.

## Source links are broken or a source has changed

Keep the entire pack together and check `evidence_manifest.json` for the mapping
between original and saved filenames. A copied PDF can have an invoice-based name;
merged PDFs can serve several records. Match the recorded identity and hash, not
just the basename.

Restore missing files from an intact copy of the original pack. If the available
source really has changed, re-extract it as new evidence. Do not edit hashes to
make an unverified replacement appear valid. Excluded sources may need an earlier
pack when reinstated.

## Claims say “needs review” or show no combined total

A changed source can make a previous position, claim group or overlap decision
stale. Recheck and save the affected decisions. Ungrouped challenges, overlapping
claims and invalid credit allocations can also withhold a total. See
[claim reconciliation](REBILLING_LEDGER.md#consolidate-claims-and-account-for-credits).
Withdrawing a group does not automatically withdraw its individual positions.

## Reporting a bug

Include the OS, source revision from **About this build**, action you took,
exact error and whether it occurs in a packaged build or source installation.
Diagnostics omit selected evidence/account details, but native error messages
can contain local paths. Check before sharing. Use synthetic or redacted samples
in public issues; never upload an entire private evidence pack just to report a bug.
