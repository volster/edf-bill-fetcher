# Documentation

Start with the [project README](../README.md) to download the app or create your
first evidence pack. The guides below describe the current working branch;
older downloads may have different controls or fewer review features.

## Using the app

| Guide | What it covers |
| --- | --- |
| [User guide](USER_GUIDE.md) | Build a pack, inspect evidence, correct records in the app or Excel, and regenerate outputs. |
| [Understanding your outputs](OUTPUTS.md) | Which files to keep, where to find information, what figures mean and how source links work. |
| [Rebilling and claim review](REBILLING_LEDGER.md) | Bill relationships, dispute positions, consolidated claims, credit allocations and complaint packs. |
| [Troubleshooting](TROUBLESHOOTING.md) | Startup, OCR, extraction gaps, workbook imports and missing evidence. |
| [Opening on macOS](MACOS.md) | Apple silicon downloads and the first-launch approval. |
| [Command-line usage](CLI.md) | Extraction, saved reports, amended workbooks, record comparisons and diagnostics. |

## Developing and distributing

| Guide | What it covers |
| --- | --- |
| [Source setup and builds](BUILDING.md) | Python/native prerequisites, local builds and all three platforms through GitHub. |
| [Development](DEVELOPMENT.md) | Daily checks, useful regression tests and where to make changes. |
| [Architecture](ARCHITECTURE.md) | Evidence flow, accepted results, financial contracts, review drafts and publication. |
| [Coverage](COVERAGE.md) | Reproducing the CI measurement and choosing meaningful tests. |
| [CI troubleshooting](CI.md) | Matrix and release gates, diagnostic artifacts and past platform failures. |

## History

[CHANGELOG.md](../CHANGELOG.md) records product changes.
[Application review fixes](REVIEW_FIXES.md) records an earlier R01–R20 review
and its validation at the time. The [August 2026 coverage baseline](coverage/baseline-2026-08-02.txt)
is also historical; neither is a statement of current test results.

## Keeping these guides useful

Keep the root README focused on choosing, installing and starting the app.
Put detailed user tasks here, implementation contracts in the architecture guide,
and setup commands in the build guide. Link to the owning guide instead of
copying its instructions into several places. Update the relevant guide when
controls, input formats or output behavior change; retain old validation claims
only when they are clearly labelled as history.
