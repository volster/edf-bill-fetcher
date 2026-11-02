# Documentation

Start with the [web branch README](../README.md) to launch the browser workspace
with Docker or from source, set up LAN access, and create your first evidence
pack. The [browser guide](WEB_UI.md) covers hosting and web workflows in detail;
the [desktop user guide](USER_GUIDE.md) describes the Qt controls. Desktop
downloads track `how-its-going` and may have different controls or features.

## Using the app

[Web accounts, sharing and database administration](WEB_ADMINISTRATION.md) covers
first login, workspace roles, account recovery, PostgreSQL upgrades and backups.

| Guide | What it covers |
| --- | --- |
| [User guide](USER_GUIDE.md) | Build a pack, inspect evidence, correct records in the app or Excel, and regenerate outputs. |
| [Browser / Docker workspace](WEB_UI.md) | Source and container hosting, LAN access, resumable folder/ZIP intake and browser review on `feature/web-ui`. |
| [Configuration](CONFIGURATION.md) | Preferences, CLI JSON, filters, duplicate handling, valuation dates and report sections. |
| [Understanding your outputs](OUTPUTS.md) | Which files to keep, where to find information, what figures mean and how source links work. |
| [Rebilling and claim review](REBILLING_LEDGER.md) | Bill relationships, dispute positions, consolidated claims, credit allocations and complaint packs. |
| [Troubleshooting](TROUBLESHOOTING.md) | Startup, OCR, extraction gaps, workbook imports and missing evidence. |
| [Opening on macOS](MACOS.md) | Apple silicon downloads and the first-launch approval. |
| [Command-line usage](CLI.md) | Extraction, saved reports, amended workbooks, record comparisons and diagnostics. |

## Suggested reading routes

- **First browser run:** [Docker or source quick start](../README.md#get-the-app) →
  [browser workflow](../README.md#build-your-first-evidence-pack) → [output meanings](OUTPUTS.md).
- **Desktop:** [source setup](BUILDING.md) or [downloads](../README.md#download-a-build) →
  [desktop user guide](USER_GUIDE.md).
- **Correct an existing pack:** [review workflow](USER_GUIDE.md#correct-records-in-excel) →
  [claim review](REBILLING_LEDGER.md) → [sharing checks](OUTPUTS.md#before-sharing-a-pack).
- **Maintain the code:** [development setup](BUILDING.md#development-and-checks) →
  [architecture](ARCHITECTURE.md) → [development recipes](DEVELOPMENT.md#where-to-make-a-change).
- **Investigate a failure:** [app troubleshooting](TROUBLESHOOTING.md) or
  [CI troubleshooting](CI.md), depending on where it happened.

## Developing and distributing

| Guide | What it covers |
| --- | --- |
| [Source setup and builds](BUILDING.md) | Python/native prerequisites, local builds and all three platforms through GitHub. |
| [Development](DEVELOPMENT.md) | Daily checks, useful regression tests and where to make changes. |
| [Architecture](ARCHITECTURE.md) | Evidence flow, accepted results, financial contracts, review drafts and publication. |
| [Coverage](COVERAGE.md) | Reproducing the CI measurement and choosing meaningful tests. |
| [CI troubleshooting](CI.md) | Matrix and release gates, diagnostic artifacts and past platform failures. |
| [Web reliability checks](WEB_RELIABILITY.md) | Repeatable synthetic upload, archive, recovery and concurrency checks. |

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

Each guide has a specific owner topic. Configuration defaults live in the
configuration reference; output file/status meanings live in the output guide;
financial and persistence contracts live in architecture. Historical review
counts stay in historical records. The current docs describe implemented behavior,
not proposed features.

Use neutral example names and synthetic inputs. Public guides should explain
product behavior without personal references, private local paths or quotations
from development conversations.
