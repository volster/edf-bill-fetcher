# Handoff: EDF Bill Fetcher

## Current state

- Branch: `how-its-going`
- Latest committed review-fix baseline: `2a5d147`
- The working tree contains one uncommitted batch of 12 targeted fixes from the adversarial-review todo list. These changes are intentionally not yet mixed with the next architectural refactor.
- Do not push from a fresh handoff unless the user explicitly requests it.

## Uncommitted work to commit

The current batch covers:

- `C-2`: CLI `--strict` makes clean-empty extraction exit 0 by default and 1 in strict mode.
- `C-3/L-15`: PDF/DOCX numeric formatting defaults to two decimals.
- `C-6`: delta amounts for `LARGE JUMP` and `BALANCE REDUCTION` dispute flags.
- `M-1`: safe numeric coercion for string amounts such as `N/A`.
- `M-3`: evidence hyperlinks indexed from the full evidence frame.
- `M-5`: derived reconciliation row offset.
- `M-7`: sheet reordering on the single-row early exit.
- `M-13`: HTM processing failures surface through the UI/stderr.
- `L-10`: literal glob handling in sequential output naming.
- `L-11`: OFGEM carry-forward returned separately instead of a sentinel dict key.
- `L-12`: all-NaN OFGEM quarters render as `N/A`.
- `L-14`: documented/warned day-first fallback assumption.

Before committing, inspect the diff and group implementation with its direct tests. The files currently marked modified by Git are the complete batch.

## Completed implementation already committed

The branch already contains the back-billing/SAP correction arc and its verification, including:

- SLC 21BA per-consumption-day legal eligibility using bill Date versus Period From/To.
- Period Charge as the canonical charge, with Amount as documented fallback.
- Unlawful-charge proration and recursive supersession handling.
- SAP matcher Posting Date + Clearing Date axis and cluster-unmatched internal-mechanism tagging.
- Production wiring for domination and cluster-unmatched handling.
- macOS ad-hoc signing fix and Windows headless Tk test stabilization.
- ConfigDict/mypy 2.x drift fixes.

The main evidence and design documents are:

- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/COVERAGE.md`
- `docs/superpowers/refs/backbilling-legal-definition.md` (internal/ignored)
- `docs/superpowers/specs/2026-08-07-backbilling-sap-corrections-design.md` (internal/ignored)
- `docs/superpowers/specs/2026-08-10-adversarial-code-review-design.md` (internal/ignored)
- `.omo/plans/2026-08-07-backbilling-sap-corrections.md` (internal/ignored)
- `.omo/plans/2026-08-10-adversarial-code-review.md` (internal/ignored)
- `scratch/docs/non-developer-technical/README.md`
- `scratch/reviews/final-report.md`

## Remaining agreed work

1. Commit the outstanding 12-fix batch.
2. `C-1/C-4`: extract duplicated production functions into neutral shared modules; do not merely delete one copy and repoint it.
3. `C-7`: create one shared payment-figure helper used by Excel, PDF, and DOCX; prefer Period Charge and fall back to Amount.
4. Arch #1: complete the duplicate-function sweep.
5. Arch #2: introduce a typed `BillingRecord` schema.
6. Arch #3: introduce shared `report_models.py` for report calculations.
7. Arch #4: introduce a shared sheet factory for banner/header/data row layout.
8. Arch #5: move OFGEM caps to a data file with freshness/carry-forward behavior.
9. Arch #6: add PDF slicer edge-case coverage.

The user explicitly prefers one subagent stage at a time because NVIDIA is rate-limited. Use CodeGraph for code exploration before raw searches. Keep implementation delegated rather than editing directly in the main session when possible.

## Verification commands

Use the dedicated Conda environment if available:

```bash
conda run -n edf-bill-fetcher ruff check .
conda run -n edf-bill-fetcher ruff format --check .
conda run -n edf-bill-fetcher mypy . --exclude scratch/
conda run -n edf-bill-fetcher pytest --cov=. --cov-report=xml
```

The repository has a root `conftest.py` that boots a virtual display on Linux when needed; an `xvfb-run -a` wrapper remains compatible.

## Configuration handoff

The NVIDIA model context limits were raised in `/home/matthew/.config/opencode/opencode.json` to `202752` for the configured NVIDIA models. Restart OpenCode after configuration changes so the new limits load.
