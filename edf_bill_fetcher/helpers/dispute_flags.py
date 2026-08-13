"""Shared dispute-flag detection thresholds for compute_dispute_flags.

The heuristics (Ombudsman evidence criteria) historically hardcoded these
magic numbers inside the per-row loops of ``compute_dispute_flags`` — which
existed in two places (``processors/analysis.py`` and the legacy
``writers/_helpers.py`` re-export).  Centralising them here keeps the
business rules self-documenting and lets a threshold be tuned in one spot
without hunting through nested loops.
"""

from __future__ import annotations

# LARGE JUMP: >25% balance increase within 90 days.
LARGE_JUMP_PCT = 0.25
LARGE_JUMP_MAX_DAYS = 90
# Above 50% the jump is flagged HIGH instead of MEDIUM.
LARGE_JUMP_HIGH_PCT = 0.5

# BILLING GAP: >60 days without a bill.
BILLING_GAP_MIN_DAYS = 60
# Above 120 days the gap is flagged HIGH instead of MEDIUM.
BILLING_GAP_HIGH_DAYS = 120

# ESTIMATED RUN: 3+ consecutive estimated readings.
ESTIMATED_RUN_MIN = 3

# HIGH DAILY RATE: daily charge >2.5x the account's mean daily rate.
HIGH_DAILY_RATE_RATIO = 2.5
# Above 4x the ratio is flagged HIGH instead of MEDIUM.
HIGH_DAILY_RATE_HIGH_RATIO = 4

# BALANCE REDUCTION: balance fell by more than £500 (payment/credit).
BALANCE_REDUCTION_AMOUNT = 500.0

# RECONCILIATION MISMATCH: balance delta vs period charge; tolerated
# difference is max(10% of the period charge, £50) — above half the period
# charge it is flagged HIGH instead of MEDIUM.
RECON_PCT_TOLERANCE = 0.10
RECON_MIN_TOLERANCE = 50.0
RECON_HIGH_PCT = 0.5
