"""
payroll_types.py

Data model and types for the payroll calculation feature (Tab 3).

Calculates RETSS (9%) and PARTSS (25% or 12.5%) from BRUTSS.
NETPAI is read directly from the input file.

Two categories:
  - Titulaires (confirmed): PARTSS = 25%
  - Non Titulaires (non-confirmed / handicap): PARTSS = 12.5%

All monetary values are stored as integer cents internally to avoid
floating-point rounding errors (e.g. 1234.56 → 123456 cents).
"""

from typing import NamedTuple


# ─────────────────────────────────────────────────────────────────────────────
# Rate constants
# ─────────────────────────────────────────────────────────────────────────────

RETSS_RATE_PERCENT = 9          # 9% — same for both categories
PARTSS_CONFIRMED_PERCENT = 25   # 25% — confirmed employees
PARTSS_NON_CONFIRMED_PERMILLE = 125  # 12.5% stored as 125‰ to stay integer

# ─────────────────────────────────────────────────────────────────────────────
# Non-confirmed employee NUMCPTs (normalized — leading zeros stripped)
# ─────────────────────────────────────────────────────────────────────────────
# Raw values:
#   007999990005329601  →  7999990005329601
#   007999990005674483  →  7999990005674483
#   007999990014277776  →  7999990014277776
#   007999990014363226  →  7999990014363226
#   007999990005298833  →  7999990005298833

NON_CONFIRMED_NUMCPTS: frozenset = frozenset({
    "7999990005329601",
    "7999990005674483",
    "7999990014277776",
    "7999990014363226",
    "7999990005298833",
})


# ─────────────────────────────────────────────────────────────────────────────
# Data model
# ─────────────────────────────────────────────────────────────────────────────

class PayrollEmployee(NamedTuple):
    """One parsed employee from the input file."""
    numcpt_raw: str     # Original NUMCPT (preserves leading zeros)
    numcpt_norm: str    # Normalized NUMCPT (leading zeros stripped)
    nom: str
    prenom: str
    brutss_cents: int   # BRUTSS in integer cents
    netpai_cents: int   # NETPAI in integer cents (read from file)


class PayrollResult(NamedTuple):
    """One employee after payroll calculation."""
    numcpt_raw: str
    nom: str
    prenom: str
    is_confirmed: bool  # True = titulaire, False = non titulaire
    brutss_cents: int
    retss_cents: int    # BRUTSS × 9%
    partss_cents: int   # BRUTSS × 25% (confirmed) or 12.5% (non-confirmed)
    netpai_cents: int   # Read from file


class PayrollStats(NamedTuple):
    """Summary statistics for the payroll calculation."""
    total_employees: int
    confirmed_count: int
    non_confirmed_count: int
    # Confirmed totals
    confirmed_brutss_cents: int
    confirmed_retss_cents: int
    confirmed_partss_cents: int
    confirmed_netpai_cents: int
    # Non-confirmed totals
    non_confirmed_brutss_cents: int
    non_confirmed_retss_cents: int
    non_confirmed_partss_cents: int
    non_confirmed_netpai_cents: int
    # Grand totals
    grand_brutss_cents: int
    grand_retss_cents: int
    grand_partss_cents: int
    grand_netpai_cents: int


class PayrollOutput(NamedTuple):
    """Full output of the payroll calculation, consumed by the exporter."""
    confirmed_rows: list    # list[PayrollResult] — titulaires
    non_confirmed_rows: list  # list[PayrollResult] — non titulaires
    stats: PayrollStats
