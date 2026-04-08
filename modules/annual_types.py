"""
annual_types.py

Data model and types for the annual declaration feature (Tab 4).

The user uploads 4 trimestrial output files (one per quarter).
Each file has a BRUTSS_TOTAL column per employee.
The annual declaration sums BRUTSS_TOTAL across all 4 quarters.

All monetary values are stored as integer cents internally to avoid
floating-point rounding errors (e.g. 1234.56 → 123456 cents).
"""

import os
from typing import NamedTuple


class EmployeeKey(NamedTuple):
    """Matching key for employees across quarterly files (NUMCPT only)."""
    numcpt: str  # Normalized NUMCPT (leading zeros stripped)


class QuarterlyEntry(NamedTuple):
    """One employee's BRUTSS_TOTAL from a single quarterly file (in cents)."""
    numcpt_raw: str   # Original NUMCPT (preserves leading zeros for output)
    nom: str
    prenom: str
    numss: str        # NUMSS, empty string if absent
    adm: str          # ADM code, empty string if absent
    datnais: str      # Date de naissance, empty string if absent
    nbrtrav: str      # Nombre de jours travaillés, empty string if absent
    datent: str       # Date d'entrée, empty string if absent
    datsor: str       # Date de sortie, empty string if absent
    brutss_total_cents: int  # BRUTSS_TOTAL in integer cents


class MissingEmployee(NamedTuple):
    """An employee absent from a specific quarterly file."""
    numcpt_raw: str
    nom: str
    prenom: str


class AnnualRow(NamedTuple):
    """One row in the final annual output."""
    numcpt_raw: str
    nom: str
    prenom: str
    numss: str
    adm: str
    datnais: str
    nbrtrav: str
    datent: str
    datsor: str
    quarterly_brutss_cents: tuple  # (q1_cents, q2_cents, q3_cents, q4_cents)
    brutss_annual: int             # Sum of all 4 quarters in cents


class AnnualStats(NamedTuple):
    """Summary statistics for the annual declaration."""
    unique_count: int
    quarterly_totals_cents: tuple   # (total_q1, total_q2, total_q3, total_q4) in cents
    grand_total_cents: int
    file_labels: tuple              # (label1, label2, label3, label4)


class AnnualResult(NamedTuple):
    """Full output of the annual merge, consumed by the exporter."""
    rows: list                  # list[AnnualRow]
    stats: AnnualStats
    missing_per_file: tuple     # (list[MissingEmployee] × 4)


def filename_to_label(filename: str) -> str:
    """Derive a clean column label from a filename (strip extension)."""
    stem, _ = os.path.splitext(filename)
    return stem.strip()
