"""
trimestrial_types.py

Data model and types for the trimestrial BRUTSS consolidation feature.

All monetary values are stored as integer cents internally to avoid
floating-point rounding errors (e.g. 1234.56 → 123456 cents).
Conversion to display values (float) happens only at export time.

Column names are derived from uploaded filenames (without extension),
e.g. "paie_janvier.xlsx" → column "BRUTSS_paie_janvier".
"""

import os
from typing import NamedTuple


class EmployeeKey(NamedTuple):
    """
    Matching key for employees across monthly files.

    Uses NUMCPT only (normalized, leading zeros stripped).
    NOM/PRENOM are excluded from the key to avoid false negatives
    caused by name typos or encoding differences between files
    (e.g. "MEYMOUNA" in one file vs "MIMOUNA" in another).
    """
    numcpt: str   # Normalized NUMCPT (leading zeros stripped)


class MonthlyEntry(NamedTuple):
    """One employee's aggregated BRUTSS for a single month (in cents)."""
    numcpt_raw: str   # Original NUMCPT (preserves leading zeros for output)
    nom: str
    prenom: str
    numss: str        # NUMSS (social security number), empty string if absent
    adm: str          # ADM code, empty string if absent
    brutss_cents: int  # BRUTSS in integer cents (e.g. 1234.56 → 123456)


class MissingEmployee(NamedTuple):
    """An employee absent from a specific monthly file."""
    numcpt_raw: str   # Original NUMCPT (with leading zeros)
    nom: str
    prenom: str


class ConsolidatedRow(NamedTuple):
    """One row in the final trimestrial output."""
    numcpt_raw: str              # Original NUMCPT (with leading zeros preserved)
    nom: str
    prenom: str
    numss: str                   # NUMSS from first available month, "" if absent
    adm: str                     # ADM from first available month, "" if absent
    monthly_brutss_cents: tuple  # (m1_cents, m2_cents, m3_cents) in file order
    brutss_total: int            # Sum of all three months in cents


class TrimestrialStats(NamedTuple):
    """Summary statistics for the consolidation."""
    unique_count: int              # Number of unique employees
    monthly_totals_cents: tuple    # (total_m1, total_m2, total_m3) in cents
    grand_total_cents: int         # Grand total all months (cents)
    file_labels: tuple             # (label1, label2, label3) derived from filenames


class TrimestrialResult(NamedTuple):
    """Full output of the trimestrial merge, consumed by the exporter."""
    rows: list                  # list[ConsolidatedRow]
    stats: TrimestrialStats
    missing_per_file: tuple     # (list[MissingEmployee], list[MissingEmployee], list[MissingEmployee])


def filename_to_label(filename: str) -> str:
    """
    Derive a clean column label from a filename.

    Strips the file extension and returns the stem.
    Example: "paie_janvier_2025.xlsx" → "paie_janvier_2025"

    Parameters
    ----------
    filename : str
        Original uploaded filename.

    Returns
    -------
    str
        Filename without extension, suitable for use in column headers.
    """
    stem, _ = os.path.splitext(filename)
    return stem.strip()
