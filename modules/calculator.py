"""
calculator.py

Key-based merge: match rows between main file and additional files
by composite key (NOM|PRENOM|NUMCPT). Only matched rows get their
BRUTSS updated. Unmatched rows stay unchanged.

Logic (mirrors the user's existing JS implementation):
1. Build a lookup from all additional files: {key → sum_of_brutss}
2. For each main row: if key matches → BRUTSS += additional BRUTSS
3. Record every match as a DuplicateMatch with full breakdown
"""

import math
from typing import NamedTuple
import pandas as pd

BRUTSS_COLUMN = "BRUTSS"
KEY_COLUMN = "_KEY"


class DuplicateMatch(NamedTuple):
    """One matched row between main file and additional files."""
    numcpt: str
    nom: str
    prenom: str
    brutss_main: float       # Original BRUTSS in main file
    brutss_additional: float  # Total BRUTSS from additional files for this key
    brutss_sum: float         # brutss_main + brutss_additional


class CalculationResult(NamedTuple):
    """Full output of run_calculation, consumed by the exporter."""
    updated_main_df: pd.DataFrame
    duplicates: list          # list[DuplicateMatch]
    stats: dict               # {total, duplicate_count, brutss_total}


def build_additional_lookup(additional_dfs: list) -> dict:
    """
    Merge all additional files into one lookup: {composite_key: sum_of_brutss}.

    If the same key appears in multiple additional files (or multiple times
    within one file), their BRUTSS values are summed together.
    Keys that are "||" (all empty NOM/PRENOM/NUMCPT) are skipped.

    Parameters
    ----------
    additional_dfs : list[pd.DataFrame]
        Each must have _KEY and BRUTSS columns (float64).

    Returns
    -------
    dict[str, float]
        {composite_key: total_brutss_from_all_additional_files}
    """
    lookup: dict = {}

    for df in additional_dfs:
        for _, row in df.iterrows():
            key = row[KEY_COLUMN]
            if key == "||":
                continue

            brutss_val = float(row[BRUTSS_COLUMN])

            if key in lookup:
                lookup[key] = lookup[key] + brutss_val
            else:
                lookup[key] = brutss_val

    return lookup


def run_calculation(
    main_df: pd.DataFrame,
    additional_dfs: list,
) -> CalculationResult:
    """
    Key-based merge between main file and additional files.

    For each main row:
    - If key exists in additional lookup → BRUTSS = main + additional
    - If key does NOT exist → BRUTSS stays unchanged

    Parameters
    ----------
    main_df : pd.DataFrame
        Must have _KEY and BRUTSS columns (float64).
    additional_dfs : list[pd.DataFrame]

    Returns
    -------
    CalculationResult
    """
    lookup = build_additional_lookup(additional_dfs)

    duplicates = []
    updated_brutss = []

    for _, row in main_df.iterrows():
        key = row[KEY_COLUMN]
        brutss_main = float(row[BRUTSS_COLUMN])

        additional_brutss = lookup.get(key)

        if additional_brutss is not None and key != "||":
            brutss_sum = brutss_main + additional_brutss
            updated_brutss.append(brutss_sum)

            duplicates.append(DuplicateMatch(
                numcpt=str(row.get("NUMCPT", "")).strip(),
                nom=str(row.get("NOM", "")).strip(),
                prenom=str(row.get("PRENOM", "")).strip(),
                brutss_main=brutss_main,
                brutss_additional=additional_brutss,
                brutss_sum=brutss_sum,
            ))
        else:
            updated_brutss.append(brutss_main)

    # Build result DataFrame
    result_df = main_df.copy()
    result_df[BRUTSS_COLUMN] = updated_brutss

    # Compute final column total using math.fsum for precision
    brutss_total = math.fsum(updated_brutss)

    stats = {
        "total": len(main_df),
        "duplicate_count": len(duplicates),
        "brutss_total": brutss_total,
    }

    return CalculationResult(
        updated_main_df=result_df,
        duplicates=duplicates,
        stats=stats,
    )
