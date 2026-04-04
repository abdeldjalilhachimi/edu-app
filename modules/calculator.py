"""
calculator.py

Full outer merge between main file and additional files by NUMCPT.

Logic:
1. Build a lookup from all additional files: {NUMCPT_key → (sum_brutss, representative_row)}
2. For each main row:
   - If NUMCPT matches in lookup → BRUTSS = main + additional (duplicate match)
   - If no match → keep row unchanged (main-only)
3. For each additional key NOT found in main → append as a new row (additional-only)

Result = union of main + additional. No NUMCPT is ever lost.
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
    brutss_main: float        # Original BRUTSS in main file (0.0 for additional-only)
    brutss_additional: float  # Total BRUTSS from additional files for this key
    brutss_sum: float         # brutss_main + brutss_additional


class CalculationResult(NamedTuple):
    """Full output of run_calculation, consumed by the exporter."""
    updated_main_df: pd.DataFrame
    duplicates: list          # list[DuplicateMatch]
    stats: dict               # {total, duplicate_count, added_count, brutss_total}


def build_additional_lookup(additional_dfs: list) -> dict:
    """
    Merge all additional files into one lookup per NUMCPT key.

    Returns dict[str, dict] where each value contains:
      - "brutss": sum of BRUTSS across all additional files for this key
      - "row":    the first encountered full row (used when appending
                  additional-only rows so we have NOM, PRENOM, etc.)

    If the same key appears multiple times, BRUTSS values are summed
    and the first row is kept as the representative.

    Keys that are empty ("", "0", or NaN-derived) are skipped.
    """
    lookup: dict = {}

    for df in additional_dfs:
        for _, row in df.iterrows():
            key = row[KEY_COLUMN]

            # Skip empty / invalid keys
            if not key or key == "" or str(key).strip() == "":
                continue

            brutss_val = float(row[BRUTSS_COLUMN])

            if key in lookup:
                lookup[key]["brutss"] += brutss_val
            else:
                # Store the full row as representative for additional-only appends
                lookup[key] = {
                    "brutss": brutss_val,
                    "row": row.copy(),
                }

    return lookup


def run_calculation(
    main_df: pd.DataFrame,
    additional_dfs: list,
) -> CalculationResult:
    """
    Full outer merge between main file and additional files by NUMCPT.

    1. Matched rows:      NUMCPT in both → BRUTSS = main + additional
    2. Main-only rows:    NUMCPT only in main → keep unchanged
    3. Additional-only:   NUMCPT only in additional → append as new row

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

    # Track which additional keys were consumed (matched with main)
    consumed_keys: set = set()
    duplicates = []
    updated_brutss = []

    # ── Step 1: Process main rows ────────────────────────────────────────────
    for _, row in main_df.iterrows():
        key = row[KEY_COLUMN]
        brutss_main = float(row[BRUTSS_COLUMN])

        entry = lookup.get(key)

        if entry is not None and key and key != "":
            # Matched: NUMCPT exists in both main and additional
            brutss_additional = entry["brutss"]
            brutss_sum = brutss_main + brutss_additional
            updated_brutss.append(brutss_sum)
            consumed_keys.add(key)

            duplicates.append(DuplicateMatch(
                numcpt=str(row.get("NUMCPT", "")).strip(),
                nom=str(row.get("NOM", "")).strip(),
                prenom=str(row.get("PRENOM", "")).strip(),
                brutss_main=brutss_main,
                brutss_additional=brutss_additional,
                brutss_sum=brutss_sum,
            ))
        else:
            # Main-only: keep unchanged
            updated_brutss.append(brutss_main)

    # Build result DataFrame from main rows
    result_df = main_df.copy()
    result_df[BRUTSS_COLUMN] = updated_brutss

    # ── Step 2: Append additional-only rows ──────────────────────────────────
    additional_only_keys = set(lookup.keys()) - consumed_keys
    added_count = 0

    if additional_only_keys:
        new_rows = []
        for key in sorted(additional_only_keys):
            entry = lookup[key]
            representative_row = entry["row"].copy()

            # Set BRUTSS to the aggregated sum from all additional files
            representative_row[BRUTSS_COLUMN] = entry["brutss"]

            new_rows.append(representative_row)

            duplicates.append(DuplicateMatch(
                numcpt=str(representative_row.get("NUMCPT", "")).strip(),
                nom=str(representative_row.get("NOM", "")).strip(),
                prenom=str(representative_row.get("PRENOM", "")).strip(),
                brutss_main=0.0,
                brutss_additional=entry["brutss"],
                brutss_sum=entry["brutss"],
            ))

        # Append new rows — keep only columns that exist in the main file
        # (additional files may have extra columns like BAREME, DATDEB, etc.)
        main_columns = result_df.columns.tolist()
        new_rows_df = pd.DataFrame(new_rows)
        # Reindex to main columns: missing cols become NaN, extra cols dropped
        new_rows_df = new_rows_df.reindex(columns=main_columns)
        result_df = pd.concat([result_df, new_rows_df], ignore_index=True)
        added_count = len(new_rows)

    # ── Step 3: Compute final stats ──────────────────────────────────────────
    brutss_total = math.fsum(result_df[BRUTSS_COLUMN].astype(float).tolist())

    stats = {
        "total": len(result_df),
        "duplicate_count": len(consumed_keys),   # Matched in both
        "added_count": added_count,               # Additional-only appended
        "brutss_total": brutss_total,
    }

    return CalculationResult(
        updated_main_df=result_df,
        duplicates=duplicates,
        stats=stats,
    )
