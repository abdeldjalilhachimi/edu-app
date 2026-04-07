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

Performance: uses vectorized pandas operations (concat, groupby, map)
instead of row-by-row iteration for ~50-100x speedup on large files.
"""

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


def build_additional_lookup(additional_dfs: list) -> tuple:
    """
    Merge all additional files into one lookup per NUMCPT key.

    Uses vectorized concat + groupby instead of row-by-row iteration.

    Returns
    -------
    tuple (brutss_sums: pd.Series, first_rows: pd.DataFrame)
        brutss_sums: Series indexed by _KEY with summed BRUTSS.
        first_rows: DataFrame indexed by _KEY with first-seen row per key.
        Both are empty if no valid additional data exists.
    """
    if not additional_dfs:
        empty_idx = pd.Index([], name=KEY_COLUMN)
        return pd.Series(dtype="float64", index=empty_idx), pd.DataFrame()

    # Concatenate all additional DataFrames at once
    all_add = pd.concat(additional_dfs, ignore_index=True)

    # Filter out rows with empty/blank keys (vectorized)
    key_col = all_add[KEY_COLUMN].astype(str).str.strip()
    valid_mask = key_col.ne("") & key_col.notna()
    all_add = all_add[valid_mask]

    if all_add.empty:
        empty_idx = pd.Index([], name=KEY_COLUMN)
        return pd.Series(dtype="float64", index=empty_idx), pd.DataFrame()

    # Vectorized groupby: sum BRUTSS per key
    brutss_sums = all_add.groupby(KEY_COLUMN)[BRUTSS_COLUMN].sum()

    # Keep first-seen row per key (for additional-only appending)
    first_rows = all_add.groupby(KEY_COLUMN).first()

    return brutss_sums, first_rows


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
    brutss_sums, first_rows = build_additional_lookup(additional_dfs)

    # ── Step 1: Vectorized merge — map additional BRUTSS onto main keys ──
    result_df = main_df.copy()
    add_brutss = result_df[KEY_COLUMN].map(brutss_sums).fillna(0.0)

    # Save original BRUTSS before updating (for DuplicateMatch records)
    brutss_main_col = result_df[BRUTSS_COLUMN].astype(float)
    result_df[BRUTSS_COLUMN] = brutss_main_col + add_brutss

    # ── Build duplicates list from matched rows ─────────────────────────
    matched_mask = add_brutss > 0
    consumed_keys = set(result_df.loc[matched_mask, KEY_COLUMN])

    duplicates = []
    if matched_mask.any():
        matched = result_df[matched_mask]
        matched_main_brutss = brutss_main_col[matched_mask]
        matched_add_brutss = add_brutss[matched_mask]

        # Use itertuples (much faster than iterrows) on the small matched subset
        for tup in zip(
            matched["NUMCPT"].astype(str).str.strip(),
            matched["NOM"].astype(str).str.strip(),
            matched["PRENOM"].astype(str).str.strip(),
            matched_main_brutss,
            matched_add_brutss,
            matched[BRUTSS_COLUMN],
        ):
            duplicates.append(DuplicateMatch(*tup))

    # ── Step 2: Append additional-only rows ──────────────────────────────
    added_count = 0

    if not first_rows.empty:
        additional_only_keys = set(brutss_sums.index) - consumed_keys

        if additional_only_keys:
            # Select additional-only rows from the grouped first_rows
            add_only = first_rows.loc[first_rows.index.isin(additional_only_keys)].copy()
            add_only[BRUTSS_COLUMN] = brutss_sums[add_only.index]

            # Build DuplicateMatch records for additional-only rows
            for key in sorted(additional_only_keys):
                row = add_only.loc[key]
                b = float(brutss_sums[key])
                duplicates.append(DuplicateMatch(
                    numcpt=str(row.get("NUMCPT", "")).strip(),
                    nom=str(row.get("NOM", "")).strip(),
                    prenom=str(row.get("PRENOM", "")).strip(),
                    brutss_main=0.0,
                    brutss_additional=b,
                    brutss_sum=b,
                ))

            # Reindex to main columns and append
            add_only = add_only.reset_index()
            main_columns = result_df.columns.tolist()
            add_only = add_only.reindex(columns=main_columns)
            result_df = pd.concat([result_df, add_only], ignore_index=True)
            added_count = len(additional_only_keys)

    # ── Step 3: Compute final stats ──────────────────────────────────────
    brutss_total = result_df[BRUTSS_COLUMN].sum()

    stats = {
        "total": len(result_df),
        "duplicate_count": len(consumed_keys),
        "added_count": added_count,
        "brutss_total": brutss_total,
    }

    return CalculationResult(
        updated_main_df=result_df,
        duplicates=duplicates,
        stats=stats,
    )
