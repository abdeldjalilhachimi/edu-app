"""
cleaner_fill.py

Clean an annual declaration file and reconcile it with the official CNAS file.

Three stages:
  1. CLEAN  — normalize DATNAIS / DATENT / DATSOR to DD/MM/YYYY and blank out
     junk dates (30/12/1899, 00/00/…, impossible years).
  2. FILL / OVERRIDE — the official file (keyed on NUMSS) is authoritative:
       - empty NUMSS  → matched by NOM + PRENOM (+ DATNAIS) so the row can match,
       - DATNAIS / DATENT → OVERWRITTEN with the official value (cleared if the
         official has none),
       - DATSOR → cleared for anyone in the official file (active → no exit date).
     Rows not found in the official file keep their data (only junk dates cleaned).
  3. MERGE  — rows sharing the same NUMSS are gathered into ONE row: the
     per-quarter days and salaries (NBRTRAV_*, BRUTSS_*) and the totals are
     summed; the most complete NUMCPT and the identity fields are kept.

Output is a fresh single-sheet workbook (Déclaration Annuelle). Grand totals
are preserved — merging only sums duplicates, it never loses money or days.
"""

import io
import re
import datetime

import pandas as pd

from modules.comparator import (
    load_keyed_file,
    _norm_text,
    _norm_date,
    _normalize_key,
)

DATE_COLUMNS = ["DATNAIS", "DATENT", "DATSOR"]
_SLASH = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$")


def clean_date(value) -> str:
    """Return a valid DD/MM/YYYY date, or "" for empty/junk/implausible dates."""
    if value is None:
        return ""
    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "none", "nat"):
        return ""

    m = _SLASH.match(s)
    if m:
        d, mo, y = (int(x) for x in m.groups())
        try:
            datetime.date(y, mo, d)
        except ValueError:
            return ""
        return f"{d:02d}/{mo:02d}/{y}" if 1920 <= y <= 2100 else ""

    parsed = pd.to_datetime(s, errors="coerce")
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%d/%m/%Y") if 1920 <= parsed.year <= 2100 else ""


def _norm_numss(value) -> str:
    """Normalized NUMSS: trimmed, no trailing .0, no leading zeros."""
    s = str(value).strip()
    return s.replace(".0", "").lstrip("0")


def _build_official_lookups(official_obj, official_name: str):
    """Build NUMSS→(datnais,datent) and name→NUMSS lookups from the official file."""
    off = load_keyed_file(official_obj, official_name)
    off_ns = _normalize_key(off["NUMSS"])

    by_numss = {}                       # nsnorm -> (datnais, datent)
    name3, amb3 = {}, set()             # (nom,prenom,ymd) -> numss
    name2, amb2 = {}, set()             # (nom,prenom) -> numss

    has_datent = "DATENT" in off.columns
    has_datnais = "DATNAIS" in off.columns

    for i in range(len(off)):
        ns_raw = str(off["NUMSS"].iloc[i]).strip()
        nsn = off_ns.iloc[i]
        dn = off["DATNAIS"].iloc[i] if has_datnais else ""
        de = off["DATENT"].iloc[i] if has_datent else ""

        if nsn not in ("", "0"):
            by_numss.setdefault(nsn, (clean_date(dn), clean_date(de)))

        nom, pre = _norm_text(off["NOM"].iloc[i]), _norm_text(off["PRENOM"].iloc[i])
        if not nom:
            continue
        k3 = (nom, pre, _norm_date(dn))
        if k3 in name3 and name3[k3] != ns_raw:
            amb3.add(k3)
        else:
            name3.setdefault(k3, ns_raw)
        k2 = (nom, pre)
        if k2 in name2 and name2[k2] != ns_raw:
            amb2.add(k2)
        else:
            name2.setdefault(k2, ns_raw)

    return by_numss, name3, amb3, name2, amb2


def _numeric_columns(columns) -> list:
    """Columns to sum when merging duplicates: days + salaries (per quarter + totals)."""
    return [
        c for c in columns
        if c == "NBRTRAV" or c.startswith("NBRTRAV_") or c.startswith("BRUTSS_")
    ]


def clean_and_fill(annual_obj, annual_name: str,
                   official_obj, official_name: str) -> tuple:
    """
    Clean, reconcile against the official file, and merge duplicate-NUMSS rows.
    Returns (xlsx_bytes, stats).
    """
    by_numss, name3, amb3, name2, amb2 = _build_official_lookups(official_obj, official_name)

    annual_obj.seek(0)
    xls = pd.ExcelFile(annual_obj, engine="openpyxl")
    sheet = "Déclaration Annuelle" if "Déclaration Annuelle" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
    df.columns = df.columns.astype(str).str.strip()
    df = df.fillna("")

    stats = {
        "rows_in": len(df),
        "junk_removed": 0,
        "numss_filled": 0,
        "datnais_set": 0,
        "datent_set": 0,
        "datsor_cleared": 0,
        "rows_merged": 0,
        "rows_out": 0,
    }

    has = lambda c: c in df.columns  # noqa: E731

    # ── 1. Clean date columns ────────────────────────────────────────────
    for dc in DATE_COLUMNS:
        if has(dc):
            cleaned = df[dc].apply(clean_date)
            stats["junk_removed"] += int(((df[dc].str.strip() != "") & (cleaned == "")).sum())
            df[dc] = cleaned

    # ── 2. Fill empty NUMSS by name (so the row can match the official) ──
    if has("NUMSS") and has("NOM"):
        for i in df.index:
            if _norm_numss(df.at[i, "NUMSS"]) in ("", "0"):
                nom = _norm_text(df.at[i, "NOM"])
                pre = _norm_text(df.at[i, "PRENOM"]) if has("PRENOM") else ""
                dn = df.at[i, "DATNAIS"] if has("DATNAIS") else ""
                k3, k2 = (nom, pre, _norm_date(dn)), (nom, pre)
                found = None
                if nom and k3 in name3 and k3 not in amb3:
                    found = name3[k3]
                elif nom and k2 in name2 and k2 not in amb2:
                    found = name2[k2]
                if found:
                    df.at[i, "NUMSS"] = found
                    stats["numss_filled"] += 1

    # ── 3. Official file is authoritative — overwrite/clear ─────────────
    if has("NUMSS"):
        for i in df.index:
            official = by_numss.get(_norm_numss(df.at[i, "NUMSS"]))
            if not official:
                continue
            odn, ode = official
            if has("DATNAIS") and df.at[i, "DATNAIS"] != odn:
                df.at[i, "DATNAIS"] = odn
                if odn:
                    stats["datnais_set"] += 1
            if has("DATENT") and df.at[i, "DATENT"] != ode:
                df.at[i, "DATENT"] = ode
                if ode:
                    stats["datent_set"] += 1
            if has("DATSOR") and df.at[i, "DATSOR"].strip() != "":
                df.at[i, "DATSOR"] = ""        # official has no exit date → active
                stats["datsor_cleared"] += 1

    # ── 4. Merge rows that share the same NUMSS ──────────────────────────
    num_cols = _numeric_columns(df.columns)
    for c in num_cols:
        df[c] = pd.to_numeric(
            df[c].astype(str).str.replace(" ", "", regex=False).str.replace(",", ".", regex=False),
            errors="coerce",
        ).fillna(0.0)

    # Group key: valid NUMSS → shared; empty NUMSS → unique per row (never merged)
    keys = []
    for pos, idx in enumerate(df.index):
        nsn = _norm_numss(df.at[idx, "NUMSS"]) if has("NUMSS") else ""
        keys.append(f"S:{nsn}" if nsn not in ("", "0") else f"R:{pos}")
    df["_key"] = keys
    df["_nclen"] = df["NUMCPT"].astype(str).str.strip().str.len() if has("NUMCPT") else 0

    keeper_idx = df.groupby("_key")["_nclen"].idxmax()
    keepers = df.loc[keeper_idx].set_index("_key")
    sums = df.groupby("_key")[num_cols].sum()

    merged = keepers.copy()
    for c in num_cols:
        merged[c] = sums[c]
    merged = merged.reset_index(drop=True).drop(columns=["_nclen"])

    stats["rows_merged"] = len(df) - len(merged)
    stats["rows_out"] = len(merged)

    # Integer days, 2-decimal salaries, renumbered line counter
    for c in num_cols:
        if c == "NBRTRAV" or c.startswith("NBRTRAV_"):
            merged[c] = merged[c].round().astype(int)
        else:
            merged[c] = merged[c].round(2)
    if "N" in merged.columns:
        merged["N"] = range(1, len(merged) + 1)

    # ── 5. Write fresh workbook ──────────────────────────────────────────
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="Déclaration Annuelle", index=False)
    return buffer.getvalue(), stats
