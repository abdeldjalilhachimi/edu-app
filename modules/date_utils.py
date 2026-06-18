"""
date_utils.py

Shared date formatting for output files.

Source files mix date formats: some cells are plain text "13/09/1987", others
are real Excel dates that pandas reads as Timestamps → "1987-09-13 00:00:00".
This normalizes everything to the French convention DD/MM/YYYY for display.
"""

import re

import pandas as pd

# Already in DD/MM/YYYY (or D/M/YYYY) form — kept as-is, only zero-padded
_SLASH_DMY = re.compile(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*$")


def format_ddmmyyyy(value) -> str:
    """
    Return a date string as DD/MM/YYYY.

    - "13/09/1987"            → "13/09/1987"  (already day-first, just padded)
    - "5/9/1987"              → "05/09/1987"
    - "1987-09-13"            → "13/09/1987"
    - "1987-09-13 00:00:00"   → "13/09/1987"
    - Timestamp / datetime    → "13/09/1987"
    - "00/00/1953" or junk    → returned unchanged (cannot be parsed safely)
    - empty / NaN / NaT       → ""
    """
    if value is None:
        return ""
    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "none", "nat"):
        return ""

    # Already day-first with slashes → keep, just zero-pad day and month
    m = _SLASH_DMY.match(s)
    if m:
        d, mo, y = m.groups()
        return f"{int(d):02d}/{int(mo):02d}/{y}"

    # Otherwise parse (ISO / Timestamp text) and reformat; keep original if it
    # cannot be parsed (e.g. "00/00/1953" placeholders).
    parsed = pd.to_datetime(s, errors="coerce")
    if pd.isna(parsed):
        return s
    return parsed.strftime("%d/%m/%Y")
