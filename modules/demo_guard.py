"""
demo_guard.py

Demo protection and activation code management.

Tracks the number of downloads in a persistent JSON file (data/license.json).
After MAX_FREE_DOWNLOADS, the app is locked until an activation code is entered.

Two activation codes:
  - TRIAL_CODE  ("00020") → unlocks for 1 month from activation date
  - FOREVER_CODE ("00034") → unlocks permanently, no expiration

The JSON file is stored locally — each machine has its own demo state.
"""

from __future__ import annotations

import json
import os
from datetime import datetime, timedelta
from typing import Optional

# ─────────────────────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────────────────────

MAX_FREE_DOWNLOADS = 3
TRIAL_CODE = "00020"          # 1-month trial
FOREVER_CODE = "00034"        # Permanent unlock
TRIAL_DAYS = 30               # Duration of trial in days
CONTACT_EMAIL = "abdeldjalil_hachimi@hotmail.com"

# Path to the license state file (relative to the app's working directory)
_DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
LICENSE_FILE = os.path.join(_DATA_DIR, "license.json")

# ── License types ───────────────────────────────────────────────────────────
# "none"      → demo mode (download counter active)
# "trial"     → 1-month trial (has expiration date)
# "permanent" → unlocked forever

# Default state for a fresh install
_DEFAULT_LICENSE = {
    "download_count": 0,
    "unlock_type": "none",          # "none" | "trial" | "permanent"
    "trial_activated_at": None,     # ISO date string, e.g. "2026-04-11"
}


# ─────────────────────────────────────────────────────────────────────────────
# Core functions
# ─────────────────────────────────────────────────────────────────────────────

def _ensure_data_dir() -> None:
    """Create the data/ directory if it doesn't exist."""
    os.makedirs(_DATA_DIR, exist_ok=True)


def load_license() -> dict:
    """
    Read the license state from disk.

    Returns the default state if the file doesn't exist or is corrupted.
    Handles migration from old format (unlocked: bool) to new format (unlock_type).
    """
    try:
        with open(LICENSE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)

        # ── Migrate old format ──────────────────────────────────────────
        # Old format had {"unlocked": true/false}.
        # Convert to new format on the fly.
        if "unlocked" in data and "unlock_type" not in data:
            if data.get("unlocked", False):
                data["unlock_type"] = "permanent"
            else:
                data["unlock_type"] = "none"
            data.pop("unlocked", None)
            data.setdefault("trial_activated_at", None)
            # Save migrated format
            _ensure_data_dir()
            with open(LICENSE_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)

        # Validate structure — merge with defaults for missing keys
        result = dict(_DEFAULT_LICENSE)
        result.update(data)
        return result
    except (FileNotFoundError, json.JSONDecodeError, TypeError):
        return dict(_DEFAULT_LICENSE)


def save_license(data: dict) -> None:
    """Write the license state to disk."""
    _ensure_data_dir()
    with open(LICENSE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def _is_trial_valid(lic: dict) -> bool:
    """Check if a trial license is still within its 30-day window."""
    activated_at = lic.get("trial_activated_at")
    if not activated_at:
        return False
    try:
        start_date = datetime.fromisoformat(activated_at)
        expiry_date = start_date + timedelta(days=TRIAL_DAYS)
        return datetime.now() < expiry_date
    except (ValueError, TypeError):
        return False


def _get_trial_expiry(lic: dict) -> Optional[datetime]:
    """Return the trial expiration datetime, or None if not a trial."""
    activated_at = lic.get("trial_activated_at")
    if not activated_at:
        return None
    try:
        start_date = datetime.fromisoformat(activated_at)
        return start_date + timedelta(days=TRIAL_DAYS)
    except (ValueError, TypeError):
        return None


def increment_downloads() -> None:
    """
    Increment the download counter by 1 and save.

    Called every time the user downloads a result file.
    Does nothing if the app is unlocked (trial or permanent).
    """
    lic = load_license()
    unlock_type = lic.get("unlock_type", "none")

    if unlock_type == "permanent":
        return
    if unlock_type == "trial" and _is_trial_valid(lic):
        return

    lic["download_count"] = lic.get("download_count", 0) + 1
    save_license(lic)


def is_demo_expired() -> bool:
    """
    Check if the app should be locked.

    Returns True if:
    - Demo mode AND all free downloads used, OR
    - Trial mode AND the 30-day window has passed
    """
    lic = load_license()
    unlock_type = lic.get("unlock_type", "none")

    if unlock_type == "permanent":
        return False

    if unlock_type == "trial":
        if _is_trial_valid(lic):
            return False
        # Trial expired — lock the app
        return True

    # Demo mode: check download count
    return lic.get("download_count", 0) >= MAX_FREE_DOWNLOADS


def is_unlocked() -> bool:
    """
    Check if the app is currently active (not in limited demo mode).

    Returns True for both valid trial and permanent licenses.
    """
    lic = load_license()
    unlock_type = lic.get("unlock_type", "none")

    if unlock_type == "permanent":
        return True
    if unlock_type == "trial" and _is_trial_valid(lic):
        return True
    return False


def get_remaining_downloads() -> int:
    """
    Return how many free downloads the user has left.

    Returns -1 if the app is unlocked (trial or permanent).
    """
    lic = load_license()
    unlock_type = lic.get("unlock_type", "none")

    if unlock_type == "permanent":
        return -1
    if unlock_type == "trial" and _is_trial_valid(lic):
        return -1

    used = lic.get("download_count", 0)
    remaining = MAX_FREE_DOWNLOADS - used
    return max(0, remaining)


def get_trial_days_remaining() -> Optional[int]:
    """
    Return the number of days left in the trial, or None if not a trial.

    Returns 0 if the trial has expired.
    """
    lic = load_license()
    if lic.get("unlock_type") != "trial":
        return None
    expiry = _get_trial_expiry(lic)
    if expiry is None:
        return None
    delta = expiry - datetime.now()
    return max(0, delta.days)


def get_lock_reason() -> str:
    """
    Return a human-readable reason why the app is locked.

    Used by the activation dialog to show the right message.
    """
    lic = load_license()
    unlock_type = lic.get("unlock_type", "none")

    if unlock_type == "trial" and not _is_trial_valid(lic):
        return "trial_expired"
    return "demo_exhausted"


def try_activate(code: str) -> str:
    """
    Attempt to activate the app with the given code.

    Parameters
    ----------
    code : str
        The activation code entered by the user.

    Returns
    -------
    str
        "permanent" — if FOREVER_CODE matched (unlocked forever)
        "trial"     — if TRIAL_CODE matched (1-month trial started)
        ""          — if code is wrong (no change)
    """
    stripped = code.strip()

    if stripped == FOREVER_CODE:
        lic = load_license()
        lic["unlock_type"] = "permanent"
        save_license(lic)
        return "permanent"

    if stripped == TRIAL_CODE:
        lic = load_license()
        lic["unlock_type"] = "trial"
        lic["trial_activated_at"] = datetime.now().isoformat()
        save_license(lic)
        return "trial"

    return ""
