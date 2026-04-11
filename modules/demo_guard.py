"""
demo_guard.py

Demo protection and activation code management.

Tracks the number of downloads in a persistent JSON file (data/license.json).
After MAX_FREE_DOWNLOADS, the app is locked until the correct activation
code is entered. Once activated, the app is unlocked forever on that machine.

The JSON file is stored locally — each machine has its own demo state.
"""

import json
import os

# ─────────────────────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────────────────────

MAX_FREE_DOWNLOADS = 3
ACTIVATION_CODE = "00020"
CONTACT_EMAIL = "abdeldjalil_hachimi@hotmail.com"

# Path to the license state file (relative to the app's working directory)
_DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
LICENSE_FILE = os.path.join(_DATA_DIR, "license.json")

# Default state for a fresh install
_DEFAULT_LICENSE = {
    "download_count": 0,
    "unlocked": False,
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
    """
    try:
        with open(LICENSE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
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


def increment_downloads() -> None:
    """
    Increment the download counter by 1 and save.

    Called every time the user downloads a result file.
    Does nothing if already unlocked.
    """
    lic = load_license()
    if lic.get("unlocked", False):
        return  # No tracking needed once unlocked
    lic["download_count"] = lic.get("download_count", 0) + 1
    save_license(lic)


def is_demo_expired() -> bool:
    """
    Check if the demo period has ended.

    Returns True if the user has used all free downloads AND
    has not entered the activation code.
    """
    lic = load_license()
    if lic.get("unlocked", False):
        return False
    return lic.get("download_count", 0) >= MAX_FREE_DOWNLOADS


def is_unlocked() -> bool:
    """Check if the app has been permanently activated."""
    lic = load_license()
    return lic.get("unlocked", False)


def get_remaining_downloads() -> int:
    """
    Return how many free downloads the user has left.

    Returns -1 if the app is unlocked (unlimited).
    """
    lic = load_license()
    if lic.get("unlocked", False):
        return -1  # Unlimited
    used = lic.get("download_count", 0)
    remaining = MAX_FREE_DOWNLOADS - used
    return max(0, remaining)


def try_activate(code: str) -> bool:
    """
    Attempt to activate the app with the given code.

    If the code matches ACTIVATION_CODE, the app is unlocked
    permanently and the function returns True.
    Otherwise returns False.
    """
    if code.strip() == ACTIVATION_CODE:
        lic = load_license()
        lic["unlocked"] = True
        save_license(lic)
        return True
    return False
