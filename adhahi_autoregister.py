"""
adhahi.dz Auto-Registration Script
====================================
Monitors the adhahi.dz platform and registers you automatically
when an appointment slot becomes available in your wilaya.

Requirements:
    pip install requests selenium webdriver-manager playsound

Usage:
    1. Fill in your personal details in the CONFIG section below
    2. Run: python adhahi_autoregister.py
"""

import time
import requests
import logging
from datetime import datetime

# ─────────────────────────────────────────────
#  CONFIG — Fill in YOUR details here
# ─────────────────────────────────────────────
CONFIG = {
    "nin": "YOUR_NIN_HERE",                  # رقم التعريف الوطني
    "cnibe": "YOUR_CNIBE_HERE",              # رقم بطاقة التعريف
    "phone": "YOUR_PHONE_NUMBER",            # e.g. 0551234567
    "email": "",                             # Optional
    "password": "YourStrongPassword123!",
    "wilaya": "08",                          # 08 = Béchar
    "commune": "Bechar",                     # Your commune
    "payment_method": "online",              # "cash", "tpe", or "online"

    # Monitoring settings
    "check_interval_seconds": 10,            # How often to check (seconds)
    "base_url": "https://adhahi.dz",
}

# Common Wilaya codes:
# 01=Adrar, 02=Chlef, 03=Laghouat, 09=Blida, 16=Alger,
# 25=Constantine, 31=Oran, 47=Ghardaia ...

# ─────────────────────────────────────────────
#  LOGGING
# ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("adhahi_log.txt"),
    ]
)
log = logging.getLogger(__name__)

# ─────────────────────────────────────────────
#  SESSION SETUP
# ─────────────────────────────────────────────
session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "ar,fr;q=0.9,en;q=0.8",
    "Referer": CONFIG["base_url"],
    "Origin": CONFIG["base_url"],
})

# Session cookie from your authenticated browser session
# ⚠️  Cookie expires — if you get 401 again, grab a fresh one from:
#     Chrome F12 > Application tab > Cookies > adhahi.dz > cookiesession1
SESSION_COOKIE = "678A3E44E896DCFFBFB8060BF7952433"
session.cookies.set("cookiesession1", SESSION_COOKIE, domain="adhahi.dz")

# ─────────────────────────────────────────────
#  STEP 1: Check slot availability
# ─────────────────────────────────────────────
def check_availability() -> bool:
    """
    Returns True if slots are available in the configured wilaya.
    Adjust the endpoint/params once you inspect the actual API via browser DevTools.
    """
    try:
        url = f"{CONFIG['base_url']}/api/availability"
        params = {
            "wilaya": CONFIG["wilaya"],
            "commune": CONFIG["commune"],
        }
        resp = session.get(url, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        log.info(f"Availability response: {data}")

        # Adjust these keys to match the real API response
        # Common patterns — script tries several:
        available = (
            data.get("available") is True
            or data.get("status") == "available"
            or data.get("slots", 0) > 0
            or data.get("quota", 0) > 0
            or data.get("hasSlots") is True
            or data.get("open") is True
        )
        return available

    except requests.exceptions.ConnectionError:
        log.warning("Connection error — site may be down or overloaded. Retrying...")
        return False
    except requests.exceptions.Timeout:
        log.warning("Request timed out. Retrying...")
        return False
    except Exception as e:
        log.error(f"Availability check failed: {e}")
        return False


# ─────────────────────────────────────────────
#  STEP 2: Register account
# ─────────────────────────────────────────────
def register() -> bool:
    """
    Submits the registration form. Returns True on success.
    """
    url = f"{CONFIG['base_url']}/api/register"   # Adjust to real endpoint
    payload = {
        "nin": CONFIG["nin"],
        "cnibe": CONFIG["cnibe"],
        "phone": CONFIG["phone"],
        "email": CONFIG["email"],
        "password": CONFIG["password"],
        "wilaya": CONFIG["wilaya"],
        "commune": CONFIG["commune"],
        "payment_method": CONFIG["payment_method"],
    }
    try:
        resp = session.post(url, json=payload, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        log.info(f"Registration response: {data}")

        success = data.get("success") or data.get("status") == "ok"
        return success

    except Exception as e:
        log.error(f"Registration failed: {e}")
        return False


# ─────────────────────────────────────────────
#  STEP 3: Confirm OTP (if needed)
# ─────────────────────────────────────────────
def confirm_otp(otp_code: str) -> bool:
    """
    Sends the OTP code received via SMS to complete registration.
    """
    url = f"{CONFIG['base_url']}/api/confirm-otp"   # Adjust to real endpoint
    payload = {
        "phone": CONFIG["phone"],
        "otp": otp_code,
    }
    try:
        resp = session.post(url, json=payload, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        log.info(f"OTP confirmation response: {data}")
        return data.get("success") or data.get("status") == "ok"
    except Exception as e:
        log.error(f"OTP confirmation failed: {e}")
        return False


# ─────────────────────────────────────────────
#  ALERT (sound + print)
# ─────────────────────────────────────────────
def alert(message: str):
    print("\n" + "=" * 60)
    print(f"  🚨 ALERT: {message}")
    print("=" * 60 + "\n")
    try:
        # Optional: play a beep sound
        import os
        if os.name == "nt":       # Windows
            import winsound
            winsound.Beep(1000, 1500)
        else:                     # Linux/Mac
            os.system("echo '\a'")
    except Exception:
        pass


# ─────────────────────────────────────────────
#  MAIN LOOP
# ─────────────────────────────────────────────
def main():
    log.info("=" * 60)
    log.info("  adhahi.dz Auto-Registration Monitor")
    log.info(f"  Wilaya: {CONFIG['wilaya']}  |  Commune: {CONFIG['commune']}")
    log.info(f"  Checking every {CONFIG['check_interval_seconds']} seconds")
    log.info("=" * 60)

    attempt = 0
    while True:
        attempt += 1
        log.info(f"[Check #{attempt}] Checking for available slots...")

        if check_availability():
            alert("Slots are AVAILABLE! Attempting registration now...")

            if register():
                alert("Registration submitted! Check your phone for the OTP code.")
                otp = input("Enter the OTP code you received via SMS: ").strip()
                if confirm_otp(otp):
                    alert("SUCCESS! Registration fully confirmed. Check adhahi.dz for your booking details.")
                    break
                else:
                    log.error("OTP confirmation failed. You may need to try again manually.")
                    break
            else:
                log.warning("Registration attempt failed. Will retry on next check...")
        else:
            log.info(f"No slots available yet. Next check in {CONFIG['check_interval_seconds']}s...")

        time.sleep(CONFIG["check_interval_seconds"])


if __name__ == "__main__":
    main()