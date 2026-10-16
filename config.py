from __future__ import annotations

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def _data_dir() -> str:
    preferred = os.environ.get("WRS_DATA_DIR", os.path.join(BASE_DIR, "data"))
    try:
        os.makedirs(preferred, exist_ok=True)
        test = os.path.join(preferred, ".write_test")
        with open(test, "w") as f:
            f.write("1")
        os.remove(test)
        return preferred
    except Exception:
        fallback = "/tmp/wrs_data"
        os.makedirs(fallback, exist_ok=True)
        return fallback

DATA_DIR = _data_dir()
SETTINGS_FILE = os.path.join(DATA_DIR, "settings.json")
ADDRESS_FILE = os.path.join(DATA_DIR, "address_book.json")
TRACKER_FILE = os.path.join(DATA_DIR, "streak_tracker.json")
CORRECTIONS_FILE = os.path.join(DATA_DIR, "corrections.json")
RECON_FILE = os.path.join(DATA_DIR, "reconstruction.json")
EVENTS_FILE = os.path.join(DATA_DIR, "events_log.json")
HISTORY_FILE = os.path.join(DATA_DIR, "history.json")
DVORNIK_FILE = os.path.join(DATA_DIR, "dvornik_schedule.json")
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

APP_TITLE = "WRS Mobile"
APP_VER = "1.2.0"
