"""Submittal Tracker — watches the drop folder and logs submittal PDFs to Excel."""
import json
import os
import time
import traceback
from datetime import datetime

import anthropic

from extractor import extract_submittal
from excel_writer import append_submittal

_CFG = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "config.json")
with open(_CFG, "r", encoding="utf-8") as _f:
    _cfg = json.load(_f)
WATCH = _cfg["watch_folder"]
XLSX = _cfg["log_xlsx"]
LOG = _cfg["activity_log"]
STATE = os.path.join(WATCH, ".processed.json")
POLL_SECONDS = int(_cfg.get("poll_seconds", 5))


def log(msg):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line, flush=True)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def load_state():
    try:
        with open(STATE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_state(state):
    with open(STATE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=1)


def is_ready(path):
    """File is fully copied (size stable across 2 seconds)."""
    try:
        s1 = os.path.getsize(path)
        time.sleep(2)
        return os.path.getsize(path) == s1
    except OSError:
        return False


def main():
    os.makedirs(WATCH, exist_ok=True)
    state = load_state()
    failures = {}
    log(f"Watching: {WATCH}")
    log(f"Logging to: {XLSX}")
    while True:
        try:
            for name in sorted(os.listdir(WATCH)):
                if not name.lower().endswith(".pdf"):
                    continue
                path = os.path.join(WATCH, name)
                key = f"{name}|{os.path.getmtime(path):.0f}"
                if key in state:
                    continue
                if not is_ready(path):
                    continue
                if "#" in name:
                    # Excel can't hyperlink to filenames containing '#'
                    new_name = name.replace("#", "_")
                    try:
                        os.rename(path, os.path.join(WATCH, new_name))
                        log(f"Renamed '{name}' -> '{new_name}' "
                            f"(Excel can't link to filenames with #)")
                    except OSError:
                        pass
                    continue  # picked up under the new name on the next pass
                log(f"Processing: {name}")
                try:
                    data = extract_submittal(path)
                    sheet = append_submittal(XLSX, data, pdf_filename=name)
                    state[key] = {"when": datetime.now().isoformat(),
                                  "sheet": sheet,
                                  "number": data.get("submittal_number"),
                                  "title": data.get("title")}
                    save_state(state)
                    log(f"  -> logged to sheet '{sheet}': "
                        f"#{data.get('submittal_number')}  {data.get('title')}")
                except (PermissionError, OSError) as e:
                    # transient: Excel has the log open, or Dropbox is still syncing.
                    log(f"  .. transient ({e.__class__.__name__}): {e} — will retry.")
                except (anthropic.RateLimitError, anthropic.InternalServerError,
                        anthropic.APIConnectionError) as e:
                    # transient: API rate limit / hiccup — never counts as a strike.
                    log(f"  .. API busy ({e.__class__.__name__}) — waiting 30s, will retry.")
                    time.sleep(30)
                except Exception as e:
                    fails = failures.get(key, 0) + 1
                    failures[key] = fails
                    log(f"  !! FAILED on {name} (attempt {fails}/3): {e}")
                    if fails >= 3:
                        log(traceback.format_exc(limit=3))
                        state[key] = {"when": datetime.now().isoformat(), "error": str(e)}
                        save_state(state)
                        log(f"  !! giving up on {name} after 3 attempts — fix the file and re-drop it.")
        except Exception as e:
            log(f"watcher loop error: {e}")
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
