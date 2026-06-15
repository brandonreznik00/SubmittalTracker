"""SubmittalTracker cloud watcher — runs on a VM (Fly.io), no PC required.

PER-PROJECT layout: every project is a folder under the Submittal Tracker
base containing a "Submittals Inbox" subfolder and its own workbook:

    Submittal Tracker/
      123 Main Street/
        123 Main Street - Submittal Log.xlsx
        Submittals Inbox/          <- PMs drop PDFs here

Projects are AUTO-DISCOVERED each poll — adding a project is just creating
that folder structure, no redeploy. Every PDF dropped in a project's inbox
is logged to that project (the folder wins over whatever address the
document shows). Hyperlinks are Dropbox share links; the log upload is
revision-guarded; processed-file state lives in each project's inbox.
"""
import json
import os
import tempfile
import time
import traceback
from datetime import datetime, timezone

import anthropic
import dropbox
from dropbox.files import FileMetadata, FolderMetadata, WriteMode

from dbx_client import make_client
from extractor import extract_submittal
from excel_writer import append_submittal

# Set DBX_BASE to the Dropbox path of your "Submittal Tracker" folder.
BASE = os.environ.get("DBX_BASE", "/Your Dropbox Folder/Submittal Tracker")
INBOX_NAME = "Submittals Inbox"
POLL_SECONDS = int(os.environ.get("POLL_SECONDS", "30"))

_states = {}      # project -> processed-file state (mirrors Dropbox copy)
_failures = {}    # (project, key) -> attempt count


def log(msg):
    print(f"[{datetime.now(timezone.utc):%Y-%m-%d %H:%M:%S}Z] {msg}", flush=True)


def list_projects(dbx):
    """Folders under BASE that contain a 'Submittals Inbox' subfolder."""
    projects = []
    for e in dbx.files_list_folder(BASE).entries:
        if not isinstance(e, FolderMetadata) or e.name.startswith(("_", ".")):
            continue
        try:
            dbx.files_get_metadata(f"{BASE}/{e.name}/{INBOX_NAME}")
            projects.append(e.name)
        except dropbox.exceptions.ApiError:
            continue
    return projects


class NamespaceLockedError(RuntimeError):
    """Dropbox team-space write lock stayed busy through every retry.
    TRANSIENT — must not count toward a file's 3-strike failure limit."""


def _write_retry(fn, what, attempts=8):
    """Run a Dropbox WRITE with backoff — a Dropbox Business team space is a
    shared namespace with a global write lock, so writes routinely lose races
    against the whole team's sync traffic (too_many_write_operations)."""
    for i in range(attempts):
        try:
            return fn()
        except dropbox.exceptions.RateLimitError as e:
            wait = min(60, max(getattr(e.error, "retry_after", 1) or 1, 2) * (1.6 ** i))
            log(f"  .. Dropbox write contention on {what} — waiting {wait:.0f}s")
            time.sleep(wait)
    raise NamespaceLockedError(
        f"Dropbox namespace still locked after {attempts} tries ({what})")


def load_state(dbx, project):
    if project not in _states:
        try:
            _, resp = dbx.files_download(f"{BASE}/{project}/{INBOX_NAME}/.processed_cloud.json")
            _states[project] = json.loads(resp.content)
        except Exception:
            _states[project] = {}
    return _states[project]


def save_state(dbx, project, state):
    _states[project] = state
    _write_retry(
        lambda: dbx.files_upload(
            json.dumps(state, indent=1).encode(),
            f"{BASE}/{project}/{INBOX_NAME}/.processed_cloud.json",
            mode=WriteMode.overwrite),
        "state file")


def share_link(dbx, path):
    """A permanent Dropbox share URL for the PDF (or None if not allowed)."""
    try:
        return _write_retry(
            lambda: dbx.sharing_create_shared_link_with_settings(path).url,
            "share link")
    except dropbox.exceptions.ApiError:
        try:  # link probably exists already
            links = dbx.sharing_list_shared_links(path=path, direct_only=True).links
            if links:
                return links[0].url
        except Exception:
            pass
    except RuntimeError:
        pass
    return None


def list_inbox(dbx, inbox):
    entries = []
    res = dbx.files_list_folder(inbox)
    entries.extend(res.entries)
    while res.has_more:
        res = dbx.files_list_folder_continue(res.cursor)
        entries.extend(res.entries)
    return [e for e in entries
            if isinstance(e, FileMetadata) and e.name.lower().endswith(".pdf")]


def update_log(dbx, log_path, data, pdf_name, link_url):
    """Download-append-upload with a revision guard (3 attempts)."""
    for attempt in range(3):
        with tempfile.TemporaryDirectory() as td:
            xlsx = os.path.join(td, "log.xlsx")
            rev = None
            try:
                meta = dbx.files_download_to_file(xlsx, log_path)
                rev = meta.rev
            except dropbox.exceptions.ApiError:
                pass  # first submittal for this project — workbook gets created
            sheet = append_submittal(xlsx, data, pdf_filename=pdf_name,
                                     link_url=link_url)
            mode = WriteMode.update(rev) if rev else WriteMode.add
            try:
                with open(xlsx, "rb") as f:
                    payload = f.read()
                _write_retry(
                    lambda: dbx.files_upload(payload, log_path, mode=mode),
                    "log workbook")
                return sheet
            except dropbox.exceptions.ApiError as e:
                if "conflict" in str(e).lower() and attempt < 2:
                    log("  .. log changed while writing (a PM saved it) — "
                        "re-applying on the fresh copy")
                    continue
                raise
    raise RuntimeError("could not write the log after 3 attempts")


def process_project(dbx, project):
    inbox = f"{BASE}/{project}/{INBOX_NAME}"
    log_path = f"{BASE}/{project}/{project} - Submittal Log.xlsx"
    state = load_state(dbx, project)

    for e in sorted(list_inbox(dbx, inbox), key=lambda x: x.name):
        key = f"{e.name}|{e.content_hash[:16]}"
        if key in state:
            continue
        if "#" in e.name:
            new_name = e.name.replace("#", "_")
            log(f"[{project}] Renaming '{e.name}' -> '{new_name}' "
                f"(Excel can't link to filenames with #)")
            try:
                dbx.files_move_v2(e.path_display, f"{inbox}/{new_name}")
            except dropbox.exceptions.ApiError as ex:
                log(f"  rename failed ({ex.error}); skipping file")
                state[key] = {"error": "unrenameable # in filename"}
                save_state(dbx, project, state)
            continue
        log(f"[{project}] Processing: {e.name}")
        try:
            with tempfile.TemporaryDirectory() as td:
                pdf = os.path.join(td, e.name)
                dbx.files_download_to_file(pdf, e.path_display)
                data = extract_submittal(pdf)
            data["project"] = project       # the folder decides the project
            link = share_link(dbx, e.path_display)
            sheet = update_log(dbx, log_path, data, e.name, link)
            state[key] = {"when": datetime.now(timezone.utc).isoformat(),
                          "sheet": sheet,
                          "number": data.get("submittal_number"),
                          "title": data.get("title")}
            save_state(dbx, project, state)
            log(f"  -> logged to '{sheet}': "
                f"#{data.get('submittal_number')}  {data.get('title')}")
        except (anthropic.RateLimitError, anthropic.InternalServerError,
                anthropic.APIConnectionError,
                dropbox.exceptions.RateLimitError,
                dropbox.exceptions.InternalServerError,
                NamespaceLockedError) as ex:
            log(f"  .. API busy ({ex.__class__.__name__}) — 30s wait, will retry.")
            log(f"     detail: {str(ex)[:300]}")
            time.sleep(30)
        except Exception as ex:
            fk = (project, key)
            n = _failures.get(fk, 0) + 1
            _failures[fk] = n
            log(f"  !! FAILED on {e.name} (attempt {n}/3): {ex}")
            if n >= 3:
                log(traceback.format_exc(limit=3))
                state[key] = {"when": datetime.now(timezone.utc).isoformat(),
                              "error": str(ex)}
                save_state(dbx, project, state)
                log(f"  !! giving up on {e.name} — fix the file and re-drop it.")


def main():
    dbx = make_client()
    known = set()
    log(f"cloud watcher up — per-project mode, base {BASE} (poll {POLL_SECONDS}s)")
    while True:
        try:
            projects = list_projects(dbx)
            for p in projects:
                if p not in known:
                    known.add(p)
                    log(f"watching project: {p}")
                process_project(dbx, p)
        except Exception as ex:
            log(f"watcher loop error: {ex}")
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
