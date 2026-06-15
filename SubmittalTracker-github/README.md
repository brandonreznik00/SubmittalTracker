# SubmittalTracker

Drop a construction **submittal PDF** into a shared Dropbox folder and it gets read by Claude and logged automatically into a per-project Excel tracker — submittal number, trade, contractor, dates, lead time, and the reviewer approval stamp captured verbatim. Resubmittals update the existing row in place and move the old version to a "Revision History" sheet.

Project managers only need access to the shared Dropbox folder — they drop a PDF, a row appears. No software to install on their end.

## Screenshots

<!-- Drop a screenshot of a filled-in submittal log here, e.g.:
![Submittal log](docs/screenshot-log.png)
A before/after (PDF on the left, logged row on the right) reads especially well. -->

_Coming soon._

## How it works

```
Submittal Tracker/
  123 Main Street/
    123 Main Street - Submittal Log.xlsx     <- the per-project log
    Submittals Inbox/                         <- PMs drop PDFs here
  456 Oak Avenue/
    456 Oak Avenue - Submittal Log.xlsx
    Submittals Inbox/
```

1. A watcher polls each project's `Submittals Inbox`.
2. New PDFs are sent to Claude (`claude-opus-4-8`), which extracts the submittal fields as JSON.
3. The row is appended to that project's workbook, with the Submittal cell hyperlinked to the PDF.

Projects are **auto-discovered** — to add one, just create the folder + `Submittals Inbox` subfolder. No redeploy.

There are two ways to run it:

| Mode | Entry point | Needs | Best for |
|------|-------------|-------|----------|
| **Local PC** | `app/watcher.py` (via `run.ps1`) | A PC that stays on; watches a synced Dropbox folder on disk | Trying it out, single machine |
| **Cloud (Fly.io)** | `app/cloud_watcher.py` | Dropbox API app + Fly.io account | Always-on, no PC required |

## Configuration (placeholders to fill in)

Nothing is hardcoded — all credentials come from environment variables, and folder paths come from `config.json` / `fly.toml`. Replace these before running:

| Placeholder | Where | What to put |
|-------------|-------|-------------|
| `ANTHROPIC_API_KEY` | env var | Your Anthropic API key (`sk-ant-...`) |
| `config.json` paths | local mode | Your Dropbox folder paths (copy from `config.example.json`) |
| `DBX_BASE` | `fly.toml` / env var | Dropbox path to your Submittal Tracker folder, e.g. `/Your Dropbox Folder/Submittal Tracker` |
| `DROPBOX_APP_KEY` / `DROPBOX_APP_SECRET` / `DROPBOX_REFRESH_TOKEN` | env var (cloud only) | From your Dropbox API app |
| `app` name in `fly.toml` | cloud only | A globally-unique Fly.io app name |

## Local PC setup

```powershell
# 1. Configure
copy config.example.json config.json   # then edit the folder paths

# 2. Set your API key (once)
setx ANTHROPIC_API_KEY "sk-ant-..."

# 3. Run (creates the venv on first run)
.\run.ps1
```

See `HOW TO RUN.txt` for the PM-facing notes.

## Cloud setup (Fly.io, always-on)

1. **Dropbox app** — create one at https://www.dropbox.com/developers/apps with full-dropbox scope, then generate a refresh token. You'll have an app key, app secret, and refresh token.
2. **Edit `fly.toml`** — set a unique `app` name and your `DBX_BASE`.
3. **Set secrets and deploy:**

   ```sh
   fly secrets set ANTHROPIC_API_KEY=sk-ant-... \
     DROPBOX_APP_KEY=... DROPBOX_APP_SECRET=... DROPBOX_REFRESH_TOKEN=...
   fly deploy --ha=false
   ```

4. **Logs:** `fly logs`

The cloud watcher polls via the Dropbox API, so the folder doesn't need to be synced anywhere. Hyperlinks in the log are real Dropbox share URLs.

## Rate limits

`extractor.py` paces requests to stay under your account's input-tokens-per-minute limit. Raise `SUBMITTAL_TPM_BUDGET` (env var / `fly.toml`) if your Anthropic tier grows — higher tiers can read the full 8 pages of big shop-drawing sets without pacing.

## Notes / gotchas baked in

- **Cover-page detection** — if page 1 is a transmittal/cover sheet, only that page is sent (≈1/8 the cost).
- **Bluebeam PDFs** — some Bluebeam Stapler exports the API can't parse are automatically retried as page images.
- **Filenames with `#`** — auto-renamed to `_` (Excel can't hyperlink to a `#`).
- **Team-space write lock** — a Dropbox Business shared namespace serializes writes; the cloud watcher retries with backoff.

## Layout

```
app/
  watcher.py        # local-PC watcher (filesystem)
  cloud_watcher.py  # Fly.io watcher (Dropbox API)
  extractor.py      # Claude PDF -> JSON extraction
  excel_writer.py   # writes/updates the per-project Excel log
  dbx_client.py     # Dropbox refresh-token auth + team-space root
config.example.json # copy to config.json for local mode
Dockerfile, fly.toml
Submittal Log - Blank Template.xlsx
```

## License

No license specified — add one if you want others to reuse it.
