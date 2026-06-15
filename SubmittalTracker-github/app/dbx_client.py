"""Dropbox client for the cloud watcher — refresh-token auth that never
expires, plus team-space path-root handling (a Dropbox Business team space
shared folder isn't visible to the API without setting the root namespace
explicitly)."""
import os

import dropbox
from dropbox.common import PathRoot


def make_client() -> dropbox.Dropbox:
    dbx = dropbox.Dropbox(
        oauth2_refresh_token=os.environ["DROPBOX_REFRESH_TOKEN"],
        app_key=os.environ["DROPBOX_APP_KEY"],
        app_secret=os.environ["DROPBOX_APP_SECRET"],
        # bound the retries — the SDK default retries rate-limits FOREVER,
        # which turns a transient Dropbox 429 storm into a silent hang
        timeout=60,
        max_retries_on_rate_limit=4,
        max_retries_on_error=4,
    )
    acct = dbx.users_get_current_account()
    root = acct.root_info
    if root.root_namespace_id != root.home_namespace_id:
        dbx = dbx.with_path_root(PathRoot.root(root.root_namespace_id))
    return dbx
