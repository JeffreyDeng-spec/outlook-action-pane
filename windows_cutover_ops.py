"""Read-only proof for freshly imported mail state and consistent rollback exports."""

import argparse
from contextlib import closing
import json
from pathlib import Path
import sqlite3
import sys
from tempfile import mkdtemp

if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from event_sync import traverse
from mail_accounts import GRAPH_KEYCHAINS, account_state
from outlook_triage import FOLDERS, PreviewError
from windows_migration import (
    CLIENT_ID, DATABASE, PROFILE, ROOT, STATE, import_backups, metadata_lock,
    require_owner, regular_file, verify_live_state, write_allowed,
)
from windows_production_grants import ProductionTrialAdapter, production_store
from worker_safety import LockBusy, exclusive_lock


def verify_imported_cursors():
    verify_live_state(STATE)
    store = production_store(PROFILE)
    reports = {}
    for account in GRAPH_KEYCHAINS:
        adapter = ProductionTrialAdapter(account, store, CLIENT_ID)
        token = adapter.graph_token()
        path = regular_file(account_state(STATE, account) / DATABASE)
        with closing(sqlite3.connect(path.as_uri() + "?mode=ro", uri=True)) as db:
            saved = dict(db.execute("SELECT folder,link FROM event_cursor"))
            for folder in FOLDERS:
                traverse(saved[folder], token, adapter.folder_reader(folder))
        reports[account] = dict.fromkeys(FOLDERS, True)
    return reports


def export_rollback_state():
    if write_allowed():
        raise ValueError("Disarm Windows mail moves before exporting rollback state")
    verify_live_state(STATE)
    directory = Path(mkdtemp(prefix="rollback-", dir=ROOT)) / "state"
    import_backups(directory, {account: account_state(STATE, account) / DATABASE for account in GRAPH_KEYCHAINS})
    return directory


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("operation", choices=("verify-imported-cursors", "export-rollback-state"))
    args = parser.parse_args()
    try:
        require_owner()
        with exclusive_lock(metadata_lock()):
            if args.operation == "verify-imported-cursors":
                print(json.dumps(verify_imported_cursors()))
            else:
                print(export_rollback_state())
        return 0
    except (OSError, ValueError, KeyError, PreviewError, LockBusy, sqlite3.Error) as error:
        print(f"Cutover operation stopped ({type(error).__name__})", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
