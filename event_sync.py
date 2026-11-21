"""Durable Graph delta catch-up, called at startup, reconnect or a new-mail event."""
import argparse
from collections import Counter
from contextlib import closing
from datetime import datetime, timedelta, timezone
import json
import os
from pathlib import Path
import re
import sqlite3
import sys
from urllib.parse import quote, urlencode, urlsplit

from mail_accounts import DEFAULT_ACCOUNT, GRAPH_KEYCHAINS, account_state
from outlook_triage import (
    CLASSIFICATION_FIELDS, FOLDERS, GRAPH, IMMUTABLE_ID, PREVIEW_COLUMNS, STATE_DIR, PreviewError, apply_moves,
    graph_token, init_cache, parse_received, pool, preview_cells, preview_rows, provider_key, request_json,
    run_lock, utcstamp, write_native_preview,
)


def migrate_floor(snapshot, now):
    """Resume just before the last regular scan's upper bound, including its partial minute."""
    if snapshot.is_file():
        info = json.loads(snapshot.read_text())
        timestamp = re.search(r"至 (\d{4}-\d{2}-\d{2} \d{2}:\d{2})", info.get("window", ""))
        if timestamp:
            previous = datetime.strptime(timestamp[1], "%Y-%m-%d %H:%M").astimezone(timezone.utc)
            return min(now, previous) - timedelta(minutes=2)
    return now


def event_tables(db, initial_floor):
    init_cache(db)
    db.executescript("""
        CREATE TABLE IF NOT EXISTS event_settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS event_cursor (
            folder TEXT PRIMARY KEY, link TEXT NOT NULL, started_utc TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS event_checked (id TEXT PRIMARY KEY);
        CREATE TABLE IF NOT EXISTS event_queue (id TEXT PRIMARY KEY, received TEXT NOT NULL);
    """)
    with db:
        inserted = db.execute("INSERT OR IGNORE INTO event_settings VALUES ('floor', ?)",
                              (utcstamp(initial_floor),)).rowcount
        if inserted:
            db.execute("INSERT OR IGNORE INTO event_checked SELECT id FROM decisions")
    return db.execute("SELECT value FROM event_settings WHERE key='floor'").fetchone()[0]


def first_link(folder, since):
    params = urlencode({"$select": "id,receivedDateTime",
                        "$filter": f"receivedDateTime ge {since}"})
    return f"{GRAPH}/me/mailFolders/{folder}/messages/delta?{params}"


def traverse(link, token, read):
    candidates = {}
    seen = set()
    while link:
        origin = urlsplit(link)
        if (origin.scheme, origin.netloc) != ("https", "graph.microsoft.com") or link in seen:
            raise PreviewError("Microsoft Graph 增量分页地址无效")
        seen.add(link)
        result = read(link, "Microsoft Graph 增量同步", token=token, headers=IMMUTABLE_ID)
        if not isinstance(result.get("value"), list):
            raise PreviewError("Microsoft Graph 增量响应无效")
        for item in result["value"]:
            if item.get("id") and "@removed" not in item:
                candidates[item["id"]] = item.get("receivedDateTime", "")
        link = result.get("@odata.nextLink")
    cursor = result.get("@odata.deltaLink")
    if not cursor:
        raise PreviewError("Microsoft Graph 没有返回增量同步位置")
    return candidates, cursor


def collect_folder(db, folder, token, floor, now, read=request_json, *, recover_from_floor=False):
    saved = db.execute("SELECT link, started_utc FROM event_cursor WHERE folder=?", (folder,)).fetchone()
    if saved:
        try:
            candidates, cursor = traverse(saved[0], token, read)
        except PreviewError as error:
            if error.http_status != 410:
                raise
            bound = parse_received(floor)
            if not recover_from_floor:
                bound = max(bound, parse_received(saved[1]) - timedelta(minutes=2))
            reseed = max(bound, now - timedelta(minutes=2))
            candidates, cursor = traverse(first_link(folder, utcstamp(reseed)), token, read)
            for item in pool(token, bound, now, metadata_only=True, folders={folder: FOLDERS[folder]},
                             request=read):
                candidates.setdefault(item["id"], item["receivedDateTime"])
    else:
        candidates, cursor = traverse(first_link(folder, floor), token, read)
    with db:
        db.executemany("""
            INSERT OR IGNORE INTO event_queue
            SELECT ?, ? WHERE NOT EXISTS (SELECT 1 FROM event_checked WHERE id=?)
        """, [(ident, received, ident) for ident, received in candidates.items()
              if not received or parse_received(received) >= parse_received(floor)])
        db.execute("INSERT OR REPLACE INTO event_cursor VALUES (?, ?, ?)", (folder, cursor, utcstamp(now)))


def complete(db, ident, *, record=None, message=None, row=None):
    with db:
        if record and message is not None:
            record(db, message, row)
        db.execute("INSERT OR IGNORE INTO event_checked VALUES (?)", (ident,))
        db.execute("DELETE FROM event_queue WHERE id=?", (ident,))


def process_queue(db, token, key, floor, *, read=request_json, classifier=None,
                  mover=apply_moves, review=None, publish=None, record=None, restored=None):
    folders = {read(f"{GRAPH}/me/mailFolders/{folder}?%24select=id", "Microsoft Graph", token=token)["id"]: name
               for folder, name in FOLDERS.items()}
    receipts = {}
    # These IDs are never reclassified or POSTed again. Pending receipts are
    # scanned independently so an orphaned/migrated queue row cannot hide one.
    for ident, prior, target, started in db.execute(
            "SELECT id, source, target, started FROM mailbox_moves WHERE state='pending' ORDER BY started, id"):
        address = f"{GRAPH}/me/messages/{quote(ident, safe='')}"
        observed = None
        reason = "Review move: source or destination not yet verified."
        try:
            proof = read(f"{address}?%24select=id,parentFolderId", "Microsoft Graph", token=token,
                         headers=IMMUTABLE_ID)
        except PreviewError as error:
            if error.http_status == 404:
                reason = "Review move: message not found (HTTP 404)."
            else:
                reason = "Review move: Graph could not verify current folder."
        else:
            if proof.get("id") != ident:
                reason = "Review move: missing or mismatched immutable message ID."
            else:
                observed = folders.get(proof.get("parentFolderId"))
                if observed == target:
                    reason = "Move verified in intended destination."
                    if target == "Inbox" and restored:
                        # Reconcile the action from the real mail, independent of the receipt's placeholder.
                        with db:
                            restored(db, ident)
                elif observed == prior:
                    reason = "Review move: message remains in the original folder."
                else:
                    reason = "Review move: current folder is missing or outside the monitored folders."
        receipts[ident] = {"id": ident, "source": observed or "Unknown", "target": "review",
                           "reason": reason, "received": started, "subject": "(move receipt)",
                           "domain": "", "change_key": None}
    emails = []
    for (ident,) in db.execute("SELECT id FROM event_queue ORDER BY received, id").fetchall():
        if ident in receipts:
            continue
        address = f"{GRAPH}/me/messages/{quote(ident, safe='')}"
        fields = f"%24select=id,receivedDateTime,parentFolderId,{CLASSIFICATION_FIELDS}"
        if record:
            fields += ",webLink"
        try:
            email = read(f"{address}?{fields}", "Microsoft Graph", token=token, headers=IMMUTABLE_ID)
        except PreviewError as error:
            if error.http_status == 404:
                complete(db, ident)
                continue
            raise
        if email.get("id") != ident:
            raise PreviewError("Missing or mismatched immutable message ID.")
        source = folders.get(email.get("parentFolderId"))
        if not source or parse_received(email["receivedDateTime"]) < parse_received(floor):
            complete(db, ident)
            continue
        emails.append({**email, "source": source})
    options = {"classifier": classifier} if classifier else {}
    rows, _, failure = preview_rows(db, emails, key, "jev-latest", review=review, **options)
    checked = list(receipts.values()) + [row for row in rows if row["target"] != "pending"]
    counts = mover(db, checked, token) if checked else Counter()
    if publish and checked:
        publish(checked)
    by_id = {mail["id"]: mail for mail in emails} if record else {}
    for row in checked:
        if not row.get("retry_move"):
            complete(db, row["id"], record=record, message=by_id.get(row["id"]), row=row)
    return checked, counts, failure


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--state-dir", type=Path, default=STATE_DIR)
    parser.add_argument("--seed-only", action="store_true")
    parser.add_argument("--account", choices=GRAPH_KEYCHAINS, default=DEFAULT_ACCOUNT)
    args = parser.parse_args()
    os.umask(0o077)
    folder = account_state(args.state_dir.expanduser(), args.account)
    folder.mkdir(mode=0o700, parents=True, exist_ok=True)
    now = datetime.now(timezone.utc)
    snapshot = folder / "preview.json"
    with run_lock(folder / "scan.lock"), closing(sqlite3.connect(folder / "decisions.sqlite3")) as db:
        floor = event_tables(db, migrate_floor(snapshot, now))
        token = graph_token(args.account, GRAPH_KEYCHAINS[args.account], write=not args.seed_only)
        for name in FOLDERS:
            collect_folder(db, name, token, floor, now)
        pending = db.execute("""SELECT count(*) FROM (
            SELECT id FROM event_queue UNION SELECT id FROM mailbox_moves WHERE state='pending'
        )""").fetchone()[0]
        if args.seed_only:
            print(f"增量位置已保存 · 待检查 {pending} 封", flush=True)
            return 0
        if not pending:
            print("事件分拣 · 没有未检查的新邮件", flush=True)
            return 0
        key = provider_key()
        if not key:
            raise PreviewError("缺少 Jev API key")
        rows, moves, failure = process_queue(db, token, key, floor)
        if rows:
            window = f"新增邮件 · {datetime.now().astimezone():%Y-%m-%d %H:%M}"
            latest = list(map(preview_cells, rows))
            keys = {(cell[0], cell[3], cell[4]) for cell in latest}
            previous = json.loads(snapshot.read_text())["messages"] if snapshot.is_file() else []
            latest.extend([row[name] for name in PREVIEW_COLUMNS] for row in previous
                          if (row["time"], row["subject"], row["domain"]) not in keys)
            write_native_preview(snapshot, window, latest[:400])
        print(f"事件分拣 · 检查 {len(rows)} 封 · 移入 Junk {moves['Junk']} 封 · "
              f"恢复 Inbox {moves['Inbox']} 封 · 待核对 {moves['unconfirmed']} 封", flush=True)
        if failure:
            raise PreviewError(failure)
        if moves["skipped"]:
            raise PreviewError("部分邮件状态发生变化，稍后补查")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except PreviewError as error:
        print(error, file=sys.stderr, flush=True)
        sys.exit(1)
