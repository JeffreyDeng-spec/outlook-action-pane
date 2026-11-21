"""Production scanner backed by the existing Outlook delta and classification code."""

from contextlib import closing
from datetime import datetime, timezone
import json
import sqlite3
from urllib.parse import parse_qs, quote, unquote, urlsplit
from uuid import uuid4

from action_items import ensure_actions, needs_full_text, open_actions, prepare_replies, record_action, update_action
from event_sync import collect_folder, event_tables, process_queue
from outlook_triage import (
    CLASSIFICATION_FIELDS,
    FOLDERS,
    GRAPH,
    IMMUTABLE_ID,
    PREVIEW_COLUMNS,
    PreviewError,
    apply_moves,
    classify,
    preview_cells,
    run_lock,
    write_native_preview,
)
from windows_jev_guard import local_review
from windows_production_grants import ProductionTrialAdapter
from windows_worker import Worker


MESSAGE_FIELDS = set(("id,receivedDateTime,parentFolderId," + CLASSIFICATION_FIELDS).split(","))
CHECK_FIELDS = {"id", "parentFolderId", "changeKey"}
RECEIPT_FIELDS = {"id", "parentFolderId"}


def request_path(url):
    parsed = urlsplit(url)
    parts = parsed.path.split("/")[1:]
    if (parsed.scheme, parsed.netloc) != ("https", "graph.microsoft.com") or (
        parsed.fragment or parts[:2] != ["v1.0", "me"]
    ):
        raise PreviewError("Unexpected Graph production endpoint")
    return parts[2:], parse_qs(parsed.query, keep_blank_values=True)


def is_quoted_id(value):
    decoded = unquote(value)
    return bool(decoded) and decoded not in {".", ".."} and quote(decoded, safe="") == value


class ProductionAdapter(ProductionTrialAdapter):
    """Use only the known account, the two monitored folders and immutable message IDs."""

    def __init__(self, account, store, client_id, write_allowed, **options):
        super().__init__(account, store, client_id, **options)
        self.write_allowed = write_allowed

    def read(self, url, label, *, token=None, headers=None):
        path, query = request_path(url)
        if set(query) != {"$select"} or len(query["$select"]) != 1:
            raise PreviewError("Unexpected Graph message projection")
        fields = set(query["$select"][0].split(","))
        if len(path) == 2 and path[0] == "mailFolders" and fields == {"id"}:
            if path[1] in self.folder_ids:
                return {"id": self.folder_ids[path[1]]}
        if headers != IMMUTABLE_ID:
            raise PreviewError("Immutable message IDs must be requested")
        message = len(path) == 2 and path[0] == "messages" and is_quoted_id(path[1])
        folder_item = (
            len(path) == 4 and path[0] == "mailFolders" and path[1] in FOLDERS
            and path[2] == "messages" and is_quoted_id(path[3])
        )
        if not ((message and fields in (MESSAGE_FIELDS, MESSAGE_FIELDS | {"webLink"},
                                        {"id", "parentFolderId", "uniqueBody"}, RECEIPT_FIELDS)) or (
            folder_item and fields == CHECK_FIELDS
        )):
            raise PreviewError("Unexpected Graph production read")
        if "uniqueBody" in fields:
            headers = {"Prefer": IMMUTABLE_ID["Prefer"] + ', outlook.body-content-type="text"'}
        return self.http(url, label, token=token, headers=headers)

    def create_reply(self, ident, comment, token):
        if not self.write_allowed():
            raise PreviewError("Windows action writer is paused")
        return self.http(f"{GRAPH}/me/messages/{quote(ident, safe='')}/createReply", "Graph reply draft",
                         token=token, headers=IMMUTABLE_ID, payload={"comment": comment})

    def capture_action(self, db, message, row, token, *, historical=False):
        if needs_full_text(message, row):
            url = f"{GRAPH}/me/messages/{quote(message['id'], safe='')}?%24select=id,parentFolderId,uniqueBody"
            try:
                expanded = self.read(url, "Graph action evidence", token=token, headers=IMMUTABLE_ID)
            except PreviewError as error:
                if error.http_status != 404:
                    raise
                return
            if (expanded.get("id") != message["id"] or
                    expanded.get("parentFolderId") != self.folder_ids["inbox"]):
                return
            message = {**message, "uniqueBody": expanded.get("uniqueBody")}
        record_action(db, message, row, historical=historical)

    def write(self, url, label, *, token=None, headers=None, payload=None):
        path, query = request_path(url)
        if (
            len(path) != 5 or path[0] != "mailFolders" or path[1] not in FOLDERS
            or path[2] != "messages" or not is_quoted_id(path[3]) or path[4] != "move"
            or query or headers != IMMUTABLE_ID or not isinstance(payload, dict)
            or set(payload) != {"destinationId"} or payload["destinationId"] not in FOLDERS
            or payload["destinationId"] == path[1]
        ):
            raise PreviewError("Only moves between Inbox and Junk are permitted")
        if not self.write_allowed():
            raise PreviewError("Windows has not been given mail move responsibility")
        return self.http(url, label, token=token, headers=headers, payload=payload)


class ProductionWorker(Worker):
    def __init__(self, root, adapters, key, write_allowed, *, classifier=classify):
        super().__init__(root, adapters, pilot=True, folder_bound=True)
        self.key = key
        self.write_allowed = write_allowed
        self.classifier = classifier

    def actions(self, account):
        with closing(sqlite3.connect((self.directory(account) / "decisions.sqlite3").as_uri() + "?mode=rw", uri=True)) as db:
            ensure_actions(db)
            last = db.execute("SELECT value FROM event_settings WHERE key='last_action_scan'").fetchone()
            return {"account": account, "items": open_actions(db),
                    "last_sync": last[0] if last else "",
                    "sync_ok": self.status(account)["push"] == "online" and self.last_scan_ok[account] is True}

    def update_action(self, account, ident, command):
        if not self.write_allowed():
            raise PreviewError("Windows action writer is paused")
        with closing(sqlite3.connect((self.directory(account) / "decisions.sqlite3").as_uri() + "?mode=rw", uri=True)) as db:
            ensure_actions(db)
            update_action(db, ident, command)
        return {"account": account, "id": ident, "command": command}

    def backfill_actions(self, db, account, token):
        # Adopt existing checked IDs independently of the move queue and Graph cursors.
        adapter = self.adapters[account]
        with db:
            first = db.execute("INSERT OR IGNORE INTO event_settings VALUES ('action_backfill_v1','1')").rowcount
            if first:
                db.execute("""INSERT OR IGNORE INTO action_backfill SELECT id FROM event_checked
                            WHERE id NOT IN (SELECT message FROM message_actions)
                            AND id NOT IN (SELECT id FROM mailbox_moves WHERE state='pending')""")
        batch = db.execute("""SELECT b.message,d.result FROM action_backfill b LEFT JOIN decisions d
                           ON d.id=b.message ORDER BY d.received DESC,b.message LIMIT 6""").fetchall()
        for ident, cached in batch:
            result = json.loads(cached) if cached else {}
            message = None
            if result.get("target") != "Junk":
                fields = "id,receivedDateTime,parentFolderId," + CLASSIFICATION_FIELDS + ",webLink"
                try:
                    message = adapter.read(f"{GRAPH}/me/messages/{quote(ident,safe='')}?%24select={fields}",
                                           "Graph historical action", headers=IMMUTABLE_ID, token=token)
                except PreviewError as error:
                    if error.http_status != 404:
                        raise
            with db:
                if message and message.get("id") == ident and message.get("parentFolderId") == adapter.folder_ids["inbox"]:
                    adapter.capture_action(db, message, {**result, "source": "Inbox", "target": "Inbox"},
                                           token, historical=True)
                db.execute("DELETE FROM action_backfill WHERE message=?", (ident,))
        if batch and hasattr(self, "submit"):
            self.submit(account)

    def preview(self, account):
        path = self.directory(account) / "preview.json"
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except FileNotFoundError:
            data = {"window": "", "messages": []}
        return {"account": account, **data}

    def status(self, account):
        result = super().status(account)
        enabled = self.write_allowed()
        message = "正在监听 Inbox 和 Junk" if enabled else "等待启用 Windows 邮件移动"
        if self.last_scan_ok[account] is False:
            message = "上次增量检查未完成，后台会重试"
        return {**result, "message": message, "apply": enabled}

    def publish(self, account, rows):
        current = list(map(preview_cells, rows))
        keys = {(cells[0], cells[3], cells[4]) for cells in current}
        previous = self.preview(account)["messages"]
        current.extend([row[name] for name in PREVIEW_COLUMNS] for row in previous
                       if (row["time"], row["subject"], row["domain"]) not in keys)
        window = f"新增邮件 · {datetime.now().astimezone():%Y-%m-%d %H:%M}"
        write_native_preview(self.directory(account) / "preview.json", window, current[:400])
        self.snapshots[account] = {"revision": uuid4().hex}

    def scan(self, account):
        job = self.jobs[account]
        if not job.acquire(blocking=False):
            return False
        try:
            if not self.write_allowed():
                raise PreviewError("Windows is waiting for a confirmed handover")
            adapter = self.adapters[account]
            directory = self.directory(account)
            database = directory / "decisions.sqlite3"
            if not database.is_file():
                raise PreviewError("A verified account database must be installed first")
            with run_lock(directory / "scan.lock"), closing(sqlite3.connect(database)) as db:
                db.execute("PRAGMA trusted_schema=OFF")
                now = datetime.now(timezone.utc)
                floor = event_tables(db, now)
                ensure_actions(db)
                token = adapter.graph_token()
                for folder in FOLDERS:
                    collect_folder(db, folder, token, floor, now,
                                   read=adapter.folder_reader(folder), recover_from_floor=True)

                def provider(message, key, model):
                    if not self.write_allowed():
                        raise PreviewError("Windows handover ended before classification")
                    return self.classifier(message, key, model)

                def move(db, rows, token):
                    return apply_moves(db, rows, token, read=adapter.read, write=adapter.write)

                _, counts, failure = process_queue(
                    db, token, self.key, floor, read=adapter.read, classifier=provider,
                    mover=move, review=local_review, publish=lambda rows: self.publish(account, rows),
                    record=lambda db, msg, row: adapter.capture_action(db, msg, row, token),
                    restored=lambda db, ident: db.execute(
                        "INSERT OR IGNORE INTO action_backfill(message) VALUES (?)", (ident,)),
                )
                if failure or counts["skipped"]:
                    raise PreviewError("Some Outlook messages must be checked again")
                prepare_replies(db, adapter, token)
                self.backfill_actions(db, account, token)
                with db:
                    db.execute("INSERT OR REPLACE INTO event_settings VALUES ('last_action_scan',?)",
                               (datetime.now(timezone.utc).isoformat(),))
            self.last_scan_ok[account] = True
            return True
        except Exception:
            self.last_scan_ok[account] = False
            raise
        finally:
            job.release()
