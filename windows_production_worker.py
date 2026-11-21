"""Production scanner backed by the existing Outlook delta and classification code."""

from contextlib import closing
from datetime import datetime, timezone
import json
import re
import sqlite3
import unicodedata
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

# An opted-in program must itself be the subject of the change. A member's own
# reservation, balance or tier can change independently of the program.
PROGRAM_CHANGE = r"(?:会员计划|會員計劃)?(?:全新升级|全面升级|焕新升级|计划更新|规则更新)"
PERSONAL_EVENT = re.compile(
    r"(?:您的|你的|本次)(?:酒店)?(?:预订|订单|入住|退款|账单|付款|会籍|会员等级|积分即将到期)"
    r"|(?:预订|订单|确认)(?:号|编号)|预订确认|入住确认|积分到期|验证码|安全提醒"
    r"|(?:reservation|booking|confirmation)\s*(?:number|code|confirmed|#)|your (?:stay|reservation|refund|invoice)"
    r"|请(?:您)?(?:确认预订|支付|取消预订|验证身份)", re.I)


def program_announcement(message, programs, adapter, token):
    sender = message.get("from", {}).get("emailAddress", {}).get("address", "").casefold()
    program = programs.get(sender)
    if not program:
        return None
    name = re.escape(program)
    subject = unicodedata.normalize("NFKC", message.get("subject", "")).strip()
    if not re.fullmatch(r"(?:20\d\d\s*年\s*[,，:：]?\s*)?" + name + PROGRAM_CHANGE, subject):
        return None
    ident = message["id"]
    url = f"{GRAPH}/me/messages/{quote(ident, safe='')}?%24select=id,parentFolderId,uniqueBody"
    try:
        expanded = adapter.read(url, "Graph program announcement", token=token, headers=IMMUTABLE_ID)
    except PreviewError:
        return None
    if expanded.get("id") != ident or expanded.get("parentFolderId") != message.get("parentFolderId"):
        return None
    body = expanded.get("uniqueBody") or {}
    content = body.get("content")
    if body.get("contentType", "").lower() != "text" or not isinstance(content, str) or len(content) > 50000:
        return None
    text = unicodedata.normalize("NFKC", content)
    if PERSONAL_EVENT.search(text):
        return None
    if re.search(name + r"(?:会员计划)?(?:将于|将|于)[^。！？\n]{0,36}?(?:迎来|实施|推出)?" + PROGRAM_CHANGE, text):
        return {"target": "Junk", "reason": "已选择过滤此会员计划的整体更新公告",
                "model": "local-program-announcement", "scores": [], "signals": {}}
    return None


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
        if row.get("model") == "local-no-provider":
            saved = db.execute("SELECT result FROM decisions WHERE id=?", (message["id"],)).fetchone()
            if saved:
                cached = json.loads(saved[0])
                scores = cached.get("scores", [])
                if cached.get("target") == "Junk" or (scores and scores[0][0] in {"junk", "marketing"}):
                    return
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
                    "backfilling": bool(db.execute("SELECT 1 FROM action_backfill LIMIT 1").fetchone()),
                    "sync_ok": self.status(account)["push"] == "online" and self.last_scan_ok[account] is True}

    def update_action(self, account, ident, command):
        if not self.write_allowed():
            raise PreviewError("Windows action writer is paused")
        with closing(sqlite3.connect((self.directory(account) / "decisions.sqlite3").as_uri() + "?mode=rw", uri=True)) as db:
            ensure_actions(db)
            visible = update_action(db, ident, command)
        return {"account": account, "id": ident, "command": command, "visible": visible}

    def backfill_actions(self, db, account, token):
        # Adopt existing checked IDs independently of the move queue and Graph cursors.
        adapter = self.adapters[account]
        with db:
            first = db.execute("INSERT OR IGNORE INTO event_settings VALUES ('action_backfill_v1','1')").rowcount
            if first:
                db.execute("""INSERT OR IGNORE INTO action_backfill SELECT id FROM event_checked
                            WHERE id NOT IN (SELECT message FROM message_actions)
                            AND id NOT IN (SELECT id FROM mailbox_moves WHERE state='pending')""")
            revised = db.execute("INSERT OR IGNORE INTO event_settings VALUES ('assessment_reminder_v3','1')").rowcount
            if revised and not first:
                db.execute("""INSERT OR IGNORE INTO action_backfill SELECT d.id FROM decisions d
                           JOIN event_checked c ON c.id=d.id
                           WHERE json_valid(d.result) AND json_extract(d.result,'$.scores[0][0]') IN ('junk','unclear')
                           AND json_extract(d.result,'$.target')<>'Junk'
                           AND d.id NOT IN (SELECT message FROM message_actions)
                           AND d.id NOT IN (SELECT id FROM mailbox_moves WHERE state='pending')""")
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

                try:
                    programs = json.loads((directory / "program-senders.json").read_text(encoding="utf-8"))
                    if not isinstance(programs, dict) or not all(isinstance(v, str) and v.strip()
                                                                 for v in programs.values()):
                        programs = {}
                except (OSError, ValueError):
                    programs = {}

                def review(message):
                    return local_review(message) or program_announcement(message, programs, adapter, token)

                _, counts, failure = process_queue(
                    db, token, self.key, floor, read=adapter.read, classifier=provider,
                    mover=move, review=review, publish=lambda rows: self.publish(account, rows),
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
