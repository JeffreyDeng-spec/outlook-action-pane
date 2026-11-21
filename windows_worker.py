"""Authenticated Outlook Triage worker API (Windows mock-only staging entrypoint)."""
from __future__ import annotations
import json, re, secrets, ssl
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlsplit
from threading import Lock
from uuid import uuid4
import sqlite3
from contextlib import closing, contextmanager
from datetime import datetime, timezone

from event_sync import collect_folder, event_tables, process_queue
from bridge_client import action_command
from outlook_triage import apply_moves, GRAPH, PREVIEW_COLUMNS, PreviewError, preview_cells, run_lock
from mail_accounts import GRAPH_KEYCHAINS
from worker_safety import account_state, exclusive_lock, worker_state_root

FOLDERS = ("inbox", "junkemail")
FAKE_ACCOUNTS = ("alpha@example.test","beta@example.test")

class Worker:
    """Graph operations are injected; no implicit credential or network access."""
    def __init__(self, root, adapters, *, apply=False, pilot=False, graph_only=False, folder_bound=False):
        self.root, self.adapters, self.apply = worker_state_root(str(root)), adapters, apply
        self.pilot = pilot
        self.graph_only = graph_only
        self.folder_bound = graph_only or folder_bound
        self.metadata_counts = {}
        if self.folder_bound and (not pilot or apply):
            raise ValueError("Folder-bound read mode requires the read-only pilot guard")
        self.last_scan_ok = {a: None for a in adapters}
        expected = tuple(GRAPH_KEYCHAINS) if pilot else FAKE_ACCOUNTS
        if set(adapters) != set(expected):
            raise ValueError("Unexpected worker accounts")
        if pilot and apply:
            raise ValueError("The live-account pilot cannot apply moves")
        self.jobs = {a: Lock() for a in adapters}
        self.generation = uuid4().hex
        self.snapshots = {}
        self.online = {a: {f: False for f in FOLDERS} for a in adapters}

    def connectivity(self, account, folder, online):
        if folder not in FOLDERS or account not in self.online:
            raise ValueError("unknown event source")
        self.online[account][folder] = bool(online)

    def directory(self, account):
        if self.pilot:
            from mail_accounts import account_state as real_account_state
            return account_state(self.root, account, real_account_state(self.root,account).name)
        return account_state(self.root, account, "alpha" if account == FAKE_ACCOUNTS[0] else "beta")

    def status(self, account):
        folders = dict(self.online[account])
        if self.graph_only:
            ok = self.last_scan_ok[account]
            message = ("Graph metadata only: " + str(self.metadata_counts.get(account, 0)) +
                       " queued IDs; no IMAP, classification or moves" if ok else
                       "Graph metadata check failed" if ok is False else "Graph metadata waiting")
            return {"account": account, "revision": self.generation, "running": self.jobs[account].locked(),
                    "message": message, "event_mode": False, "push": None, "folders": {}, "apply": False}
        return {"account": account, "revision": self.generation + ":" + str(self.snapshots.get(account, {}).get("revision", "empty")),
                "running": self.jobs[account].locked(), "message": (("只读试运行" if self.last_scan_ok[account] is not False else "只读同步暂时失败") if self.pilot else "MOCK only"), "event_mode": True,
                "push": "online" if all(folders.values()) else "connecting", "folders": folders, "apply": self.apply}

    def preview(self, account):
        data = self.snapshots.get(account, {})
        return {"account": account, "window": "Graph metadata only" if self.graph_only else "MOCK",
                "messages": data.get("messages",[])}

    def scan(self, account):
        guard = self.jobs[account]
        if not guard.acquire(blocking=False):
            return False
        try:
            adapter = self.adapters[account]
            directory = self.directory(account)
            directory.mkdir(parents=True, exist_ok=True)
            with run_lock(directory/"scan.lock"), closing(sqlite3.connect(directory/"decisions.sqlite3")) as db:
                now = datetime.now(timezone.utc)
                floor = event_tables(db, now if self.pilot else datetime(2026,1,1,tzinfo=timezone.utc))
                token = adapter.graph_token() if self.pilot else "dummy-token"
                for folder in FOLDERS:
                    reader = adapter.folder_reader(folder) if self.folder_bound else adapter.read
                    collect_folder(db, folder, token, floor, now, read=reader)
                if self.graph_only:
                    self.metadata_counts[account] = db.execute("SELECT count(*) FROM event_queue").fetchone()[0]
                self.last_scan_ok[account] = True
                if self.pilot or not self.apply:
                    # This must never consume queued mail or trigger a move, including via the API.
                    return True

                def mover(db, rows, token):
                    return apply_moves(db, rows, token, read=adapter.read, write=adapter.write)
                rows, _, failure = process_queue(db, "dummy-token", "dummy-key", floor, read=adapter.read,
                                                     classifier=adapter.classify, mover=mover)
                if rows:
                    self.snapshots[account] = {"revision": uuid4().hex,
                       "messages": [dict(zip(PREVIEW_COLUMNS,preview_cells(row))) for row in rows][:400]}
                if failure:
                    raise RuntimeError("mock classification failed")
        except Exception:
            self.last_scan_ok[account] = False
            raise
        finally:
            guard.release()

def handler(worker, bearer):
    if not re.fullmatch(r"[A-Za-z0-9._~-]{32,512}", bearer):
        raise ValueError("Invalid worker bearer format")

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *_):
            pass
        def reply(self, code, payload):
            raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
        def execute(self, method):
            provided = self.headers.get_all("Authorization", [])
            if len(provided)!=1 or not provided[0].isascii() or not secrets.compare_digest(provided[0],"Bearer " + bearer):
                return self.reply(401, {"error": "Unauthorized"})
            parts = urlsplit(self.path)
            if parts.netloc or parts.fragment:
                return self.reply(404, {"error": "Not found"})
            params = parse_qs(parts.query,keep_blank_values=True)
            if method=="GET" and parts.path=="/accounts" and not params:
                return self.reply(200,{"accounts":list(worker.adapters),"default":next(iter(worker.adapters))})
            if set(params)!={"account"} or len(params["account"])!=1 or params["account"][0] not in worker.adapters:
                return self.reply(404,{"error":"Unknown account"})
            account=params["account"][0]
            command = action_command(parts.path[1:]) if parts.path.startswith("/") else None
            if method == "GET" and parts.path == "/actions" and hasattr(worker, "actions"):
                try:
                    return self.reply(200, worker.actions(account))
                except (OSError, sqlite3.Error, PreviewError):
                    return self.reply(503, {"error": "Action list unavailable"})
            if method == "POST" and command and hasattr(worker, "update_action"):
                if self.headers.get_all("Content-Length", []) != ["0"] or self.headers.get("Transfer-Encoding"):
                    return self.reply(400, {"error": "Expected empty request"})
                try:
                    return self.reply(200, worker.update_action(account, *command))
                except (KeyError, ValueError):
                    return self.reply(404, {"error": "Unknown action"})
                except (OSError, sqlite3.Error, PreviewError):
                    return self.reply(503, {"error": "Action list unavailable"})
            if method=="GET" and parts.path=="/status":
                return self.reply(200,worker.status(account))
            if method=="GET" and parts.path=="/preview":
                return self.reply(200,worker.preview(account))
            if method=="POST" and parts.path=="/scan" and self.headers.get_all("Content-Length", []) in ([],["0"]) and not self.headers.get("Transfer-Encoding"):
                try:
                    getattr(worker,"submit",worker.scan)(account)
                except Exception:
                    return self.reply(503,{"error":"Scan could not complete"})
                return self.reply(202,worker.status(account))
            return self.reply(404,{"error":"Not found"})
        def do_GET(self):
            return self.execute("GET")
        def do_POST(self):
            return self.execute("POST")
    return Handler

def tls_server(worker,bearer,certificate,key,port=0):
    if not (certificate and key):
        raise ValueError("Worker API requires TLS")
    server = ThreadingHTTPServer(("127.0.0.1",port),handler(worker,bearer))
    try:
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        ctx.minimum_version = ssl.TLSVersion.TLSv1_2
        ctx.load_cert_chain(certificate,key)
        server.socket = ctx.wrap_socket(server.socket,server_side=True)
    except BaseException:
        server.server_close()
        raise
    return server


@contextmanager
def locked_tls_server(worker, bearer, certificate, key, port=0):
    """Only one staged worker per explicit state root; no Internet listener."""
    with exclusive_lock(worker.root / "worker.lock"):
        with tls_server(worker, bearer, certificate, key, port) as server:
            yield server
