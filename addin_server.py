"""Local HTTPS host for the Outlook message-read task pane."""
import argparse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import os
from pathlib import Path
import secrets
import ssl
import struct
import subprocess
import sys
from threading import Event, Lock, Thread
import zlib
from urllib.parse import parse_qs, urlsplit

from mail_accounts import DEFAULT_ACCOUNT, GRAPH_KEYCHAINS, account_state
from outlook_triage import STATE_DIR, PreviewError, atomic_report, export_existing_preview, keychain, run_lock
from imap_trigger import SERVICE as IMAP_SERVICE, watch_folder
from bridge_client import BridgeClient, BridgeError, action_command

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "addin"
CERT_CONFIG = """[req]
prompt = no
distinguished_name = dn
x509_extensions = leaf
[dn]
CN = localhost
[leaf]
basicConstraints = critical,CA:FALSE
keyUsage = critical,digitalSignature,keyEncipherment
extendedKeyUsage = serverAuth
subjectAltName = DNS:localhost,IP:127.0.0.1
"""


def listen_url(port):
    return f"https://localhost:{port}"


def install_files(state, port):
    state.mkdir(parents=True, exist_ok=True, mode=0o700)
    token_file = state / "access"
    if not token_file.exists():
        atomic_report(token_file, secrets.token_urlsafe(32))
    token = token_file.read_text()
    if not (len(token) == 43 and all(ch.isascii() and (ch.isalnum() or ch in "-_") for ch in token)):
        raise PreviewError("本地组件访问配置无效")
    cert, key = state / "localhost.pem", state / "localhost-key.pem"
    if not cert.is_file() or not key.is_file():
        config = state / "localhost.cnf"
        config.write_text(CERT_CONFIG)
        subprocess.run([
            "/usr/bin/openssl", "req", "-x509", "-newkey", "rsa:2048", "-nodes",
            "-days", "365", "-sha256", "-config", str(config),
            "-keyout", str(key), "-out", str(cert),
        ], check=True, capture_output=True)
    base = f"{listen_url(port)}/p/{token}"
    manifest = (ASSETS / "manifest.xml.in").read_text().replace("__BASE__", base)
    atomic_report(state / "manifest.xml", manifest)
    return cert, key, token


def icon(size):
    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data))

    blue, paper = b"\x00\x5a\x9e", b"\xf3\xf8\xfc"
    rows = []
    for y in range(size):
        row = bytearray(b"\x00")
        for x in range(size):
            nx, ny = x / size, y / size
            inside = .18 <= nx <= .82 and .26 <= ny <= .74
            seam = abs(ny - (.55 - abs(nx - .5) * .77)) < .022
            row.extend(paper if inside and not seam else blue)
        rows.append(row)
    info = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", info) + chunk(b"IDAT", zlib.compress(b"".join(rows))) + chunk(b"IEND", b"")


class PreviewJob:
    def __init__(self, state_dir, event_mode=False, account=DEFAULT_ACCOUNT):
        self.state_dir = state_dir
        self.account = account
        self.event_mode = event_mode
        self.imap_configured = event_mode and bool(keychain(IMAP_SERVICE, account))
        self.online_folders = set()
        self.arrivals = Event()
        reports = account_state(state_dir, account)
        self.snapshot = reports / "preview.json"
        export_existing_preview(reports / "preview.html", self.snapshot)
        self.lock = Lock()
        self.running = False
        self.message = ""

    def status(self):
        try:
            revision = str(self.snapshot.stat().st_mtime_ns)
        except FileNotFoundError:
            revision = "empty"
        with self.lock:
            push = "waiting_auth" if not self.imap_configured else (
                "online" if len(self.online_folders) == 2 else "connecting")
            return {"revision": revision, "running": self.running, "message": self.message,
                    "event_mode": self.event_mode, "push": push if self.event_mode else None}

    def start(self):
        if self.event_mode:
            self.arrivals.set()
            return

        with self.lock:
            if self.running:
                return
            self.running = True
            self.message = "正在扫描 Inbox 和 Junk…"
            Thread(target=self.run, daemon=True).start()

    def connection(self, folder, ready):
        with self.lock:
            if ready:
                self.online_folders.add(folder)
            else:
                self.online_folders.discard(folder)

    def run_events(self, stop):
        delay = 2
        while not stop.is_set():
            if not self.arrivals.wait(2):
                continue
            if stop.wait(2):
                return
            self.arrivals.clear()
            with self.lock:
                self.running = True
                self.message = "正在补查未检查的新邮件…"
            if self.run():
                delay = 2
            elif not stop.wait(delay):
                self.arrivals.set()
                delay = min(delay * 2, 60)

    def run(self):
        message = "扫描失败，请查看本地 Outlook Triage 运行环境。"
        auto_summary = ""
        succeeded = False
        try:
            command = ["event_sync.py"] if self.event_mode else ["outlook_triage.py", "scan"]
            args = [sys.executable, "-B", str(ROOT / command[0]), *command[1:],
                    "--state-dir", str(self.state_dir), "--account", self.account]
            with subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                  text=True, encoding="utf-8", errors="replace") as process:
                for line in process.stdout:
                    line = line.strip()
                    if line.startswith("事件分拣 ·"):
                        auto_summary = line
                    if line.startswith(("本轮新增分类 ", "扫描 ", "事件分拣 ·", "Jev ", "Microsoft ",
                                        "缺少 Jev ", "读不到 Outlook ", "已有分类进程", "部分邮件")):
                        message = line
                        with self.lock:
                            self.message = line
                succeeded = process.wait() == 0
                if succeeded:
                    message = auto_summary or "预览已更新。"
        except OSError:
            message = "无法启动本地扫描程序。"
        finally:
            with self.lock:
                self.running = False
                self.message = message
        return succeeded


def make_handler(job, token, port, jobs=None, *, bridge=None):
    jobs = jobs or {DEFAULT_ACCOUNT: job}
    origin = listen_url(port)
    csp = ("default-src 'none'; script-src 'self' https://appsforoffice.microsoft.com https://officeapis.public.onecdn.static.microsoft; "
           "style-src 'self'; img-src 'self'; connect-src 'self'; base-uri 'none'; "
           "form-action 'none'; frame-ancestors "
           "https://outlook.live.com https://outlook.cloud.microsoft https://outlook.office.com "
           "https://outlook.office365.com")
    assets = {
        "": ("index.html", "text/html; charset=utf-8"),
        "pane.css": ("pane.css", "text/css; charset=utf-8"),
        "pane.js": ("pane.js", "text/javascript; charset=utf-8"),
    }

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *_):
            pass

        def route(self):
            if self.headers.get("Host") != f"localhost:{port}":
                return None
            segments = self.path.partition("?")[0].split("/", 3)
            if len(segments) != 4:
                return None
            prefix, match, supplied, path = segments
            if prefix or match != "p" or not supplied.isascii() or not secrets.compare_digest(supplied, token):
                return None
            return path

        def send(self, body, kind, status=200):
            self.send_response(status)
            self.send_header("Content-Type", kind)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("Referrer-Policy", "no-referrer")
            self.send_header("Content-Security-Policy", csp)
            self.end_headers()
            self.wfile.write(body)

        def json(self, value, status=200):
            self.send(json.dumps(value, ensure_ascii=False).encode(), "application/json; charset=utf-8", status)

        def selected_job(self):
            query = parse_qs(urlsplit(self.path).query, keep_blank_values=True)
            accounts = query.get("account")
            if accounts is None:
                return job
            return jobs.get(accounts[0]) if len(accounts) == 1 else None

        def remote(self, route, account=None):
            try:
                data = bridge.request(route, account)
            except BridgeError:
                self.json({"error": "Windows 分类服务暂不可用"}, 502)
            else:
                self.json(data, 202 if route == "scan" else 200)

        def do_GET(self):
            route = self.route()
            selected = self.selected_job()
            if route == "accounts":
                if bridge:
                    self.remote("accounts")
                else:
                    self.json({"accounts": list(jobs), "default": DEFAULT_ACCOUNT})
            elif route == "actions" and bridge and self.explicit_account() is not None:
                self.remote(route, selected)
            elif route in assets:
                filename, kind = assets[route]
                self.send((ASSETS / filename).read_bytes(), kind)
            elif route == "status" and selected:
                if bridge:
                    self.remote("status", selected)
                else:
                    self.json(selected.status())
            elif route == "preview" and selected:
                if bridge:
                    self.remote("preview", selected)
                    return
                try:
                    report = selected.snapshot.read_bytes()
                except FileNotFoundError:
                    self.json({"window": "", "messages": []})
                else:
                    self.send(report, "application/json; charset=utf-8")
            elif route in ("icon-16.png", "icon-32.png", "icon-64.png", "icon-80.png"):
                self.send(icon(int(route[5:-4])), "image/png")
            else:
                self.json({"error": "Not found"}, 404)

        def explicit_account(self):
            query = parse_qs(urlsplit(self.path).query, keep_blank_values=True)
            values = query.get("account", [])
            if set(query) == {"account"} and len(values) == 1 and values[0] in jobs:
                return values[0]
            return None

        def do_POST(self):
            route = self.route()
            selected = self.selected_job()
            command = action_command(route or "")
            if ((route != "scan" and not command) or not selected or
                    (command and (not bridge or not self.explicit_account()))):
                self.json({"error": "Not found"}, 404)
            elif (self.headers.get_all("Origin") != [origin] or self.headers.get_all("X-Outlook-Triage") != ["1"]
                  or self.headers.get_all("Content-Length", []) not in ([], ["0"]) or self.headers.get("Transfer-Encoding")):
                self.json({"error": "Forbidden"}, 403)
            elif bridge:
                self.remote(route, selected)
            else:
                selected.start()
                self.json(selected.status(), 202)

    return Handler


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=("prepare", "serve", "bridge"))
    parser.add_argument("--state-dir", type=Path, default=STATE_DIR)
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--event-mode", action="store_true", help="catch up by event instead of full scans")
    parser.add_argument("--bridge-url")
    parser.add_argument("--bridge-token-file", type=Path)
    parser.add_argument("--bridge-ca-file", type=Path)
    args = parser.parse_args()
    os.umask(0o077)
    state_dir = args.state_dir.expanduser()
    addin_dir = state_dir / "addin"
    event_mode = args.event_mode or (addin_dir / "event-mode").is_file()
    addin_dir.mkdir(parents=True, exist_ok=True, mode=0o700)
    with run_lock(addin_dir / "prepare.lock"):
        cert, key, token = install_files(addin_dir, args.port)
    if args.command == "prepare":
        print(f"Outlook XML: {addin_dir / 'manifest.xml'}")
        print(f"Local certificate: {cert}")
        return
    bridge = None
    if args.command == "bridge":
        if not args.bridge_url or not args.bridge_token_file:
            parser.error("bridge requires --bridge-url and --bridge-token-file")
        try:
            bridge = BridgeClient(args.bridge_url, args.bridge_token_file, GRAPH_KEYCHAINS,
                                  args.bridge_ca_file)
        except BridgeError as error:
            raise SystemExit(str(error)) from None
        jobs = {account: account for account in GRAPH_KEYCHAINS}
    else:
        accounts = GRAPH_KEYCHAINS if event_mode else (DEFAULT_ACCOUNT,)
        jobs = {account: PreviewJob(state_dir, event_mode, account) for account in accounts}
    handler = make_handler(jobs[DEFAULT_ACCOUNT], token, args.port, jobs, bridge=bridge)
    with ThreadingHTTPServer(("127.0.0.1", args.port), handler) as server:
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        context.minimum_version = ssl.TLSVersion.TLSv1_2
        context.load_cert_chain(cert, key)
        server.socket = context.wrap_socket(server.socket, server_side=True)
        print(f"Outlook Triage listening at {listen_url(args.port)}", flush=True)
        stop = Event()
        if bridge is None and event_mode:
            for account, job in jobs.items():
                Thread(target=job.run_events, args=(stop,), daemon=True).start()
                if job.imap_configured:
                    for folder in ("Inbox", "Junk"):
                        Thread(target=watch_folder, args=(folder, stop, job.arrivals.set,
                                                          job.connection, account), daemon=True).start()
        try:
            server.serve_forever()
        finally:
            stop.set()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
