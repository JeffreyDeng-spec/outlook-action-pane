"""Prepare and run the current-account Outlook service after the Mac handover."""

import argparse
from contextlib import closing
import csv
import ctypes
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import secrets
import sqlite3
import stat
import subprocess
import sys
import tempfile
from urllib.parse import parse_qs, urlsplit

if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from mail_accounts import DEFAULT_ACCOUNT, GRAPH_KEYCHAINS, account_state
from outlook_triage import FOLDERS, PREVIEW_COLUMNS, PreviewError, atomic_report, classify, parse_received, utcstamp
from windows_pilot import ListenerSupervisor
from windows_production_grants import production_roots, production_store
from windows_production_worker import ProductionAdapter, ProductionWorker
from windows_worker import Worker, tls_server
from worker_safety import LockBusy, exclusive_lock


PROFILE = Path(os.getenv("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
ROOT = PROFILE / "OutlookTriage"
APP = ROOT / "production-app"
STATE = ROOT / "state"
BRIDGE = ROOT / "bridge"
PORT = 18765
CLIENT_ID = os.getenv("OUTLOOK_TRIAGE_CLIENT_ID", "")
OWNER_SID = os.getenv("OUTLOOK_TRIAGE_OWNER_SID", "")
WHOAMI = str(Path(os.getenv("SystemRoot", "C:/Windows")) / "System32" / "whoami.exe")
OPENSSL = Path(os.getenv("OUTLOOK_TRIAGE_OPENSSL") or "openssl")
JEV_STORE_KEY = ("TYPESAFE_API_KEY", "outlook-triage")
DATABASE = "decisions.sqlite3"
MARKER = "windows-move-owner"
COUNTS = ("event_checked", "event_queue", "mailbox_moves")


def require_owner():
    if not OWNER_SID or not CLIENT_ID:
        raise ValueError("Configure the owner SID and application client ID first")
    if os.name != "nt" or Path(__file__).resolve().parent != APP or PROFILE.resolve(strict=True) != PROFILE:
        raise ValueError("Use the installed Windows package under the current profile")
    reply = subprocess.check_output([WHOAMI, "/user", "/fo", "csv", "/nh"], text=True, timeout=5)
    identity = list(csv.reader(reply.splitlines()))
    if (
        len(identity) != 1 or len(identity[0]) != 2 or identity[0][1] != OWNER_SID
        or ctypes.windll.shell32.IsUserAnAdmin()
    ):
        raise ValueError("Use the existing non-elevated Windows account")


def regular_file(path):
    info = path.lstat()
    if not stat.S_ISREG(info.st_mode) or getattr(info, "st_file_attributes", 0) & 0x400:
        raise ValueError("Expected a regular local file")
    return path


def private_dir(path):
    if not path.is_dir() or path.resolve(strict=True) != path:
        raise ValueError("Expected a regular local directory")
    if getattr(path.lstat(), "st_file_attributes", 0) & 0x400:
        raise ValueError("Reparse points are unavailable for Outlook state")
    return path


def metadata_lock():
    return production_roots(PROFILE)[1] / "worker.lock"


def inspect_backup(db):
    db.execute("PRAGMA trusted_schema=OFF")
    if db.execute("PRAGMA integrity_check").fetchall() != [("ok",)]:
        raise ValueError("SQLite integrity check failed")
    tables = dict(db.execute("SELECT name,type FROM sqlite_schema"))
    required = set(COUNTS) | {"decisions", "event_settings", "event_cursor"}
    if any(tables.get(name) != "table" for name in required) or (
        set(tables.values()) - {"index", "table"}
    ):
        raise ValueError("Unexpected SQLite schema")
    cursors = dict(db.execute("SELECT folder,link FROM event_cursor"))
    if set(cursors) != set(FOLDERS):
        raise ValueError("Backup must have both Graph cursors")
    for link in cursors.values():
        parsed = urlsplit(link)
        if (
            (parsed.scheme, parsed.netloc) != ("https", "graph.microsoft.com")
            or parsed.fragment or "$deltatoken" not in parse_qs(parsed.query)
        ):
            raise ValueError("Backup contains an invalid Graph cursor")
    if not db.execute("SELECT 1 FROM event_settings WHERE key='floor' AND length(value)>0").fetchone():
        raise ValueError("Backup sync start is missing")
    return dict(zip(COUNTS, (db.execute(f"SELECT count(*) FROM {name}").fetchone()[0] for name in COUNTS)))


def import_backups(destination, files):
    if destination.exists() or destination.is_symlink() or set(files) != set(GRAPH_KEYCHAINS):
        raise ValueError("Require two account backups and an unused production state directory")
    sources = {account: regular_file(path) for account, path in files.items()}
    if len({path.resolve(strict=True) for path in sources.values()}) != len(sources):
        raise ValueError("Backup source files must differ")
    private_dir(destination.parent)
    temporary = Path(tempfile.mkdtemp(prefix=".outlook-import-", dir=destination.parent))
    reports = {}
    try:
        for account, path in sources.items():
            directory = account_state(temporary, account)
            directory.mkdir()
            with closing(sqlite3.connect(path.absolute().as_uri() + "?mode=ro&immutable=1", uri=True)) as source:
                report = inspect_backup(source)
                with closing(sqlite3.connect(directory / DATABASE)) as replica:
                    source.backup(replica)
                    if inspect_backup(replica) != report:
                        raise ValueError("Staged copy did not match the source")
            reports[account] = report
        atomic_report(temporary / "import.json", json.dumps({"accounts": reports, "imported": utcstamp(
            datetime.now(timezone.utc)
        )}, separators=(",", ":")))
        temporary.rename(destination)
    finally:
        if temporary.exists():
            for account in sources:
                directory = account_state(temporary, account)
                if directory.exists():
                    (directory / DATABASE).unlink(missing_ok=True)
                    directory.rmdir()
            (temporary / "import.json").unlink(missing_ok=True)
            temporary.rmdir()
    return reports


def verify_live_state(state):
    private_dir(state)
    receipt = json.loads(regular_file(state / "import.json").read_text(encoding="utf-8"))
    if set(receipt["accounts"]) != set(GRAPH_KEYCHAINS):
        raise ValueError("Two-account backup import receipt required")
    for account in GRAPH_KEYCHAINS:
        regular_file(private_dir(account_state(state, account)) / DATABASE)


def write_allowed():
    try:
        marker = regular_file(STATE / MARKER)
        return marker.stat().st_size <= 64 and marker.read_text(encoding="ascii").strip() == CLIENT_ID
    except (OSError, ValueError):
        return False


def bridge_material():
    private_dir(BRIDGE)
    cert = regular_file(BRIDGE / "bridge-cert.pem")
    key = regular_file(BRIDGE / "bridge-key.pem")
    token = regular_file(BRIDGE / "bridge-token").read_text(encoding="ascii").strip()
    return token, cert, key


def prepare_bridge():
    bridge = BRIDGE
    bridge.mkdir(exist_ok=True)
    private_dir(bridge)
    token, cert, key = (bridge / name for name in ("bridge-token", "bridge-cert.pem", "bridge-key.pem"))
    if any(path.exists() for path in (token, cert, key)):
        bridge_material()
        return
    env = dict(os.environ)
    subprocess.run(
        [str(OPENSSL), "req", "-x509", "-newkey", "rsa:2048", "-nodes", "-days", "180",
         "-sha256", "-subj", "/CN=localhost", "-addext", "basicConstraints=critical,CA:FALSE",
         "-addext", "extendedKeyUsage=serverAuth", "-addext", "subjectAltName=DNS:localhost,IP:127.0.0.1",
         "-keyout", str(key), "-out", str(cert)],
        check=True, env=env, capture_output=True, timeout=30,
    )
    with token.open("x", encoding="ascii") as file:
        file.write(secrets.token_urlsafe(32) + "\n")
    bridge_material()


def set_jev(store):
    from tkinter import Tk, messagebox, simpledialog

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        key = simpledialog.askstring(
            "Outlook Triage", "输入 Jev API key。Windows 会发送一封合成测试内容以验证密钥。",
            show="*", parent=root,
        )
        if not key:
            return 2
        message = {
            "from": {"emailAddress": {"address": "sender@example.test"}},
            "subject": "Synthetic Outlook Triage migration test",
            "bodyPreview": "May I confirm the time for the mock project handoff tomorrow?",
            "internetMessageHeaders": [],
        }
        try:
            classify(message, key.strip(), "jev-latest")
            store.set(*JEV_STORE_KEY, key.strip())
        except (OSError, PreviewError, ValueError) as error:
            messagebox.showerror("Outlook Triage", f"密钥验证未完成（{type(error).__name__}）。", parent=root)
            return 1
        messagebox.showinfo("Outlook Triage", "Jev 合成测试成功。密钥已用 Windows DPAPI 加密保存。", parent=root)
        return 0
    finally:
        root.destroy()


class BridgeProbe(Worker):
    def __init__(self, trial_state):
        super().__init__(trial_state, dict.fromkeys(GRAPH_KEYCHAINS), pilot=True, folder_bound=True)
        self.receipt = trial_state / "resident-status.json"

    def status(self, account):
        data = json.loads(self.receipt.read_text(encoding="utf-8"))
        folders = data["accounts"][account]
        fresh = 0 <= (datetime.now(timezone.utc) - parse_received(data["updated_utc"])).total_seconds() <= 30
        return {**super().status(account), "message": "Windows 只读桥接验收",
                "push": "online" if fresh and all(folders[k] for k in ("inbox", "junkemail", "graph")) else "connecting"}

    def preview(self, account):
        row = dict.fromkeys(PREVIEW_COLUMNS, "")
        row["subject"] = f"桥接验收（合成） · {account}"
        return {"account": account, "window": "只读桥接验收", "messages": [row]}

    def scan(self, account):
        raise PreviewError("The bridge probe does not scan or classify messages")


def verify_cutover_ready(store, *, need_state=True):
    if need_state:
        verify_live_state(STATE)
    if any(not store.get(resource, account) for account in GRAPH_KEYCHAINS
           for resource in ("graph-write", "imap-examine")):
        raise ValueError("Verified OAuth grants are missing")
    key = store.get(*JEV_STORE_KEY)
    if not key:
        raise ValueError("The current Windows account needs its Jev key")
    return key, bridge_material()


def run():
    if not write_allowed():
        raise ValueError("A confirmed Mac handover must be recorded first")
    store = production_store(PROFILE)
    with exclusive_lock(metadata_lock()):
        key, (bearer, cert, certificate_key) = verify_cutover_ready(store)
        worker = ProductionWorker(
            STATE, {account: ProductionAdapter(account, store, CLIENT_ID, write_allowed)
                    for account in GRAPH_KEYCHAINS}, key, write_allowed,
        )
        supervisor = ListenerSupervisor(worker)
        worker.submit = supervisor.wake
        with tls_server(worker, bearer, cert, certificate_key, PORT) as server:
            try:
                server.timeout = 1
                supervisor.start()
                while write_allowed():
                    if not all(thread.is_alive() for thread in supervisor.threads):
                        raise RuntimeError("A listener exited")
                    server.handle_request()
            finally:
                if not supervisor.close_before_unlock():
                    raise RuntimeError("Listener stop timed out")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    command = parser.add_subparsers(dest="command", required=True)
    for mode in ("prepare-bridge", "probe-bridge", "set-jev", "check", "run", "disarm"):
        command.add_parser(mode)
    stage = command.add_parser("import")
    stage.add_argument("first_backup", type=Path, help=f"Consistent backup for {DEFAULT_ACCOUNT}")
    stage.add_argument("second_backup", type=Path, help="Consistent backup for the second account")
    arm = command.add_parser("arm")
    arm.add_argument("--mac-writer-stopped", action="store_true", required=True)
    args = parser.parse_args()
    try:
        require_owner()
        if args.command == "prepare-bridge":
            prepare_bridge()
            print("Local bridge materials are ready")
        elif args.command == "probe-bridge":
            token, cert, key = bridge_material()
            with tls_server(BridgeProbe(production_roots(PROFILE)[1]), token, cert, key, PORT) as server:
                server.serve_forever()
        elif args.command == "set-jev":
            return set_jev(production_store(PROFILE))
        elif args.command == "check":
            verify_cutover_ready(production_store(PROFILE), need_state=False)
            print("OAuth, Jev and bridge materials are present")
        elif args.command == "import":
            files = dict(zip(GRAPH_KEYCHAINS, (args.first_backup, args.second_backup)))
            with exclusive_lock(ROOT / "import.lock"):
                print(json.dumps(import_backups(STATE, files)))
        elif args.command == "arm":
            verify_cutover_ready(production_store(PROFILE))
            with exclusive_lock(metadata_lock()):
                with (STATE / MARKER).open("x", encoding="ascii") as file:
                    file.write(CLIENT_ID + "\n")
            print("Windows move authorization armed; start the worker next")
        elif args.command == "disarm":
            (STATE / MARKER).unlink(missing_ok=True)
            print("Windows move authorization withdrawn")
        elif args.command == "run":
            run()
        return 0
    except KeyboardInterrupt:
        return 0
    except (OSError, ValueError, KeyError, LockBusy, PreviewError, RuntimeError, sqlite3.Error,
            subprocess.SubprocessError) as error:
        safe = str(error) if type(error) is ValueError and not isinstance(error, json.JSONDecodeError) else type(error).__name__
        print(f"Outlook migration did not complete: {safe}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
