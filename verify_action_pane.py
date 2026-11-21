"""Exercise the production pane, HTTPS bridge and action store using synthetic mail."""
import argparse
import os
from contextlib import closing, contextmanager
from datetime import datetime, timezone
from http.server import ThreadingHTTPServer
from pathlib import Path
import secrets
import shutil
import sqlite3
import ssl
import subprocess
from tempfile import TemporaryDirectory
from threading import Thread
from unittest.mock import patch
from urllib.request import HTTPSHandler, ProxyHandler, build_opener

from action_items import ensure_actions, record_action
from addin_server import make_handler
from bridge_client import BridgeClient, _NoRedirect
from event_sync import event_tables
from windows_production_worker import ProductionWorker
from windows_worker import FAKE_ACCOUNTS, tls_server


@contextmanager
def serving(server):
    with server:
        thread = Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            yield server.server_port
        finally:
            server.shutdown()
            thread.join()


def seed(root):
    adapters = dict.fromkeys(FAKE_ACCOUNTS)
    with patch("windows_worker.GRAPH_KEYCHAINS", adapters):
        worker = ProductionWorker(root, adapters, None, lambda: True)
    now = datetime.now(timezone.utc)
    examples = (
        ("Your verification code is valid for 30 minutes.", "We have received your application."),
        ("Please complete the online assessment by Friday.", "Your parcel was delivered."),
    )
    for account, subjects in zip(FAKE_ACCOUNTS, examples):
        folder = worker.directory(account)
        folder.mkdir()
        with closing(sqlite3.connect(folder / "decisions.sqlite3")) as db, db:
            event_tables(db, now)
            ensure_actions(db)
            for i, subject in enumerate(subjects):
                record_action(db, {
                    "id": str(i), "receivedDateTime": now.isoformat(), "subject": subject,
                    "from": {"emailAddress": {"address": "hiring@example.test"}},
                    "bodyPreview": subject, "webLink": "https://outlook.live.com/mail/",
                }, {"source": "Inbox", "target": "Inbox", "scores": []})
            db.execute("INSERT OR REPLACE INTO event_settings VALUES ('last_action_scan',?)", (now.isoformat(),))
        worker.connectivity(account, "inbox", True)
        worker.connectivity(account, "junkemail", True)
        worker.last_scan_ok[account] = True
    return worker


def verify(root, screenshot=None, channel=None):
    from playwright.sync_api import sync_playwright, expect
    worker = seed(root)
    cert, key = root / "cert.pem", root / "key.pem"
    openssl = os.environ.get("OUTLOOK_TRIAGE_OPENSSL") or shutil.which("openssl") or r"C:\Program Files\Git\usr\bin\openssl.exe"
    subprocess.run([openssl, "req", "-x509", "-newkey", "rsa:2048", "-nodes", "-days", "1",
                    "-subj", "/CN=localhost", "-addext", "subjectAltName=DNS:localhost,IP:127.0.0.1",
                    "-keyout", str(key), "-out", str(cert)], check=True, capture_output=True)
    bearer = secrets.token_urlsafe(48)
    accounts = {a: a for a in FAKE_ACCOUNTS}
    with serving(tls_server(worker, bearer, cert, key)) as worker_port:
        # The production constructor reads a POSIX owner-only file on the Mac.
        # Use ephemeral in-memory credentials to run these same HTTPS requests on Windows.
        bridge = BridgeClient.__new__(BridgeClient)
        bridge.base = f"https://127.0.0.1:{worker_port}"
        bridge.token, bridge.accounts = bearer, frozenset(accounts)
        bridge.opener = build_opener(ProxyHandler({}), HTTPSHandler(context=ssl.create_default_context(cafile=str(cert))), _NoRedirect())
        server = ThreadingHTTPServer(("127.0.0.1", 0), None)
        pane_token = secrets.token_urlsafe(32)
        server.RequestHandlerClass = make_handler(FAKE_ACCOUNTS[0], pane_token, server.server_port, accounts, bridge=bridge)
        tls = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        tls.load_cert_chain(cert, key)
        server.socket = tls.wrap_socket(server.socket, server_side=True)
        with serving(server) as pane_port, sync_playwright() as p:
            browser = p.chromium.launch(channel=channel)
            try:
                page = browser.new_page(viewport={"width":350,"height":700}, ignore_https_errors=True)
                errors = []
                # Browser-only verification exercises the pane's own API and no live Office or mail accounts.
                page.route("https://officeapis.public.onecdn.static.microsoft/**", lambda route: route.abort())
                page.on("pageerror", lambda error: errors.append(str(error)))
                page.goto(f"https://localhost:{pane_port}/p/{pane_token}/")
                expect(page.locator("article")).to_have_count(2)
                expect(page.get_by_text(FAKE_ACCOUNTS[0], exact=False)).to_be_visible()
                expect(page.get_by_text(FAKE_ACCOUNTS[1], exact=False)).to_be_visible()
                expect(page.get_by_text("We have received your application.")).to_have_count(0)
                assert page.evaluate("document.documentElement.scrollWidth <= innerWidth")
                if screenshot:
                    page.screenshot(path=screenshot)
                first = page.locator("article").first
                first.locator("summary").click()
                expect(first.get_by_role("link", name="打开原邮件")).to_have_attribute("href", "https://outlook.live.com/mail/")
                first.get_by_role("button", name="标记已完成").click()
                expect(page.locator("article")).to_have_count(1)
                # A successful undo restores the card even while the next full list refresh is unavailable.
                page.route("**/accounts", lambda route: route.fulfill(status=503))
                page.get_by_role("button", name="撤销").click()
                expect(page.locator("article")).to_have_count(2)
                page.unroute("**/accounts")
                for _ in range(2):
                    page.locator("article").first.locator("summary").click()
                    page.locator("article").first.get_by_role("button", name="无需处理").click()
                    expect(page.locator("article")).to_have_count(1 - _)
                page.reload()
                expect(page.locator("#empty")).to_be_visible()
                worker.connectivity(FAKE_ACCOUNTS[0], "inbox", False)
                page.reload()
                expect(page.locator("#connection")).to_contain_text("上次成功同步")
                expect(page.locator("#empty")).to_be_hidden()
                assert not errors, errors
            finally:
                browser.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--screenshot")
    parser.add_argument("--channel", help="Installed Chromium channel, for example msedge")
    args = parser.parse_args()
    with TemporaryDirectory() as scratch:
        verify(Path(scratch), args.screenshot, args.channel)
    print("PASS: merged actions, notification exclusion, links, narrow layout, durable commands, undo, empty and offline states")
