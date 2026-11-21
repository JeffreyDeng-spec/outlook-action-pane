"""Use Outlook IMAP IDLE only as a wake-up signal. Graph reads and classifies mail."""
import imaplib
import re
import select
import socket
import ssl
import time

from outlook_triage import AUTH, CLIENT_ID, DEFAULT_ACCOUNT, PreviewError, keychain, request_json

SERVICE = "codex.outlook-triage-imap"
SCOPE = "https://outlook.office.com/IMAP.AccessAsUser.All"
SERVER = "outlook.office365.com"
MAILBOX_EXISTS = re.compile(rb"^\* [0-9]+ EXISTS$", re.ASCII)
LIST_MAILBOX = re.compile(rb'^\(([^)]*)\) (?:NIL|"(?:\\.|[^"])*") (.+)$')


class IdleClient(imaplib.IMAP4_SSL):
    def open(self, host="", port=993, timeout=None):
        super().open(host, port, timeout)
        self.file.close()
        self.file = self.sock.makefile("rb", buffering=0)
        self.sock.setsockopt(socket.SOL_SOCKET, socket.SO_KEEPALIVE, 1)
        for name, value in (("TCP_KEEPALIVE", 25), ("TCP_KEEPIDLE", 25),
                            ("TCP_KEEPINTVL", 5), ("TCP_KEEPCNT", 3)):
            if hasattr(socket, name):
                try:
                    self.sock.setsockopt(socket.IPPROTO_TCP, getattr(socket, name), value)
                except OSError:
                    continue

    def idle(self, stop, wake, connected):
        # Python 3.9's imaplib has no public IDLE helper. Keep command parsing in imaplib.
        tag = self._new_tag()
        self.send(tag + b" IDLE\r\n")
        if self._get_response() is not None:
            raise PreviewError("Outlook IMAP 未开启 IDLE")
        connected()
        deadline = time.monotonic() + 24 * 60
        wall = time.time()
        while not stop.is_set() and time.monotonic() < deadline:
            ready = self.sock.pending() or select.select([self.sock], [], [], 5)[0]
            now = time.time()
            if now - wall > 24:
                raise imaplib.IMAP4.abort("system resumed")
            wall = now
            if ready:
                frame = self._get_response()
                if tag and self.tagged_commands.get(tag) is not None:
                    raise imaplib.IMAP4.abort("server ended IDLE")
                if frame and MAILBOX_EXISTS.fullmatch(frame):
                    wake()
        self.send(b"DONE\r\n")
        self._command_complete("IDLE", tag)


def junk_name(mailboxes):
    fallback = None
    for line in mailboxes:
        match = LIST_MAILBOX.fullmatch(line) if isinstance(line, bytes) else None
        if not match:
            continue
        flags, name = match.groups()
        if b"\\junk" in flags.lower().split():
            return name
        if name.strip(b'"').lower() in (b"junk", b"junk email"):
            fallback = name
    if not fallback:
        raise PreviewError("Outlook IMAP 未返回 Junk 文件夹")
    return fallback


def imap_token(account=DEFAULT_ACCOUNT):
    refresh = keychain(SERVICE, account)
    if not refresh:
        raise PreviewError("等待 Outlook IMAP 授权")
    reply = request_json(AUTH, "Microsoft IMAP 授权", form={
        "client_id": CLIENT_ID, "grant_type": "refresh_token",
        "refresh_token": refresh, "scope": SCOPE,
    })
    if not reply.get("access_token") or not any(
        scope.rsplit("/", 1)[-1] == "IMAP.AccessAsUser.All" for scope in reply.get("scope", "").split()
    ):
        raise PreviewError("Microsoft 未返回 IMAP.AccessAsUser.All 授权")
    return reply["access_token"]


def watch_folder(folder, stop, wake, status, account=DEFAULT_ACCOUNT, *, get_token=None):
    delay = 2
    while not stop.is_set():
        client = None
        try:
            token = get_token(account) if get_token else imap_token(account)
            client = IdleClient(SERVER, ssl_context=ssl.create_default_context(), timeout=35)
            credential = f"user={account}\x01auth=Bearer {token}\x01\x01".encode()
            responses = iter((credential, b""))
            client.authenticate("XOAUTH2", lambda _: next(responses, b""))
            if "IDLE" not in client.capabilities:
                typ, capabilities = client.capability()
                if typ != "OK" or b"IDLE" not in b" ".join(capabilities).split():
                    raise PreviewError("Outlook IMAP 服务器不支持 IDLE")
            name = "INBOX"
            if folder == "Junk":
                typ, folders = client.list()
                if typ != "OK":
                    raise PreviewError("无法通过 IMAP 定位 Junk")
                name = junk_name(folders)
            typ, _ = client.select(name, readonly=True)
            if typ != "OK":
                raise PreviewError("Outlook IMAP 文件夹无法打开")
            delay = 2

            def online():
                status(folder, True)
                wake()

            client.idle(stop, wake, online)
            while not stop.is_set():
                client.idle(stop, wake, lambda: None)
        except (PreviewError, OSError, imaplib.IMAP4.error):
            status(folder, False)
            if not stop.wait(delay):
                delay = min(30, delay * 2)
        finally:
            if client:
                try:
                    client.shutdown()
                except OSError:
                    pass
