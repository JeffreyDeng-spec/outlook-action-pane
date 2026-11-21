"""Private, authenticated connection from the Mac pane host to the Windows worker."""
import json
import os
import re
import ssl
import stat
from urllib.parse import urlencode, urlsplit
from urllib.request import HTTPSHandler, HTTPRedirectHandler, ProxyHandler, Request, build_opener

MAX_BODY = 2 * 1024 * 1024


class BridgeError(Exception):
    pass


def action_command(route):
    match = re.fullmatch(r"actions/([0-9a-f]{32})/(done|ignore|undo)", route)
    return match.groups() if match else None


class _NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, *args):
        return None


class BridgeClient:
    @staticmethod
    def _read_mac_token(token_file):
        with os.fdopen(os.open(token_file.expanduser(), os.O_RDONLY | os.O_NOFOLLOW),
                       "r", encoding="ascii") as handle:
            info = os.fstat(handle.fileno())
            if (not stat.S_ISREG(info.st_mode) or info.st_uid != os.getuid()
                    or info.st_mode & 0o077):
                raise ValueError()
            return handle.read(1025).strip()

    def __init__(self, url, token_file, accounts, ca_file=None, *, read_protected_token=None):
        try:
            parsed = urlsplit(url)
            _ = parsed.port
            if (parsed.scheme != "https" or not parsed.hostname or not parsed.netloc
                    or parsed.username is not None or parsed.password is not None
                    or parsed.query or parsed.fragment or parsed.path not in ("", "/")):
                raise ValueError()
            self.token = (read_protected_token or self._read_mac_token)(token_file)
            if not re.fullmatch(r"[A-Za-z0-9._~-]{32,512}", self.token):
                raise ValueError()
            tls = ssl.create_default_context(cafile=str(ca_file.expanduser()) if ca_file else None)
            tls.minimum_version = ssl.TLSVersion.TLSv1_2
        except (OSError, ValueError, UnicodeError):
            raise BridgeError("Windows bridge needs an HTTPS origin, trusted CA and private token file") from None
        self.base = url.rstrip("/")
        self.accounts = frozenset(accounts)
        # Never send credentials via ambient proxy settings and never follow redirects.
        self.opener = build_opener(ProxyHandler({}), HTTPSHandler(context=tls), _NoRedirect())

    def request(self, route, account=None):
        if route == "accounts" and account is None:
            suffix = ""
        elif (route in {"status", "preview", "scan", "actions"} or action_command(route)) and account in self.accounts:
            suffix = "?" + urlencode({"account": account})
        else:
            raise BridgeError("Invalid worker route or account")
        command = action_command(route)
        post = route == "scan" or bool(command)
        try:
            req = Request(self.base + "/" + route + suffix,
                          data=b"" if post else None, method="POST" if post else "GET",
                          headers={"Authorization": "Bearer " + self.token,
                                   "Accept": "application/json"})
            with self.opener.open(req, timeout=8) as response:
                if response.status != (202 if route == "scan" else 200) or response.headers.get_content_type() != "application/json":
                    raise ValueError()
                content = response.read(MAX_BODY + 1)
            if len(content) > MAX_BODY:
                raise ValueError()
            result = json.loads(content)
            if not isinstance(result, dict):
                raise ValueError()
            if route == "accounts":
                items = result.get("accounts")
                if not isinstance(items, list) or set(items) != self.accounts or len(items) != len(self.accounts) or result.get("default") not in self.accounts:
                    raise ValueError()
                return {"accounts": items, "default": result["default"]}
            if result.get("account") != account:
                raise ValueError()
            if route == "actions":
                fields = ("id", "kind", "title", "source", "received", "due", "detail", "mail_url", "draft_url")
                rows = result.get("items")
                if (not isinstance(rows, list) or type(result.get("sync_ok")) is not bool or
                        not isinstance(result.get("last_sync"), str)):
                    raise ValueError()
                items = []
                for row in rows:
                    if (not isinstance(row, dict) or any(not isinstance(row.get(k), str) for k in fields)
                            or not re.fullmatch(r"[0-9a-f]{32}", row["id"])):
                        raise ValueError()
                    items.append({k: row[k] for k in fields})
                return {"account": account, "items": items, "sync_ok": result["sync_ok"],
                        "last_sync": result["last_sync"], "backfilling": result.get("backfilling") is True}
            if command:
                if (result.get("id"), result.get("command")) != command:
                    raise ValueError()
                receipt = {"account": account, "id": command[0], "command": command[1]}
                if type(result.get("visible")) is bool:
                    receipt["visible"] = result["visible"]
                return receipt
            if route == "preview":
                if not isinstance(result.get("window"), str) or not isinstance(result.get("messages"), list):
                    raise ValueError()
                columns = ("time", "source", "target", "subject", "domain", "reason", "scores")
                items = result["messages"]
                if len(items) > 400 or any(not isinstance(row, dict) or
                    any(not isinstance(row.get(k), str) for k in columns) for row in items):
                    raise ValueError()
                return {"account": account, "window": result["window"],
                        "messages": [{k: row[k] for k in columns} for row in items]}
            if route == "status":
                fields = {"revision": str, "running": bool, "message": str, "event_mode": bool}
                if any(type(result.get(k)) is not v for k, v in fields.items()):
                    raise ValueError()
                if result.get("push") not in ("online", "connecting", "waiting_auth", None):
                    raise ValueError()
                status = {"account": account, **{k: result[k] for k in fields}, "push": result.get("push")}
                if type(result.get("apply")) is bool:
                    status["apply"] = result["apply"]
                return status
            return {"account": account}
        except (OSError, ValueError, UnicodeError, TypeError):
            raise BridgeError("Windows worker connection or response verification failed") from None
