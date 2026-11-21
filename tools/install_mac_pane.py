"""Update a running Mac localhost pane from a verified release checkout."""
import argparse
from contextlib import closing
from http.client import HTTPSConnection, HTTPException
import json
import os
from pathlib import Path
import plistlib
import re
import shutil
import ssl
import stat
import subprocess
import sys
import tempfile
from time import monotonic, sleep
from urllib.parse import urlencode, urlsplit
from xml.etree import ElementTree

SOURCE = Path(__file__).resolve().parents[1]
SERVER_FILES = ("addin_server.py", "bridge_client.py")
ASSETS = ("index.html", "pane.js", "pane.css")
LABEL = "ai.codex.outlook-triage-bridge"


def regular(path):
    if path.is_symlink() or not path.is_file():
        raise ValueError("An expected installation file is missing or redirected")
    return path


def replace_from(src, target):
    with tempfile.NamedTemporaryFile(dir=target.parent, delete=False) as handle:
        temporary = Path(handle.name)
    try:
        shutil.copyfile(src, temporary)
        temporary.chmod(stat.S_IMODE(target.stat().st_mode))
        temporary.replace(target)
    finally:
        temporary.unlink(missing_ok=True)


def manifest_origin(manifest, token):
    if not re.fullmatch(r"[A-Za-z0-9_-]{43}", token):
        raise ValueError("Invalid installed pane access configuration")
    ports = set()
    for node in ElementTree.parse(manifest).iter():
        if node.tag.rsplit("}", 1)[-1] == "SourceLocation":
            url = urlsplit(node.get("DefaultValue", ""))
            if (url.scheme != "https" or url.hostname != "localhost" or url.username or url.password
                    or url.path.rstrip("/") != "/p/" + token or url.query or url.fragment):
                raise ValueError("Expected the existing authenticated localhost pane")
            ports.add(url.port or 443)
    if len(ports) != 1:
        raise ValueError("Could not identify the installed localhost port")
    return ports.pop()


def request(port, tls, token, route):
    # Direct HTTPS prevents redirects or proxy forwarding of the local access path.
    with closing(HTTPSConnection("localhost", port, timeout=8, context=tls)) as conn:
        conn.request("GET", "/p/" + token + "/" + route)
        res = conn.getresponse()
        if res.status != 200:
            raise ValueError("Pane endpoint did not return success")
        data = res.read(2 * 1024 * 1024 + 1)
        if len(data) > 2 * 1024 * 1024:
            raise ValueError("Oversized pane response")
        return data


def installed_accounts(fetch):
    names = json.loads(fetch("accounts")).get("accounts")
    if not isinstance(names, list) or len(names) != 2 or not all(isinstance(n, str) for n in names) or len(set(names)) != 2:
        raise ValueError("Expected the two configured mailbox views")
    return frozenset(names)


def verify(fetch, accounts, assets):
    if installed_accounts(fetch) != accounts:
        raise ValueError("Mailbox configuration changed")
    for name in ASSETS:
        if fetch("" if name == "index.html" else name) != (assets / name).read_bytes():
            raise ValueError("The running pane did not serve the installed assets")
    for name in accounts:
        view = json.loads(fetch("actions?" + urlencode({"account": name})))
        if view.get("account") != name or not isinstance(view.get("items"), list) or view.get("sync_ok") is not True:
            raise ValueError("An action view is unavailable or not current")


def retry(check):
    limit = monotonic() + 22
    while True:
        try:
            return check()
        except (ValueError, OSError, HTTPException):
            if monotonic() >= limit:
                raise
            sleep(0.5)


def ctl(command, target):
    args = ["kickstart", "-k", target] if command == "restart" else ["print", target]
    return subprocess.run(["/bin/launchctl", *args], stdout=subprocess.DEVNULL,
                          stderr=subprocess.DEVNULL, timeout=20).returncode == 0


def install(project, state, verify_only):
    if sys.platform != "darwin":
        raise ValueError("Run this utility on the Mac that hosts the existing pane")
    import fcntl

    bridge, addin, api = state / "windows-bridge", project / "addin", state / "addin"
    for folder in (project, state, bridge, addin, api):
        if folder.is_symlink() or not folder.is_dir() or folder.stat().st_uid != os.getuid():
            raise ValueError("Expected user-owned installation directories")
    regular(bridge / "mac_handoff.py")
    domain = f"gui/{os.getuid()}/"
    service = domain + LABEL
    plist = regular(Path.home() / "Library/LaunchAgents" / (LABEL + ".plist"))
    args = plistlib.loads(plist.read_bytes()).get("ProgramArguments", [])
    if args[-2:] != [str(bridge / "mac_handoff.py"), "serve"] or not ctl("loaded", service):
        raise ValueError("The existing bridge LaunchAgent must be active before updating")
    if ctl("loaded", domain + "ai.codex.outlook-triage"):
        raise ValueError("The original Mac writer is active")
    token = regular(api / "access").read_text(encoding="ascii")
    manifest = regular(api / "manifest.xml")
    manifest_bytes = manifest.read_bytes()
    port = manifest_origin(manifest, token)
    tls = ssl.create_default_context(cafile=str(regular(api / "localhost.pem")))
    tls.minimum_version = ssl.TLSVersion.TLSv1_2
    def fetch(route):
        return request(port, tls, token, route)
    accounts = installed_accounts(fetch)
    def ready():
        verify(fetch, accounts, SOURCE / "addin")
        if manifest_origin(manifest, token) != port or manifest.read_bytes() != manifest_bytes:
            raise ValueError("The existing manifest changed during restart")
    if verify_only:
        ready()
        return
    project_template = regular(addin / "manifest.xml.in").read_bytes()
    rendered = project_template.replace(b"__BASE__", f"https://localhost:{port}/p/{token}".encode())
    if rendered.replace(b"\r\n", b"\n") != manifest_bytes:
        raise ValueError("The project template differs from the installed manifest")
    target_files = [(SOURCE / name, bridge / name) for name in SERVER_FILES]
    target_files.extend((SOURCE / "addin" / name, addin / name) for name in ASSETS)
    for source, target in target_files:
        regular(source)
        regular(target)
    with (bridge / "pane-update.lock").open("a+") as lock:
        fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        backup = Path(tempfile.mkdtemp(prefix="pane-backup-", dir=bridge))
        for _, target in target_files:
            shutil.copy2(target, backup / target.name)
        print("Previous pane files saved in", backup)
        try:
            for source, target in target_files:
                replace_from(source, target)
            if not ctl("restart", service):
                raise ValueError("The pane LaunchAgent could not restart")
            retry(ready)
        except BaseException:
            for _, target in target_files:
                replace_from(backup / target.name, target)
            if not ctl("restart", service):
                raise RuntimeError("Rollback restored the files but the pane could not restart") from None
            def restored():
                if installed_accounts(fetch) != accounts:
                    raise ValueError("Rollback account check failed")
            retry(restored)
            raise


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project", required=True, type=Path, help="Existing private project directory")
    parser.add_argument("--state", type=Path, default=Path.home() / "Library/Application Support/OutlookTriage")
    parser.add_argument("--verify-only", action="store_true")
    options = parser.parse_args()
    try:
        install(options.project.expanduser().absolute(), options.state.expanduser().absolute(), options.verify_only)
    except (OSError, ValueError, RuntimeError, HTTPException, subprocess.SubprocessError, ElementTree.ParseError) as error:
        # Never emit raw HTTP or JSON errors containing mailbox data or pane access material.
        parser.exit(1, f"Pane update did not pass ({type(error).__name__}). Check the local agent and bridge.\n")
    print("Installed pane assets and both authenticated account views verified.")


if __name__ == "__main__":
    main()
