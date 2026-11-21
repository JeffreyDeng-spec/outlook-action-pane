"""Serve the installed Outlook action pane on Windows under its current owner."""
import argparse
from base64 import b64encode
import json
import os
from pathlib import Path
import re
import secrets
import ssl
import subprocess
import sys
from time import monotonic, sleep
from http.server import ThreadingHTTPServer
from urllib.parse import urlencode
from urllib.request import HTTPSHandler, ProxyHandler, build_opener

HERE = Path(__file__).resolve().parent
if __name__ == "__main__":
    sys.path.insert(0, str(HERE))

from addin_server import make_handler
from bridge_client import BridgeClient, BridgeError, _NoRedirect
from mail_accounts import GRAPH_KEYCHAINS
from outlook_triage import atomic_report
from windows_migration import APP, BRIDGE, OWNER_SID, PROFILE, ROOT, private_dir, regular_file
from worker_safety import exclusive_lock

PANE = ROOT / "windows-pane"
CERT = BRIDGE / "bridge-cert.pem"
KEY = BRIDGE / "bridge-key.pem"
BEARER = BRIDGE / "bridge-token"
MANIFEST = PANE / "manifest.xml"
ACCESS = PANE / "access"
PORT = 18767

def powershell(code, **env):
    path = Path(os.environ["SystemRoot"]) / "System32/WindowsPowerShell/v1.0/powershell.exe"
    result = subprocess.run([str(path), "-NoLogo", "-NoProfile", "-NonInteractive", "-Command",
                             "$ErrorActionPreference='Stop'\n" + code],
                            env={**os.environ, **env, "PSModulePath": str(path.parent / "Modules")},
                            capture_output=True, timeout=40)
    if result.returncode:
        raise ValueError("Windows security or startup check failed")


def check_owner_and_acl(paths):
    powershell(r"""
$me = [Security.Principal.WindowsIdentity]::GetCurrent()
if ($me.User.Value -ne $env:PANE_OWNER -or
    ([Security.Principal.WindowsPrincipal]::new($me)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) -or
    [Environment]::GetFolderPath('LocalApplicationData') -ne $env:PANE_PROFILE) {throw 'Wrong owner'}
$allowed = @($me.User.Value, 'S-1-5-18', 'S-1-5-32-544')
foreach ($p in (ConvertFrom-Json $env:PANE_PATHS)) {
    $item = Get-Item -Force -LiteralPath $p
    if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {throw 'Reparse point'}
    $acl = Get-Acl -LiteralPath $p
    if ($acl.GetOwner([Security.Principal.SecurityIdentifier]).Value -notin $allowed) {throw 'Owner'}
    $rules = $acl.GetAccessRules($true,$true,[Security.Principal.SecurityIdentifier])
    if ($rules.Count -eq 0) {throw 'Missing access rules'}
    foreach ($rule in $rules) {
        if ($rule.AccessControlType -eq 'Allow' -and $rule.IdentityReference.Value -notin $allowed) {throw 'Shared ACL'}
    }
}
""", PANE_OWNER=OWNER_SID, PANE_PROFILE=str(PROFILE), PANE_PATHS=json.dumps(list(map(str, paths))))


def validate_installation(worker_port):
    if os.name != "nt" or HERE != (APP if worker_port == 18765 else ROOT / "pane-candidate"):
        raise ValueError("Run the installed current-user pane package")
    for path in (PROFILE, ROOT, BRIDGE, HERE):
        private_dir(path)
    for path in (KEY, CERT, BEARER):
        regular_file(path)
    check_owner_and_acl([ROOT, HERE, BRIDGE, KEY, CERT, BEARER])
    PANE.mkdir(exist_ok=True)
    private_dir(PANE)
    existing = [path for path in (ACCESS, MANIFEST) if path.exists()]
    check_owner_and_acl([PANE, *map(regular_file, existing)])


def prepare():
    with exclusive_lock(PANE / "setup.lock"):
        if not ACCESS.exists():
            with ACCESS.open("x", encoding="ascii") as output:
                output.write(secrets.token_urlsafe(32))
        token = regular_file(ACCESS).read_text(encoding="ascii")
        if not re.fullmatch(r"[A-Za-z0-9_-]{43}", token):
            raise ValueError("Invalid pane access material")
        template = regular_file(HERE / "addin/manifest.xml.in").read_text(encoding="utf-8")
        xml = template.replace("__BASE__", f"https://localhost:{PORT}/p/{token}")
        if not MANIFEST.exists() or MANIFEST.read_text(encoding="utf-8") != xml:
            atomic_report(MANIFEST, xml)
        return token


def make_bridge(worker_port):
    return BridgeClient(f"https://localhost:{worker_port}", BEARER, GRAPH_KEYCHAINS, CERT,
                        read_protected_token=lambda path: regular_file(path).read_text(encoding="ascii").strip())


def serve(token, worker_port):
    accounts = {a:a for a in GRAPH_KEYCHAINS}
    handler = make_handler(next(iter(accounts)), token, PORT, accounts, bridge=make_bridge(worker_port))
    with ThreadingHTTPServer(("127.0.0.1", PORT), handler) as server:
        tls = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        tls.minimum_version = ssl.TLSVersion.TLSv1_2
        tls.load_cert_chain(CERT, KEY)
        server.socket = tls.wrap_socket(server.socket, server_side=True)
        server.serve_forever()


def trust_certificate(*, check_only=False):
    der = ssl.PEM_cert_to_DER_cert(CERT.read_text(encoding="ascii"))
    powershell(r"""
$c = [Security.Cryptography.X509Certificates.X509Certificate2]::new([Convert]::FromBase64String($env:PANE_CERT))
$constraints = $c.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.19' }
$eku = $c.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.37' }
$san = $c.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.17' }
$expected = [byte[]](@(0x30,0x11,0x82,0x09) + [Text.Encoding]::ASCII.GetBytes('localhost') + @(0x87,0x04,127,0,0,1))
if (!$constraints -or $constraints.CertificateAuthority -or
    $c.Subject -ne 'CN=localhost' -or $c.Issuer -ne $c.Subject -or $c.NotAfter -le [DateTime]::Now -or
    $c.NotBefore -gt [DateTime]::Now -or '1.3.6.1.5.5.7.3.1' -notin $eku.EnhancedKeyUsages.Value -or
    !$san -or [Convert]::ToBase64String($san.RawData) -ne [Convert]::ToBase64String($expected)) {throw 'Invalid leaf'}
if ($env:PANE_CERT_CHECK -eq '1') {return}
if (!(Test-Path -LiteralPath ("Cert:\CurrentUser\Root\"+$c.Thumbprint)) -and
    !(Test-Path -LiteralPath ("Cert:\LocalMachine\Root\"+$c.Thumbprint))) {
    $store = [Security.Cryptography.X509Certificates.X509Store]::new('Root','CurrentUser')
    $store.Open('ReadWrite')
    try {$store.Add($c)} finally {$store.Close()}
}
""", PANE_CERT=b64encode(der).decode("ascii"), PANE_CERT_CHECK="1" if check_only else "0")


def probe(token):
    opener = build_opener(ProxyHandler({}), HTTPSHandler(context=ssl.create_default_context(cafile=str(CERT))), _NoRedirect())
    base = f"https://localhost:{PORT}/p/{token}/"
    with opener.open(base+"accounts", timeout=4) as response:
        if set(json.load(response)["accounts"]) != set(GRAPH_KEYCHAINS):
            raise ValueError("Unexpected account view")
    for account in GRAPH_KEYCHAINS:
        with opener.open(base+"actions?"+urlencode({"account":account}), timeout=8) as response:
            result = json.load(response)
            if result.get("account") != account or not isinstance(result.get("items"), list):
                raise ValueError("Unexpected actions response")


def install(token):
    headless = Path(sys.executable).with_name("pythonw.exe")
    regular_file(headless)
    powershell(r"""
$argsText='-I -S -B "'+$env:PANE_SCRIPT+'" serve'
$linkPath=Join-Path ([Environment]::GetFolderPath('Startup')) 'Outlook Action Pane.lnk'
$ws=New-Object -ComObject WScript.Shell
$link=$ws.CreateShortcut($linkPath)
if ((Test-Path -LiteralPath $linkPath) -and ($link.TargetPath -ne $env:PANE_PYTHON -or $link.Arguments -ne $argsText)) {throw 'Unexpected existing link'}
$listener=@(Get-NetTCPConnection -LocalPort 18767 -State Listen -ErrorAction SilentlyContinue)
if ($listener.Count -gt 0) {
    $p=Get-CimInstance Win32_Process -Filter ('ProcessId='+$listener[0].OwningProcess)
    if ($listener.Count -ne 1 -or $listener[0].LocalAddress -ne '127.0.0.1' -or $p.ExecutablePath -ne $env:PANE_PYTHON -or !$p.CommandLine.EndsWith($argsText)) {throw 'Port in use'}
} else {
    (New-Object -ComObject Shell.Application).ShellExecute($env:PANE_PYTHON,$argsText,$env:PANE_APP,'open',0)
}
$link.TargetPath=$env:PANE_PYTHON
$link.Arguments=$argsText
$link.WorkingDirectory=$env:PANE_APP
$link.Description='Outlook action pane on localhost'
$link.Save()
""", PANE_PYTHON=str(headless), PANE_APP=str(HERE), PANE_SCRIPT=str(HERE / "windows_pane.py"))
    deadline = monotonic() + 20
    while True:
        try:
            probe(token)
            return
        except OSError:
            if monotonic() >= deadline:
                raise
            sleep(1)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=("prepare", "trust", "serve", "probe", "install"))
    parser.add_argument("--worker-port", type=int, choices=(18765,18768), default=18765)
    args = parser.parse_args()
    if args.command == "install" and args.worker_port != 18765:
        parser.error("Startup must use the production worker")
    try:
        validate_installation(args.worker_port)
        token = prepare()
        if args.command in {"trust", "install"}:
            trust_certificate()
        if args.command == "serve":
            serve(token, args.worker_port)
        elif args.command == "probe":
            probe(token)
        elif args.command == "install":
            install(token)
        print("Pane manifest:", MANIFEST)
    except (OSError, ValueError, RuntimeError, BridgeError, subprocess.SubprocessError) as error:
        # Exception text from urllib can contain the private pane URL.
        print("Pane operation failed:", type(error).__name__, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
