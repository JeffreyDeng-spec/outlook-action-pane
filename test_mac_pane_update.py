import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from urllib.parse import parse_qs, urlsplit

from tools.install_mac_pane import ASSETS, manifest_origin, replace_from, verify


class MacPaneUpdateTests(unittest.TestCase):
    def test_manifest_keeps_one_loopback_origin_and_matching_access_path(self):
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "manifest.xml"
            token = "b" * 43
            for host, expected in (("localhost", 8765), ("example.test", None)):
                path.write_text(f'<OfficeApp><SourceLocation DefaultValue="https://{host}:8765/p/{token}/"/></OfficeApp>')
                if expected is None:
                    with self.assertRaises(ValueError):
                        manifest_origin(path, token)
                else:
                    self.assertEqual(manifest_origin(path, token), expected)
                    with self.assertRaises(ValueError):
                        manifest_origin(path, "a" * 43)

    def test_verify_both_accounts_and_unhealthy_response(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            for name in ASSETS:
                (root / name).write_text(name)
            accounts = ["alpha@example.test", "beta@example.test"]
            checked = []
            def fetch(route):
                if route == "accounts":
                    return json.dumps({"accounts": accounts}).encode()
                if route.startswith("actions?"):
                    account = parse_qs(urlsplit(route).query)["account"][0]
                    checked.append(account)
                    return json.dumps({"account": account, "items": [], "sync_ok": healthy}).encode()
                return (route or "index.html").encode()
            healthy = True
            verify(fetch, frozenset(accounts), root)
            self.assertCountEqual(checked, accounts)
            healthy = False
            with self.assertRaises(ValueError):
                verify(fetch, frozenset(accounts), root)

    def test_atomic_swap_can_restore_original_bytes(self):
        with TemporaryDirectory() as tmp:
            paths = [Path(tmp) / n for n in ("release", "installed", "backup")]
            for path, value in zip(paths, (b"new", b"old", b"old")):
                path.write_bytes(value)
            replace_from(paths[0], paths[1])
            self.assertEqual(paths[1].read_bytes(), b"new")
            replace_from(paths[2], paths[1])
            self.assertEqual(paths[1].read_bytes(), b"old")
