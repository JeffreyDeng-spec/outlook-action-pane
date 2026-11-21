"""Windows pane setup must never put the bridge credential into the add-in manifest."""
from contextlib import ExitStack
from pathlib import Path
import secrets
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from xml.etree import ElementTree

import windows_pane as pane


class PaneSetupTests(unittest.TestCase):
    def setUp(self):
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        scratch = Path(self.stack.enter_context(TemporaryDirectory()))
        pane_dir = scratch / "windows-pane"
        pane_dir.mkdir()
        self.private_bearer = secrets.token_urlsafe(32)
        bearer = scratch / "bridge-token"
        bearer.write_text(self.private_bearer, encoding="ascii")
        for name, path in {"PANE":pane_dir, "ACCESS":pane_dir/"access", "MANIFEST":pane_dir/"manifest.xml",
                           "BEARER":bearer}.items():
            self.stack.enter_context(patch.object(pane, name, path))

    def test_prepare_preserves_pane_identity_and_keeps_bridge_credential_private(self):
        access = pane.prepare()
        self.assertRegex(access, r"^[A-Za-z0-9_-]{43}$")
        self.assertNotEqual(access, self.private_bearer)
        xml = pane.MANIFEST.read_text(encoding="utf-8")
        tree = ElementTree.fromstring(xml)
        self.assertNotIn("__BASE__", xml)
        self.assertNotIn(self.private_bearer, xml)
        self.assertTrue(any(node.get("DefaultValue") == f"https://localhost:18767/p/{access}/" for node in tree.iter()))
        self.assertEqual(pane.prepare(), access)

    def test_existing_malformed_credential_requires_recovery(self):
        pane.ACCESS.write_text("malformed", encoding="ascii")
        with self.assertRaises(ValueError):
            pane.prepare()
        self.assertFalse(pane.MANIFEST.exists())

    def test_worktree_cannot_read_private_credentials_as_an_installed_pane(self):
        with patch.object(pane, "HERE", Path("uninstalled").absolute()), patch.object(pane,"check_owner_and_acl") as acl:
            with self.assertRaises(ValueError):
                pane.validate_installation(18768)
            acl.assert_not_called()
