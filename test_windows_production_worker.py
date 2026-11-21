from contextlib import closing
from datetime import datetime, timezone
import json
from pathlib import Path
import sqlite3
from tempfile import TemporaryDirectory
from unittest import TestCase

from event_sync import event_tables
from mail_accounts import GRAPH_KEYCHAINS, account_state
from outlook_triage import GRAPH, IMMUTABLE_ID, PreviewError, decision_key
from test_windows_worker import FakeGraph
from windows_migration import import_backups
from windows_production_worker import ProductionAdapter, ProductionWorker


A, B = GRAPH_KEYCHAINS
STAMP = datetime(2026, 11, 1, tzinfo=timezone.utc)




def fake_backend(account, fake, gate):
    def http(url, label, **kwargs):
        request = fake.write if "payload" in kwargs else fake.read
        return request(url, label, **kwargs)

    adapter = ProductionAdapter(account, None, None, gate, http=http)
    adapter.folder_ids = {"inbox": "inbox", "junkemail": "junkemail"}
    adapter.graph_token = lambda: "synthetic"
    return adapter


def new_database(path):
    with closing(sqlite3.connect(path)) as db:
        event_tables(db, STAMP)
        with db:
            for name in ("inbox", "junkemail"):
                delta = f"{GRAPH}/me/mailFolders/{name}/messages/delta?$deltatoken=old"
                db.execute("INSERT INTO event_cursor VALUES (?,?,?)",
                           (name, delta, "2026-11-20T00:00:00Z"))


class ProductionWorkerTests(TestCase):
    def setUp(self):
        folder = TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.root = Path(folder.name).resolve(strict=True) / "state"
        self.first, self.second = FakeGraph(), FakeGraph()
        self.enabled = True
        self.gate = lambda: self.enabled
        for account in (A, B):
            directory = account_state(self.root, account)
            directory.mkdir(parents=True)
            new_database(directory / "decisions.sqlite3")

    def worker(self):
        return ProductionWorker(
            self.root, {A: fake_backend(A, self.first, self.gate), B: fake_backend(B, self.second, self.gate)},
            "dummy", self.gate, classifier=self.first.classify,
        )

    def test_bidirectional_moves_and_reconnect_preserve_account_identity(self):
        self.first.add("same", "inbox", False, "A promo")
        self.second.add("same", "junkemail", True, "B personal")
        worker = self.worker()
        for account in (A, B):
            self.assertTrue(worker.scan(account))
        reopened = self.worker()
        for account in (A, B):
            self.assertTrue(reopened.scan(account))
        self.assertEqual(self.first.posts, [("same", "junkemail")])
        self.assertEqual(self.second.posts, [("same", "inbox")])
        self.assertEqual(self.first.classify_count, 2)
        self.assertEqual(reopened.preview(A)["messages"][0]["subject"], "A promo")
        self.assertEqual(reopened.preview(B)["messages"][0]["subject"], "B personal")

    def test_sensitive_mail_bypasses_migrated_cache_and_provider(self):
        self.first.add("private", "inbox", False, "Your verification code")
        email = self.first.messages["private"]
        with closing(sqlite3.connect(account_state(self.root, A) / "decisions.sqlite3")) as db:
            with db:
                db.execute("INSERT INTO decisions VALUES (?,?,?,?)", (
                    email["id"], email["receivedDateTime"], decision_key(email, "jev-latest"),
                    json.dumps({"target": "Junk", "reason": "old prediction"}),
                ))
        worker = self.worker()
        worker.scan(A)
        self.assertEqual(self.first.posts, [])
        self.assertEqual(self.first.classify_count, 0)
        self.assertIn("待复核", worker.preview(A)["messages"][0]["reason"])

    def test_move_gate_blocks_graph_posts_even_during_scan(self):
        self.first.add("test", "inbox", False, "sale")
        worker = self.worker()
        self.enabled = False
        with self.assertRaises(PreviewError):
            worker.scan(A)
        with self.assertRaises(PreviewError):
            worker.adapters[A].write(
                GRAPH + "/me/mailFolders/inbox/messages/test/move", "test",
                headers=IMMUTABLE_ID, payload={"destinationId": "junkemail"},
            )
        self.enabled = True
        with self.assertRaises(PreviewError):
            worker.adapters[A].write(
                GRAPH + "/me/mailFolders/inbox/messages/test/move", "test",
                headers=IMMUTABLE_ID, payload={"destinationId": "deleteditems"},
            )
        self.assertEqual(self.first.posts, [])

    def test_invalid_second_backup_does_not_publish_a_partial_import(self):
        destination = self.root.parent / "imported"
        first = account_state(self.root, A) / "decisions.sqlite3"
        second = account_state(self.root, B) / "decisions.sqlite3"
        with closing(sqlite3.connect(second)) as db:
            with db:
                db.execute("DELETE FROM event_cursor WHERE folder='junkemail'")
        with self.assertRaises(ValueError):
            import_backups(destination, {A: first, B: second})
        self.assertFalse(destination.exists())
        self.assertEqual(list(destination.parent.glob(".outlook-import-*")), [])

    def test_import_keeps_original_cursors_and_preserves_both_archives(self):
        destination = self.root.parent / "imported"
        files = {account: account_state(self.root, account) / "decisions.sqlite3" for account in (A, B)}
        expected = {}
        for index, (account, path) in enumerate(files.items()):
            with closing(sqlite3.connect(path)) as db:
                with db:
                    db.execute("INSERT INTO event_checked VALUES (?)", (str(index),))
            expected[account] = path.read_bytes()
        reports = import_backups(destination, files)
        self.assertEqual(set(reports), {A, B})
        for index, account in enumerate((A, B)):
            with closing(sqlite3.connect(account_state(destination, account) / "decisions.sqlite3")) as db:
                self.assertEqual(db.execute("SELECT id FROM event_checked").fetchall(), [(str(index),)])
                self.assertEqual(db.execute("SELECT count(*) FROM event_cursor").fetchone(), (2,))
            self.assertEqual(files[account].read_bytes(), expected[account])
