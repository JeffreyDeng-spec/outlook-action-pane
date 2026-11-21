"""Synthetic business cases, including replay after an ambiguous remote draft."""
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path
import sqlite3
from tempfile import TemporaryDirectory
import unittest

from action_items import action_for, ensure_actions, needs_full_text, open_actions, outlook_link, prepare_replies, record_action, update_action
from outlook_triage import PreviewError
from test_windows_production_worker import A, B, fake_backend, new_database
from test_windows_worker import FakeGraph
from mail_accounts import account_state
from windows_production_worker import ProductionWorker


ROW = {"id": "one", "source": "Inbox", "target": "Inbox", "scores": [
    ("legitimate", .97, .95), ("required", .95, .9), ("other", .95, .95)]}


def mail(subject, body):
    return {"id": "one", "source": "Inbox", "receivedDateTime": datetime.now(timezone.utc).isoformat(),
            "from": {"emailAddress": {"address": "team@example.test"}}, "subject": subject,
            "webLink": "https://outlook.live.com/mail/read/one", "bodyPreview": body[:255],
            "uniqueBody": {"contentType": "text", "content": body}}


class ActionItemsTests(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(":memory:")
        self.addCleanup(self.db.close)
        ensure_actions(self.db)

    def test_positive_and_negative_controls(self):
        for subject, body, kind in (
            ("Application received", "Your application has been received. No reply is required.", None),
            ("Action required", "Your package was delivered.", None),
            ("Interview scheduled", "Your interview is confirmed. No action required.", None),
            ("你的申请已收到", "请在周五前完成在线测评。", "material"),
            ("下一步", "你的申请已收到，请在周五前完成在线测评。", "material"),
            ("Interview invitation", "Please select your interview time: Monday at 10:00 UTC.", "interview"),
            ("Application", "Please upload your transcript by Friday.", "material"),
            ("Checking in", "Please reply to confirm receipt.", "reply"),
            ("继续申请", "请回复确认已收到邮件。", "reply"),
            ("已收到申请", "无需回复确认收到。", None),
        ):
            with self.subTest(body=body):
                item = action_for(mail(subject, body), ROW)
                self.assertEqual(item["kind"] if item else None, kind)

    def test_request_after_preview_is_found_without_fetching_mail_links(self):
        msg = mail("Upcoming assessment", "Thank you for your interest. " * 15 +
                   "Please finish the assessment by Friday. https://example.test/secret-token")
        item = action_for(msg, ROW)
        self.assertEqual((item["kind"],item["due"]), ("material","by Friday"))
        self.assertNotIn("secret-token", str(item))

    def test_otp_local_and_expiration(self):
        msg = mail("Your verification code is 843612", "Use it to sign in. Valid for 10 minutes.")
        row = {**ROW, "target": "review", "scores": []}
        record_action(self.db, msg, row)
        items = open_actions(self.db)
        self.assertEqual(len(items), 1)
        self.assertNotIn("843612", str(items))
        self.assertEqual(items[0]["kind"], "verify")
        self.db.execute("UPDATE message_actions SET expires='2000-01-01T00:00:00Z'")
        self.assertEqual(open_actions(self.db), [])

    def test_sensitive_and_junk_evidence_gate(self):
        otp = mail("Your verification code", "Use this code to sign in.")
        self.assertIsNone(action_for(otp, {**ROW, "scores": [("junk", .98, .99)]}))
        self.assertIsNone(action_for(otp, {**ROW, "source": "Junk"}))
        quoted = mail("Re: Your verification code", "Thanks!\n> Your verification code was issued.")
        self.assertIsNone(action_for(quoted, ROW))
        expired = mail("Verification code", "Your code has expired.")
        self.assertIsNone(action_for(expired, ROW))
        private = mail("Your bank statement", "Bank account balance. " * 22)
        self.assertFalse(needs_full_text(private, ROW))
        self.assertFalse(needs_full_text(otp, ROW))
        long_mail = mail("Application", "Thank you for your interest. " * 12 + "Please complete the assessment.")
        self.assertTrue(needs_full_text(long_mail, ROW))
        self.assertFalse(needs_full_text(long_mail, {**ROW, "target":"Junk"}))

    def test_normal_mail_recovered_from_junk_including_lost_receipt(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp).resolve() / "state"
            fakes = {A:FakeGraph(), B:FakeGraph()}
            for account in (A,B):
                directory = account_state(root,account)
                directory.mkdir(parents=True)
                new_database(directory / "decisions.sqlite3")
                fakes[account].add("same", "junkemail" if account == A else "inbox", True, "测评邀请")
                fakes[account].messages["same"]["bodyPreview"] = "请在周五前完成在线测评。"
            with closing(sqlite3.connect(account_state(root,B) / "decisions.sqlite3")) as db, db:
                db.execute("INSERT INTO mailbox_moves VALUES ('same','Junk','Inbox','pending','2026-11-19T00:00:00Z')")
            adapters = {a:fake_backend(a,fakes[a],lambda: True) for a in (A,B)}
            w = ProductionWorker(root, adapters, "", lambda:True, classifier=fakes[A].classify)
            for account in (A,B):
                w.scan(account)
                w.scan(account)
                self.assertEqual([i["kind"] for i in w.actions(account)["items"]], ["material"])
            self.assertEqual(fakes[A].posts, [("same", "inbox")])
            self.assertEqual(fakes[B].posts, [])

    def test_user_state_persists_on_replay(self):
        msg = mail("Test", "Please complete the assessment by Friday.")
        record_action(self.db, msg, ROW)
        item, = open_actions(self.db)
        update_action(self.db, item["id"], "ignore")
        record_action(self.db, msg, ROW)
        self.assertEqual(open_actions(self.db), [])
        update_action(self.db, item["id"], "undo")
        self.assertEqual(open_actions(self.db), [item])
        self.assertEqual(outlook_link("https://outlook.live.com@evil.example/"), "")

    def test_remote_draft_succeeds_but_response_is_lost(self):
        msg = mail("Request", "Please reply to confirm receipt.")
        record_action(self.db, msg, ROW)
        class Adapter:
            calls = 0
            def create_reply(self, *args):
                self.calls += 1
                raise PreviewError("simulated timeout")
        adapter = Adapter()
        with self.assertRaises(PreviewError):
            prepare_replies(self.db, adapter, "")
        record_action(self.db, msg, ROW)
        prepare_replies(self.db, adapter, "")
        self.assertEqual(adapter.calls, 1)
        self.assertIn("草稿", open_actions(self.db)[0]["detail"])

    def test_real_worker_scan_backfill_isolated_and_reconnect(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp).resolve() / "state"
            fakes = {A:FakeGraph(), B:FakeGraph()}
            for account in (A,B):
                directory = account_state(root,account)
                directory.mkdir(parents=True)
                new_database(directory / "decisions.sqlite3")
                fakes[account].add("same", "inbox", True, "Inbox test")
            fakes[A].messages["same"].update(mail("面试邀请","请选择面试时段，周一 14:00 或周二 16:00（北京时间）。"))
            fakes[B].messages["same"].update(mail("Application receipt","Your application was received."))
            for account in (A,B):
                fakes[account].messages["same"]["id"] = "same"
            with closing(sqlite3.connect(account_state(root,A) / "decisions.sqlite3")) as db, db:
                db.execute("INSERT INTO event_checked VALUES ('same')")
            adapters = {a:fake_backend(a,fakes[a],lambda: True) for a in (A,B)}
            w = ProductionWorker(root, adapters, "", lambda:True, classifier=fakes[B].classify)
            for account in (A,B):
                w.scan(account)
                w.scan(account)
            self.assertEqual([i["kind"] for i in w.actions(A)["items"]], ["interview"])
            self.assertEqual(w.actions(B)["items"], [])
            self.assertEqual(fakes[A].posts+fakes[B].posts,[])
