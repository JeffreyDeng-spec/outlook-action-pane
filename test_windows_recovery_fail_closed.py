import sqlite3
from datetime import datetime, timezone
from urllib.parse import parse_qs, urlsplit
from unittest import TestCase
from event_sync import collect_folder,event_tables,process_queue
from outlook_triage import GRAPH,IMMUTABLE_ID,PreviewError,apply_moves

FLOOR="2026-11-01T00:00:00Z"
CURSOR=GRAPH+"/me/mailFolders/inbox/messages/delta?$deltatoken=old"
NEXT=GRAPH+"/me/mailFolders/inbox/messages/delta?$skiptoken=second"
NEW=GRAPH+"/me/mailFolders/inbox/messages/delta?$deltatoken=new"
NOW=datetime(2026,11,20,12,tzinfo=timezone.utc)
OLDER="2026-11-10T00:00:00Z"
STARTED="2026-11-20T00:00:00Z"

class RecoveryTests(TestCase):
    def db(self):
        db=sqlite3.connect(":memory:")
        self.addCleanup(db.close)
        event_tables(db,datetime(2026,11,1,tzinfo=timezone.utc))
        return db

    def expired_db(self):
        db=self.db()
        with db:
            db.execute("INSERT INTO event_cursor VALUES ('inbox',?,?)",(CURSOR,STARTED))
            db.execute("INSERT INTO event_checked VALUES ('already')")
            db.execute("INSERT INTO event_queue VALUES ('queued',?)",(OLDER,))
        return db

    def state(self,db):
        return {t:db.execute("SELECT * FROM "+t).fetchall()
                for t in ("event_cursor","event_checked","event_queue")}

    def test_410_page_failure_in_reseed_or_backfill_keeps_original_state(self):
        for failure_path in ("delta","listing"):
            with self.subTest(failure_path=failure_path):
                db=self.expired_db(); before=self.state(db)
                def read(url,_,**opts):
                    self.assertEqual(opts["headers"],IMMUTABLE_ID)
                    if url==CURSOR: raise PreviewError("expired",http_status=410)
                    if url==NEXT: raise PreviewError("second page failed",http_status=503)
                    if urlsplit(url).path.endswith("/delta"):
                        return {"value":[{"id":"new","receivedDateTime":STARTED}],
                                **({"@odata.nextLink":NEXT} if failure_path=="delta" else {"@odata.deltaLink":NEW})}
                    return {"value":[{"id":"new","receivedDateTime":STARTED}],"@odata.nextLink":NEXT}
                with self.assertRaises(PreviewError): collect_folder(db,"inbox","DUMMY",FLOOR,NOW,read)
                self.assertEqual(self.state(db),before)

    def test_production_410_recovers_an_unseen_message_received_before_cursor(self):
        db=self.expired_db()
        # Server honors received-time filters. Fixture history says this old
        # unreviewed message just moved into Inbox, but reset replies cannot
        # express whether it was moved recently or had always been in Inbox.
        def read(url,_,**opts):
            if url==CURSOR: raise PreviewError("expired",http_status=410)
            if "/delta?" in url:
                return {"value":[],"@odata.deltaLink":NEW}
            since=parse_qs(urlsplit(url).query)["$filter"][0]
            return {"value":[{"id":"old-moved-after-expiry","receivedDateTime":OLDER}]
                    if FLOOR in since else []}
        collect_folder(db,"inbox","DUMMY",FLOOR,NOW,read,recover_from_floor=True)
        ids={r[0] for r in db.execute("SELECT id FROM event_queue")}
        self.assertIn("old-moved-after-expiry",ids)

    def test_pending_receipts_visible_and_never_posted_even_without_queue(self):
        for label,proof,target_ok in (
            ("missing",PreviewError("not found",http_status=404),False),
            ("source",{"id":"old","parentFolderId":"inbox"},False),
            ("third",{"id":"old","parentFolderId":"archive"},False),
            ("noid",{"parentFolderId":"junkemail"},False),
            ("wrongid",{"id":"wrong","parentFolderId":"junkemail"},False),
            ("destination",{"id":"old","parentFolderId":"junkemail"},True)):
            with self.subTest(label=label):
                db=self.db()
                with db: db.execute("INSERT INTO mailbox_moves VALUES ('old','Inbox','Junk','pending',?)",(STARTED,))
                if label=="source":
                    with db: db.execute("INSERT INTO event_queue VALUES ('old',?)",(OLDER,))
                def fail(*a,**kw): self.fail("A pending receipt was classified or POSTed")
                def read(url,_,**opts):
                    if "/mailFolders/" in url: return {"id":urlsplit(url).path.split("/")[-1]}
                    self.assertEqual(opts["headers"],IMMUTABLE_ID)
                    self.assertIn("select=id,parentFolderId",url)
                    if isinstance(proof,Exception): raise proof
                    return proof
                mover=lambda d,r,t: apply_moves(d,r,t,read=read,write=fail)
                rows,counts,_=process_queue(db,"DUMMY","DUMMY",FLOOR,read=read,classifier=fail,mover=mover)
                self.assertEqual(len(rows),1)
                self.assertEqual(rows[0]["id"],"old")
                self.assertEqual(counts["unconfirmed"],int(not target_ok))
                self.assertEqual(db.execute("SELECT state FROM mailbox_moves").fetchone()[0],
                                 "moved" if target_ok else "pending")
                self.assertEqual(db.execute("SELECT count(*) FROM event_checked").fetchone()[0],int(target_ok))
                if label=="source":
                    self.assertEqual(db.execute("SELECT id FROM event_queue").fetchall(),[("old",)])

    def test_preflight_and_response_require_exact_identity(self):
        for initial,returned,expected_posts in (("other","old",0),("old","other",1),("old",None,1),("old","old",1)):
            with self.subTest(initial=initial,returned=returned):
                db=self.db(); posts=[]
                data={"id":"old","source":"Inbox","target":"Junk","change_key":"mock","reason":""}
                def read(url,_,**opts):
                    if "/messages/" not in url: return {"id":urlsplit(url).path.split("/")[-1]}
                    self.assertEqual(opts["headers"],IMMUTABLE_ID)
                    return {"id":initial,"parentFolderId":"inbox","changeKey":"mock"}
                def write(url,_,**opts):
                    self.assertEqual(opts["headers"],IMMUTABLE_ID); posts.append(url)
                    return {"id":returned,"parentFolderId":"junkemail"}
                apply_moves(db,[data],"DUMMY",read=read,write=write)
                self.assertEqual(len(posts),expected_posts)
                state=db.execute("SELECT state FROM mailbox_moves").fetchone()
                self.assertEqual(state,None if not expected_posts else ("moved",) if returned=="old" else ("pending",))
