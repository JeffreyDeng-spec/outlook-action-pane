from datetime import datetime,timedelta,timezone
from pathlib import Path
from tempfile import TemporaryDirectory
from time import sleep,monotonic
from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import patch
from urllib.parse import urlsplit,parse_qs
from uuid import uuid4
from contextlib import closing
import sqlite3

from mail_accounts import GRAPH_KEYCHAINS
from outlook_triage import AUTH,GRAPH,PreviewError,utcstamp
from windows_graph_metadata import GraphMetadata
from windows_pilot import build_pilot,ListenerSupervisor
from windows_worker import Worker

A,B=tuple(GRAPH_KEYCHAINS)
class Store:
    def __init__(self):
        self.items={("graph-read-basic",A):"DUMMY_A",("graph-read-basic",B):"DUMMY_B"}
    def get(self,kind,account):
        if kind not in {"graph-read-basic","graph-basic-bridge-bearer"}:
            raise AssertionError("Unapproved credentials")
        return self.items.get((kind,account),"DUMMY_TOKEN_ONLY_FOR_TEST")
    def set(self,kind,account,value):
        if kind!="graph-read-basic": raise AssertionError("Wrong secret type")
        self.items[kind,account]=value
class Fake:
    def __init__(self):
        self.scope="User.Read Mail.ReadBasic"
        self.items={}
        self.calls=[]
        self.cross=None
    def __call__(self,url,label,**kwargs):
        self.calls.append((url,set(kwargs)))
        if url==AUTH:
            account=kwargs["form"]["refresh_token"].replace("DUMMY_","").split("_")[-1]
            token="alpha" if account=="A" else "beta"
            assert "Mail.ReadBasic" in kwargs["form"]["scope"] and "IMAP" not in kwargs["form"]["scope"]
            return {"scope":self.scope,"access_token":token,
                    "refresh_token":"DUMMY_ROTATED_"+account,"expires_in":3600}
        account=kwargs["token"]
        parts=urlsplit(url).path.split("/")
        if parts[-1]=="me": return {"mail":A if account=="alpha" else B}
        if parts[-2]=="mailFolders": return {"id":account+"-"+parts[-1]}
        assert kwargs.get("payload") is None and kwargs.get("form") is None
        folder=parts[-3].removeprefix(account+"-")
        delta=GRAPH+f"/me/mailFolders/{account}-{folder}/messages/delta?%24deltatoken=dummy"
        if self.cross==account and folder=="inbox":
            delta=GRAPH+f"/me/mailFolders/{account}-junkemail/messages/delta?%24skiptoken=dummy"
            return {"value":[],"@odata.nextLink":delta}
        return {"value":self.items.get((account,folder),[]),"@odata.deltaLink":delta}

class GraphTests(TestCase):
    def test_scope_and_isolated_request_projection(self):
        fake=Fake(); store=Store()
        a=GraphMetadata(A,store,str(uuid4()),http=fake)
        with patch("socket.create_connection",side_effect=AssertionError("No real network")):
            self.assertEqual(a.graph_token(),"alpha")
            self.assertEqual(store.items["graph-read-basic",A],"DUMMY_ROTATED_A")
            self.assertEqual(store.items["graph-read-basic",B],"DUMMY_B")
            reader=a.folder_reader("inbox")
            for addr in (
                    GRAPH+"/me/mailFolders/inbox/messages/delta",
                    GRAPH+"/me/mailFolders/inbox/messages?$select=id,receivedDateTime,subject",
                    GRAPH+"/me/mailFolders/junkemail/messages/delta?$deltatoken=x",
                    GRAPH+"/me/messages?id=x"):
                self.assertRaises(PreviewError,reader,addr,"Dummy",token="alpha")
            for route in ("mailFolders('inbox')","mailFolders(%27inbox%27)","mailFolders('alpha-inbox')"):
                backend_calls=[]
                def test_http(url,label,**kwargs):
                    backend_calls.append(url)
                    return {"value":[]}
                with patch.object(a,"http",side_effect=test_http):
                    reader(GRAPH+"/me/"+route+"/messages/delta?$deltatoken=dummy","Dummy",token="alpha")
                self.assertEqual(len(backend_calls),1)
            for bad in ("mailFolders('junkemail')","mailFolders('beta-inbox')","mailFolders('alpha-junkemail')"):
                self.assertRaises(PreviewError,reader,GRAPH+"/me/"+bad+"/messages/delta?$deltatoken=dummy",
                                  "Dummy",token="alpha")
            self.assertRaises(PreviewError,reader,GRAPH+"/users/other/mailFolders('inbox')/messages/delta?$deltatoken=x",
                              "Dummy",token="alpha")
            self.assertRaises(PreviewError,reader,GRAPH+"/me/mailFolders('inbox')/messages/delta?$deltatoken=x&$expand=attachments",
                              "Dummy",token="alpha")
            fake.scope="User.Read Mail.ReadBasic Mail.Read"
            self.assertRaises(PreviewError,GraphMetadata(A,store,str(uuid4()),http=fake).graph_token)
            fake.scope="User.Read Mail.ReadBasic Mail.ReadWrite"
            self.assertRaises(PreviewError,GraphMetadata(B,store,str(uuid4()),http=fake).graph_token)

    def test_two_polling_accounts_two_folders_catch_up_and_no_sensitive_state(self):
        fake=Fake(); store=Store()
        with TemporaryDirectory() as folder, patch("socket.create_connection",side_effect=AssertionError("No network")):
            root=Path(folder)
            w=Worker(root,{i:GraphMetadata(i,store,str(uuid4()),http=fake) for i in (A,B)},
                     pilot=True,graph_only=True)
            def forbidden_imap(*args,**kwargs):
                raise AssertionError("Graph metadata mode unexpectedly used IMAP")
            def wait_for(predicate):
                until=monotonic()+4
                while monotonic()<until:
                    if predicate():return
                    sleep(.02)
                self.fail("Synthetic poll not observed")
            supervisor=ListenerSupervisor(w,watcher=forbidden_imap,poll_seconds=.2).start()
            try:
                wait_for(lambda:w.last_scan_ok[A] and w.last_scan_ok[B])
                self.assertEqual(len(supervisor.threads),3)
                stamp=utcstamp(datetime.now(timezone.utc)+timedelta(seconds=2))
                fake.items["alpha","inbox"]=[{"id":"same","receivedDateTime":stamp,
                       "subject":"SENSITIVE_SUBJECT","from":"SENSITIVE_FROM","bodyPreview":"SENSITIVE_BODY"}]
                fake.items["beta","junkemail"]=[{"id":"same","receivedDateTime":stamp}]
                wait_for(lambda:w.metadata_counts.get(A)==1 and w.metadata_counts.get(B)==1)
                self.assertIsNone(w.status(A)["push"])
                self.assertEqual(w.preview(A)["messages"],[])
            finally:
                self.assertTrue(supervisor.close(2))
            for account in (A,B):
                file=w.directory(account)/"decisions.sqlite3"
                with closing(sqlite3.connect(file)) as db:
                    self.assertEqual(db.execute("select count(*) from event_cursor").fetchone()[0],2)
                    self.assertEqual(db.execute("select count(*) from mailbox_moves").fetchone()[0],0)
                    self.assertEqual(db.execute("select count(*) from decisions").fetchone()[0],0)
                self.assertNotIn(b"SENSITIVE_",file.read_bytes())
            w2=Worker(root,{i:GraphMetadata(i,store,str(uuid4()),http=fake) for i in (A,B)},
                      pilot=True,graph_only=True)
            w2.scan(A);w2.scan(B)
            self.assertEqual(w2.metadata_counts,{A:1,B:1})
            fake.cross="alpha"
            # Force the mock cross-folder nextLink (test-only stub); auth scope unchanged.
            self.assertRaises(PreviewError,w2.scan,A)
            self.assertTrue(w2.last_scan_ok[B])

    def test_config_does_not_read_secrets_if_imap_or_wrong_state_marker(self):
        with TemporaryDirectory() as temp:
            base=Path(temp)
            state=base/"state"; state.mkdir()
            credentials=base/"credentials"; credentials.mkdir()
            marker=state/".outlook-triage-pilot-only"; marker.write_text("pilot-metadata-only")
            settings=SimpleNamespace(mode="graph-metadata",confirm_read_only_pilot=True,state_dir=state,
                       credential_dir=credentials,graph_client_id=str(uuid4()),
                       imap_client_id=None,poll_seconds=120)
            def forbidden(_):
                self.fail("Invalid configuration read credentials")
            self.assertRaises(ValueError,build_pilot,settings,provider=forbidden)
            marker.write_text("graph-basic-polling-only")
            settings.imap_client_id=str(uuid4())
            self.assertRaises(ValueError,build_pilot,settings,provider=forbidden)
