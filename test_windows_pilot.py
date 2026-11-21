from contextlib import closing
from pathlib import Path
import os,sqlite3
from tempfile import TemporaryDirectory
from threading import Event, Lock
from types import SimpleNamespace
from time import sleep, monotonic
from unittest import TestCase
from unittest.mock import patch
from uuid import uuid4

from mail_accounts import GRAPH_KEYCHAINS
from windows_credentials import DPAPIStore
from outlook_triage import AUTH, GRAPH, PreviewError
from windows_pilot_adapters import MicrosoftReadOnly
from windows_pilot import ListenerSupervisor, build_pilot, main
from windows_worker import Worker

A,B=tuple(GRAPH_KEYCHAINS)
class Adapter:
    def __init__(self,account):
        self.account=account
        self.reads=0
    def graph_token(self):
        return "DUMMY"
    def imap_token(self,account):
        if account!=self.account:
            raise AssertionError("Wrong account")
        return "DUMMY"
    def read(self,url,label,**kwargs):
        self.reads+=1
        return {"value":[],"@odata.deltaLink":url}
    def classify(self,*args,**kwargs):
        raise AssertionError("Pilot called Jev")
    def write(self,*args,**kwargs):
        raise AssertionError("Pilot moved mail")

class PilotTests(TestCase):
    def test_no_arguments_opens_nothing_and_unsafe_start_refused(self):
        with patch("socket.create_connection",side_effect=AssertionError("No real network")):
            with patch("windows_pilot.DPAPIStore",side_effect=AssertionError("No credential load")):
                self.assertEqual(main([]),0)
            with self.assertRaises(ValueError):
                build_pilot(SimpleNamespace(confirm_read_only_pilot=False),provider=lambda _:None)
            with TemporaryDirectory() as d:
                with self.assertRaises(ValueError):
                    Worker(Path(d),{a:Adapter(a) for a in (A,B)},pilot=True,apply=True)

    def test_supervisor_four_distinct_listeners_and_reconnect(self):
        with TemporaryDirectory() as d, patch("socket.create_connection",side_effect=AssertionError("No real network")):
            result={}
            mutex=Lock()
            w=Worker(Path(d),{a:Adapter(a) for a in (A,B)},pilot=True)
            def watcher(folder,stop,wake,status,account,*,get_token):
                self.assertEqual(get_token(account),"DUMMY")
                signal=Event()
                status(folder,True)
                wake()
                with mutex:
                    result[account,folder]=signal
                while not stop.is_set():
                    if signal.wait(.05):
                        status(folder,False)
                        status(folder,True)
                        wake()
                        signal.clear()
            def wait_for(test):
                until=monotonic()+4
                while monotonic()<until:
                    if test(): return
                    sleep(.02)
                self.fail("Synthetic event not observed")
            s=ListenerSupervisor(w,watcher=watcher).start()
            try:
                wait_for(lambda: len(result)==4 and all(x.reads>=2 for x in w.adapters.values()))
                self.assertEqual(set(result),{(a,f) for a in (A,B) for f in ("Inbox","Junk")})
                self.assertEqual(w.status(A)["push"],"online")
                # Wait for all already-pending coalesced events to settle before compare.
                sleep(.3)
                start=w.adapters[A].reads,w.adapters[B].reads
                result[A,"Junk"].set()
                wait_for(lambda: w.adapters[A].reads>=start[0]+2)
                self.assertEqual(w.adapters[B].reads,start[1])
                for a in (A,B):
                    path=w.directory(a)/"decisions.sqlite3"
                    with closing(sqlite3.connect(path)) as db:
                        self.assertEqual(db.execute("select count(*) from event_cursor").fetchone()[0],2)
                        self.assertEqual(db.execute("select count(*) from mailbox_moves").fetchone()[0],0)
            finally:
                self.assertTrue(s.close(2))

    def test_graph_scope_identity_and_dpapi_atomic_rotation_separate_resources(self):
        with TemporaryDirectory() as d, patch("socket.create_connection",side_effect=AssertionError("No real network")):
            root=Path(d)
            (root/".outlook-triage-dpapi").touch()
            store=DPAPIStore(root)
            store.set("graph-read",A,"DUMMY_ALPHAGRAPH")
            store.set("imap-idle",A,"DUMMY_ALPHAIMAP")
            store.set("graph-read",B,"DUMMY_BETAGRAPH")
            def http(url,label,**kwargs):
                if url==AUTH:
                    form=kwargs["form"]
                    if "graph.microsoft.com" in form["scope"]:
                        self.assertNotIn("ReadWrite",form["scope"])
                        return {"scope":"User.Read Mail.Read","access_token":"DUMMY_ACCESS",
                                "refresh_token":"DUMMY_ROTATED","expires_in":3600}
                    return {"scope":"IMAP.AccessAsUser.All","access_token":"DUMMY_IMAP","expires_in":3600}
                self.assertTrue(url.startswith(GRAPH+"/me?"))
                return {"mail":A}
            p=MicrosoftReadOnly(A,store,str(uuid4()),str(uuid4()),http=http)
            self.assertEqual(p.graph_token(),"DUMMY_ACCESS")
            self.assertEqual(p.imap_token(A),"DUMMY_IMAP")
            self.assertEqual(store.get("graph-read",A),"DUMMY_ROTATED")
            self.assertEqual(store.get("graph-read",B),"DUMMY_BETAGRAPH")
            self.assertEqual(store.get("imap-idle",A),"DUMMY_ALPHAIMAP")
            self.assertEqual(list(root.glob(".rotate-*")),[])
            for file in root.glob("*.dpapi"):
                self.assertNotIn(b"DUMMY_",file.read_bytes())
            self.assertRaises(PreviewError,p.imap_token,B)
            self.assertRaises(RuntimeError,p.write,"unused")
            q=MicrosoftReadOnly(A,store,str(uuid4()),str(uuid4()),http=lambda *a,**kw:
                                 {"scope":"User.Read Mail.Read Mail.ReadWrite","access_token":"DUMMY_EXPAND"})
            self.assertRaises(PreviewError,q.graph_token)
            other=MicrosoftReadOnly(B,store,str(uuid4()),str(uuid4()),http=http)
            self.assertRaises(PreviewError,other.graph_token)
            self.assertEqual(store.get("graph-read",B),"DUMMY_BETAGRAPH")
