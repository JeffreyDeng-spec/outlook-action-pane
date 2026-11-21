from tempfile import TemporaryDirectory
from pathlib import Path
from unittest import TestCase
from unittest.mock import patch
from uuid import uuid4
from windows_production_grants import (production_roots,enroll_production,verify_imap_account,scope,
                                       ProductionTrialAdapter,PreviewError,GRAPH,AUTH)
from windows_graph_owner_setup import roots
from windows_production_trial import listener_trial
from windows_worker import Worker
from mail_accounts import GRAPH_KEYCHAINS

A,B=tuple(GRAPH_KEYCHAINS)
class Store:
    def __init__(self):
        self.items={}
    def set(self,resource,account,token):
        self.items[resource,account]=token
    def get(self,resource,account):
        return self.items.get((resource,account))

class ProductionPreparationTests(TestCase):
    def test_separate_approved_grants_exact_identity_and_scope(self):
        store=Store()
        identities=[]
        client=str(uuid4())
        def callback(cid,account,c,state,*,scope,timeout):
            self.assertEqual(cid,client)
            self.assertEqual(timeout,300)
            self.assertNotEqual(c,state)
            is_graph="Mail.ReadWrite" in scope
            self.assertEqual("IMAP" in scope,not is_graph)
            self.assertIn("offline_access",scope)
            return account,"http://localhost:23456/outlook-triage-callback"
        def http(url,label,**kw):
            if url==AUTH:
                form=kw["form"]
                self.assertNotIn("client_secret",form)
                self.assertTrue(form["code_verifier"])
                a=form["code"]
                role="graph" if "Mail.ReadWrite" in form["scope"] else "imap"
                return {"scope":"User.Read Mail.ReadWrite" if role=="graph" else "IMAP.AccessAsUser.All",
                        "access_token":a,"refresh_token":"DUMMY-"+role+"-"+a}
            self.assertTrue(url.startswith(GRAPH+"/me?"))
            return {"mail":kw["token"]}
        def imap(account,token):
            identities.append(account)
            self.assertEqual(account,token)
        with patch("socket.create_connection",side_effect=AssertionError):
            for a in (A,B):
                for r in ("graph-write","imap-examine"):
                    enroll_production(client,a,r,store,http=http,callback=callback,imap_identity=imap)
            self.assertEqual(identities,[A,B])
            self.assertEqual(len(set(store.items.values())),4)
            self.assertRaises(PreviewError,enroll_production,client,B,"imap-examine",store,callback=callback,
               http=lambda *a,**kw: {"scope":"IMAP.AccessAsUser.All Mail.Send","access_token":"DUMMY","refresh_token":"NONE"})
            self.assertRaises(PreviewError,enroll_production,client,B,"graph-write",store,callback=callback,
               http=lambda url,*a,**kw: ({"scope":"User.Read Mail.ReadWrite","access_token":"DUMMY","refresh_token":"NONE"}
                                      if url==AUTH else {"mail":A}))
            self.assertTrue(all(v.startswith("DUMMY-") for v in store.items.values()))

    def test_imap_probe_never_selects_write_mode(self):
        calls=[]
        class Reader:
            def __init__(self,host,**kw): calls.append(("TLS",host,kw["ssl_context"].check_hostname))
            def authenticate(self,mech,response):
                value=response(None)
                self.assert_auth=mech=="XOAUTH2" and value.startswith(f"user={A}".encode())
                calls.append(("auth-exact",self.assert_auth))
            def list(self): return "OK",[b'(\\Junk) "/" "Junk Email"']
            def select(self,name,*,readonly):
                calls.append(("EXAMINE",name,readonly))
                return "OK",[]
            def shutdown(self): calls.append(("shutdown",))
        verify_imap_account(A,"DUMMY",connect=Reader)
        self.assertEqual([c for c in calls if c[0]=="EXAMINE"],
                          [("EXAMINE","INBOX",True),("EXAMINE",b'"Junk Email"',True)])
        self.assertEqual(calls[-1],("shutdown",))
        self.assertTrue(calls[0][-1])
        self.assertTrue(calls[1][-1])

    def test_production_terminal_trial_watcher_and_no_mutations(self):
        class FakeAdapter:
            def __init__(self,account): self.account=account
            def graph_token(self): return "DUMMY"
            def imap_token(self,account):
                if account!=self.account: raise AssertionError("Wrong account")
                return "DUMMY"
            def folder_reader(self,folder):
                return lambda url,label,**kw:{"value":[],"@odata.deltaLink":url}
            def write(self,*a,**kw): raise AssertionError("Moves prohibited")
            def classify(self,*a,**kw): raise AssertionError("Jev prohibited")
        with TemporaryDirectory() as d, patch("socket.create_connection",side_effect=AssertionError):
            self.assertNotEqual(roots(d),production_roots(d))
            w=Worker(Path(d),{a:FakeAdapter(a) for a in (A,B)},pilot=True,folder_bound=True)
            ran=[]
            def watcher(f,stop,wake,status,a,*,get_token):
                self.assertEqual(get_token(a),"DUMMY")
                ran.append((a,f))
                status(f,True);wake();stop.wait(2)
            result,closed=listener_trial(w,3,watcher=watcher)
            self.assertTrue(closed)
            self.assertEqual(set(ran),{(a,f) for a in (A,B) for f in ("Inbox","Junk")})
            self.assertTrue(all(all(x.values()) for x in result.values()))
            self.assertRaises(ValueError,Worker,Path(d),w.adapters,pilot=True,folder_bound=True,apply=True)
