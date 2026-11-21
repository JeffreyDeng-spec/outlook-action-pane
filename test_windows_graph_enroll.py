from threading import Thread
from pathlib import Path
from tempfile import TemporaryDirectory
from urllib.parse import parse_qs, urlsplit, urlencode
from urllib.error import HTTPError
from urllib.request import build_opener, ProxyHandler, Request
from unittest import TestCase
from unittest.mock import patch
from uuid import uuid4
from types import SimpleNamespace
import base64,hashlib

from mail_accounts import GRAPH_KEYCHAINS
from outlook_triage import AUTH,GRAPH,PreviewError
from windows_graph_enroll import PATH,enroll,make_pkce,parse_callback,receive_code
from windows_pilot import build_pilot
A,B=tuple(GRAPH_KEYCHAINS)

class LocalStore:
    def __init__(self): self.saved=[]
    def set(self,*args): self.saved.append(args)

class EnrollTests(TestCase):
    def test_state_validation_and_live_loopback_callback_without_microsoft(self):
        v,c,state=make_pkce()
        self.assertEqual(c,base64.urlsafe_b64encode(hashlib.sha256(v.encode()).digest()).rstrip(b"=").decode())
        self.assertIsNone(parse_callback(PATH+"?code=a&state=wrong","localhost:2","localhost:2",state))
        self.assertIsNone(parse_callback(PATH+"?code=a&state="+state,"example.test","localhost:2",state))
        self.assertIsNone(parse_callback(PATH+"?code=a&state="+state+"&state="+state,"localhost:2","localhost:2",state))
        def launch(url):
            p=urlsplit(url); q=parse_qs(p.query)
            self.assertEqual(p.hostname,"login.microsoftonline.com")
            self.assertEqual(q["code_challenge_method"],["S256"])
            endpoint=q["redirect_uri"][0]
            def call():
                cli=build_opener(ProxyHandler({}))
                with cli.open(Request(endpoint+"?"+urlencode({"code":"DUMMY_CODE","state":state})), timeout=2) as reply:
                    self.assertEqual(reply.status,200)
                    self.assertNotIn(b"DUMMY_CODE",reply.read())
            Thread(target=call,daemon=True).start()
            return True
        code,uri=receive_code(str(uuid4()),A,c,state,launch=launch,timeout=3)
        self.assertEqual(code,"DUMMY_CODE")
        self.assertEqual(urlsplit(uri).hostname,"localhost")

    def test_credential_written_only_for_exact_identity_and_narrow_scope(self):
        def callback(app,account,challenge,state):
            self.assertNotEqual(challenge,state)
            return "DUMMY_CODE","http://localhost:18051"+PATH
        store=LocalStore()
        def http(url,label,**kw):
            if url==AUTH:
                self.assertTrue(kw["form"]["code_verifier"])
                self.assertNotIn("client_secret",kw["form"])
                return {"access_token":"DUMMY_ACCESS","refresh_token":"DUMMY_REFRESH","scope":"User.Read Mail.ReadBasic"}
            self.assertTrue(url.startswith(GRAPH+"/me?"))
            return {"mail":A}
        with patch("socket.create_connection",side_effect=AssertionError("No real network")):
            enroll(str(uuid4()),A,store,callback=callback,http=http)
            self.assertRaises(PreviewError,enroll,str(uuid4()),B,store,callback=callback,http=http)
            self.assertRaises(PreviewError,enroll,str(uuid4()),A,store,callback=callback,http=lambda *a,**b:{
                "access_token":"DUMMY_ACCESS","refresh_token":"DUMMY_REFRESH","scope":"User.Read Mail.ReadBasic Mail.ReadWrite"})
        self.assertEqual(store.saved,[("graph-read-basic",A,"DUMMY_REFRESH")])

    def test_allow_check_once_without_tls_or_api_bearer_only_in_graph_mode(self):
        with TemporaryDirectory() as temp:
            p=Path(temp)
            state=p/"state"; state.mkdir()
            auth=p/"auth"; auth.mkdir()
            (state/".outlook-triage-pilot-only").write_text("graph-basic-polling-only")
            args=SimpleNamespace(mode="graph-metadata",state_dir=state,credential_dir=auth,
                                 graph_client_id=str(uuid4()),imap_client_id=None,poll_seconds=120,
                                 confirm_read_only_pilot=True,check_once=True,tls_cert=None,tls_key=None)
            w,bearer=build_pilot(args,provider=lambda _:object(),
                                 graph_adapter=lambda account,*_:SimpleNamespace(account=account))
            self.assertTrue(w.graph_only)
            self.assertFalse(w.apply)
            self.assertIsNone(bearer)
            self.assertEqual(tuple(w.adapters),(A,B))
