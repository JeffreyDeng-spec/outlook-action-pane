from collections import Counter
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path
from tempfile import TemporaryDirectory
from threading import Thread
from unittest.mock import patch
from urllib.parse import urlsplit, unquote
from urllib.request import urlopen,Request
from urllib.error import HTTPError, URLError
import os, shutil
import sqlite3, ssl, subprocess, unittest, json

from outlook_triage import GRAPH, PreviewError, utcstamp
from windows_worker import Worker, FAKE_ACCOUNTS, locked_tls_server
A,B=FAKE_ACCOUNTS
STAMP=utcstamp(datetime(2026,11,19,tzinfo=timezone.utc))
TOKEN="syn_test_"+("a"*40)
SSL_EXE = os.getenv("OUTLOOK_TRIAGE_OPENSSL") or shutil.which("openssl")

class FakeGraph:
    def __init__(self):
        self.messages={}
        self.posts=[]
        self.mode=None
        self.classify_count=0
        self.delta_pages=Counter()

    def add(self, ident, folder, good, subject):
        self.messages[ident]={"id":ident,"receivedDateTime":STAMP,"parentFolderId":folder,
                              "changeKey":"v1","subject":subject,"good":good,"bodyPreview":"dummy",
                              "from":{"emailAddress":{"address":"sender@example.test"}}}
    def read(self,url,label,**kwargs):
        path=tuple(str(unquote(urlsplit(url).path)).strip("/").split("/")[2:])
        if path[:1]==("mailFolders",) and len(path)==2:
            return {"id":path[1]}
        if path[:1]==("mailFolders",) and path[-1]=="delta":
            folder=path[1]
            self.delta_pages[folder]+=1
            return {"value":[{"id":m["id"],"receivedDateTime":STAMP} for m in self.messages.values()
                            if m["parentFolderId"]==folder],
                    "@odata.deltaLink":GRAPH+"/me/mailFolders/"+folder+"/messages/delta?$deltatoken=abc"}
        if path[-2:-1]==("messages",) and path[-1] in self.messages:
            return dict(self.messages[path[-1]])
        raise AssertionError("Unexpected mock read")
    def classify(self,email,key,model):
        self.classify_count+=1
        target="Inbox" if email["good"] else "Junk"
        return {"target":target,"reason":"MOCK","scores":[("legitimate" if email["good"] else "marketing",.99,.99)],
                "signals":{"bulk_content":0 if email["good"] else .99,"recipient_specific":.99 if email["good"] else 0}}
    def write(self,url,label,**kwargs):
        ident=unquote(urlsplit(url).path).split("/")[-2]
        target=kwargs["payload"]["destinationId"]
        self.posts.append((ident,target))
        if self.mode=="wrong":
            return {"parentFolderId":"wrong"}
        self.messages[ident]["parentFolderId"]=target
        if self.mode=="crash":
            raise RuntimeError("synthetic crash after server accepted move")
        return {"id": ident, "parentFolderId":target}

class WorkerTests(unittest.TestCase):
    def setUp(self):
        self.tmp=TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root=Path(self.tmp.name)
        self.a,self.b=FakeGraph(),FakeGraph()
        self.a.add("same","inbox",False,"ALPHA mock promotion")
        self.b.add("same","junkemail",True,"BETA mock normal")

    def make(self, apply=False):
        return Worker(self.root,{A:self.a,B:self.b},apply=apply)

    def dbrows(self,w,account,query):
        with closing(sqlite3.connect(w.directory(account)/"decisions.sqlite3")) as conn:
            return conn.execute(query).fetchall()

    def test_default_no_network_no_move_and_does_not_mark_checked(self):
        w=self.make()
        with patch("socket.create_connection",side_effect=AssertionError("No external socket")):
            w.scan(A); w.scan(B)
        self.assertFalse(self.a.posts or self.b.posts)
        self.assertEqual(self.a.classify_count+self.b.classify_count,0)
        for a in (A,B):
            self.assertEqual(self.dbrows(w,a,"select count(*) from event_checked"),[(0,)])
            self.assertEqual(self.dbrows(w,a,"select count(*) from event_queue"),[(1,)])

    def test_both_moves_and_reconnect_dedup_after_restart(self):
        w=self.make(True)
        with patch("socket.create_connection",side_effect=AssertionError("No external socket")):
            w.scan(A); w.scan(B)
            w2=self.make(True)
            w2.connectivity(A,"inbox",True);w2.connectivity(A,"junkemail",True)
            self.assertEqual(w2.status(A)["push"],"online")
            self.assertEqual(w2.status(B)["push"],"connecting")
            w2.scan(A);w2.scan(B)
        self.assertEqual(self.a.posts,[("same","junkemail")])
        self.assertEqual(self.b.posts,[("same","inbox")])
        self.assertEqual(self.a.classify_count+self.b.classify_count,2)
        self.assertEqual(w.preview(A)["messages"][0]["subject"],"ALPHA mock promotion")
        self.assertEqual(w.preview(B)["messages"][0]["subject"],"BETA mock normal")
        self.assertNotEqual(w.directory(A),w.directory(B))

    def test_crash_after_move_and_wrong_destination_no_retry_post(self):
        w=self.make(True)
        with patch("socket.create_connection",side_effect=AssertionError("No external socket")):
            self.a.mode="crash"
            with self.assertRaises(RuntimeError):
                w.scan(A)
            self.assertEqual(self.dbrows(w,A,"select state from mailbox_moves"),[("pending",)])
            self.a.mode=None
            w=self.make(True)
            w.scan(A)
            self.assertEqual(self.dbrows(w,A,"select state from mailbox_moves"),[("moved",)])
            self.assertEqual(len(self.a.posts),1)
            self.b.mode="wrong"
            w.scan(B);w.scan(B)
            self.assertEqual(self.dbrows(w,B,"select state from mailbox_moves"),[("pending",)])
            self.assertEqual(self.dbrows(w,B,"select count(*) from event_queue"),[(1,)])
            self.assertEqual(len(self.b.posts),1)

    @unittest.skipUnless(SSL_EXE and Path(SSL_EXE).exists(),"Existing Windows OpenSSL utility required")
    def test_local_tls_auth_manual_scan_no_moves_and_account_identity(self):
        w=self.make()
        cert,key=self.root/"cert.pem",self.root/"key.pem"
        subprocess.run([SSL_EXE,"req","-x509","-newkey","rsa:2048","-nodes","-days","1","-subj","/CN=localhost",
            "-addext","basicConstraints=critical,CA:TRUE","-addext","subjectAltName=DNS:localhost,IP:127.0.0.1",
            "-keyout",str(key),"-out",str(cert)],check=True,capture_output=True)
        with locked_tls_server(w,TOKEN,cert,key) as server:
            worker=Thread(target=server.serve_forever,daemon=True)
            worker.start()
            try:
                base="https://localhost:"+str(server.server_address[1])
                tls=ssl.create_default_context(cafile=str(cert))
                def request(path,method="GET",key=TOKEN):
                    data=b"" if method=="POST" else None
                    with urlopen(Request(base+path,method=method,data=data,
                                 headers={"Authorization":"Bearer "+key}),context=tls,timeout=3) as response:
                        return response.status,json.load(response)
                self.assertEqual(request("/accounts")[1],{"accounts":[A,B],"default":A})
                for bad, status in [("/status",404),("/status?account="+A+"&account="+B,404)]:
                    with self.assertRaises(HTTPError) as e: request(bad)
                    self.assertEqual(e.exception.code,status)
                with self.assertRaises(HTTPError) as e: request("/accounts",key="bad")
                self.assertEqual(e.exception.code,401)
                with self.assertRaises(URLError): urlopen(base+"/accounts",timeout=3)
                self.assertEqual(request("/scan?account="+B,"POST")[0],202)
                self.assertEqual(request("/status?account="+A)[1]["account"],A)
                self.assertEqual(request("/preview?account="+B)[1]["account"],B)
                self.assertEqual(self.b.posts,[])
                self.assertFalse(w.directory(A).exists())
                self.assertTrue(w.directory(B).exists())
            finally:
                server.shutdown()
                worker.join(3)
