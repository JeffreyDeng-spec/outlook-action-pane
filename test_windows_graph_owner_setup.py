import json
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from unittest.mock import patch
from uuid import uuid4
from types import SimpleNamespace
from mail_accounts import GRAPH_KEYCHAINS
from outlook_triage import PreviewError
from windows_graph_owner_setup import roots,run,provision_folders
A,B=tuple(GRAPH_KEYCHAINS)

class OwnerSetupTests(TestCase):
    def test_sequential_owner_flow_and_minimal_record(self):
        calls=[]
        with TemporaryDirectory() as profile, patch("socket.create_connection",side_effect=AssertionError):
            auth,state=roots(profile)
            id=str(uuid4())
            def enroll(client,account,store):
                self.assertEqual(client,id)
                self.assertEqual(store,auth)
                calls.append(("enroll",account))
                if account==B:
                    raise PreviewError("TEST_ONLY_SENSITIVE_NOT_RECORDED")
            def builder(args,*,provider):
                self.assertEqual(args.mode,"graph-metadata")
                self.assertEqual(args.credential_dir,auth)
                self.assertEqual(args.state_dir,state)
                self.assertTrue(args.check_once)
                self.assertIsNone(args.imap_client_id)
                self.assertIsNone(args.tls_cert)
                return SimpleNamespace(scan=lambda a: calls.append(("scan",a))),None
            self.assertEqual(run(id,profile,enroll_account=enroll,provider=lambda p:p,builder=builder),1)
            self.assertEqual(calls,[("enroll",A),("enroll",B),("scan",A)])
            raw=(state/"verification.json").read_text(encoding="utf-8")
            self.assertNotIn("TEST_ONLY_SENSITIVE",raw)
            self.assertEqual(json.loads(raw)["accounts"][A],
                  {"authorized":True,"identity_inbox_junk_delta":True})
            self.assertFalse(json.loads(raw)["accounts"][B]["authorized"])

    def test_refuses_to_reuse_unmarked_local_directory(self):
        with TemporaryDirectory() as profile:
            auth,state=roots(profile)
            state.mkdir(parents=True)
            sentinel=state/"original.bin"
            sentinel.write_bytes(b"DUMMY ORIGINAL")
            with self.assertRaises(ValueError):
                provision_folders(auth,state)
            self.assertEqual(sentinel.read_bytes(),b"DUMMY ORIGINAL")
