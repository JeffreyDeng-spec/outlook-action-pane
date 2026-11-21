"""Offline fixtures for the actual TypeSafe Jev HTTP request serializer."""
from contextlib import closing
import io,json,sqlite3
from unittest import TestCase
from unittest.mock import patch

from outlook_triage import JEV, QUESTIONS, PreviewError, classify, init_cache, preview_rows

DUMMY_TOKEN="DUMMY_TOKEN_NEVER_USED"

class Capture:
    def __init__(self):
        self.requests=[]
    def open(self,req,*,timeout):
        self.requests.append((req.full_url,req.get_method(),dict(req.header_items()),json.loads(req.data),timeout))
        # Deliberately incomplete fake response: we're validating serialization, not model predictions.
        return io.BytesIO(b'{"answers":{}}')

def serialize(email):
    capture=Capture()
    with patch("outlook_triage.build_opener",return_value=capture), patch(
            "socket.create_connection",side_effect=AssertionError("Offline test; no external connections")):
        try:
            classify(email,DUMMY_TOKEN,"jev-latest")
        except PreviewError:
            pass
    assert len(capture.requests)==1
    return capture.requests[0]

class JevPayloadTests(TestCase):
    def fixture(self,**changes):
        sample={"id":"DUMMY_MESSAGE_ID","source":"Inbox","receivedDateTime":"2026-11-19T00:00:00Z",
                "parentFolderId":"DUMMY_FOLDER_ID","changeKey":"DUMMY_CHANGE_KEY",
                "from":{"emailAddress":{"name":"DUMMY_FROM_NAME","address":"sender@example.test"}},
                "toRecipients":[{"emailAddress":{"address":"recipient@other.test"}}],
                "attachments":[{"name":"DUMMY_ATTACHMENT"}], "body":{"content":"DUMMY_FULL_BODY"},
                "internetMessageHeaders":[{"name":"LIST-Unsubscribe","value":"DUMMY_PRIVATE_HEADER"},
                                          {"name":"X-secret","value":"DUMMY_OTHER_HEADER"}],
                "subject":"Synthetic health note for Example Name: person@other.test",
                "bodyPreview":"Synthetic ID Z9999999; account 0000-1234; +1 555 0100; visit https://demo.test/reset?token=DUMMY_QUERY",
                **changes}
        return sample

    def test_exact_transmitted_json_and_sensitive_limitations(self):
        url,method,headers,data,_=serialize(self.fixture())
        self.assertEqual((url,method),(JEV,"POST"))
        self.assertEqual(headers["Authorization"],"Bearer "+DUMMY_TOKEN)
        self.assertEqual(headers["Content-type"],"application/json")
        self.assertEqual(data,{"model":"jev-latest","questions":QUESTIONS,"state":{
            "sender_domain":"example.test",
            "subject":"Synthetic health note for Example Name: [email]",
            "body_preview":"Synthetic ID Z9999999; account 0000-1234; +1 555 0100; visit [link domain: demo.test]",
            "mailing_list_header_present":True}})
        for absent in ("DUMMY_MESSAGE_ID","DUMMY_FOLDER_ID","DUMMY_CHANGE_KEY","DUMMY_ATTACHMENT",
                       "DUMMY_FULL_BODY","DUMMY_PRIVATE_HEADER","DUMMY_OTHER_HEADER","DUMMY_FROM_NAME",
                       "recipient@other.test","sender@example.test","person@other.test","DUMMY_QUERY"):
            self.assertNotIn(absent,json.dumps(data))

    def test_subject_and_preview_character_limits_and_no_list_header_flag(self):
        _,_,_,data,_=serialize(self.fixture(
            subject="甲"*245,bodyPreview="乙"*260,
            internetMessageHeaders=[{"name":"X-List-Id","value":"DUMMY_HEADER"}]))
        self.assertEqual(data["state"],{"sender_domain":"example.test","subject":"甲"*240,
                                        "body_preview":"乙"*255})

    def test_uncached_sensitive_looking_message_reaches_provider_before_any_review(self):
        sample=self.fixture(subject="Synthetic hospital report; Test Person; passport Q000001",
                           bodyPreview="Synthetic tax identifier: 999-99-9999; bank payment reference MOCK-999")
        capture=Capture()
        with closing(sqlite3.connect(":memory:")) as db, patch(
                "outlook_triage.build_opener",return_value=capture), patch(
                "socket.create_connection",side_effect=AssertionError("No networking")):
            init_cache(db)
            preview_rows(db,[sample],DUMMY_TOKEN,"jev-latest")
        self.assertEqual(len(capture.requests),1)
        payload=capture.requests[0][3]["state"]
        self.assertIn("Test Person; passport Q000001",payload["subject"])
        self.assertIn("999-99-9999",payload["body_preview"])
        self.assertIn("MOCK-999",payload["body_preview"])
