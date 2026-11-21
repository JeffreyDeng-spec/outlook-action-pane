from unittest import TestCase
from unittest.mock import Mock

from filing_evidence import classify_with_body
from outlook_triage import IMMUTABLE_ID, PreviewError, classification_state, decision_key


# Same ambiguous preheader, different recipients and senders. These examples are
# also used to replay the policy against the configured classification provider.
EXAMPLES = (
    ("sky.example.test", "Rewards news", "All Sky Club members will earn double points under the new program rules next spring.", "Junk"),
    ("shops.example.test", "会员计划更新", "从明年开始，全体会员均可按新的比例累计积分。会员等级门槛将统一调整。", "Junk"),
    ("sky.example.test", "Rewards news", "Your reservation for Tuesday is confirmed. Booking number: ZQ8271. Members can earn extra points.", "Inbox"),
    ("shops.example.test", "会员计划更新", "本月您的账单已出具，本次入住应付金额为 280 元。会员商城另有优惠。", "Inbox"),
    ("sky.example.test", "Rewards news", "Your membership has been upgraded to Gold after your qualifying stays. Your new tier is already active.", "Inbox"),
)


def mail(domain, subject, preview="Learn more."):
    return {"id": "item", "parentFolderId": "inbox", "changeKey": "v1", "subject": subject,
            "bodyPreview": preview, "from": {"emailAddress": {"address": "mail@" + domain}}}


def response(text):
    return {"id": "item", "parentFolderId": "inbox", "uniqueBody": {"contentType": "text", "content": text}}


class FilingEvidenceTests(TestCase):
    def test_scoped_read_and_clean_payload_for_each_example(self):
        for domain, subject, text, expected in EXAMPLES:
            with self.subTest(domain=domain, expected=expected):
                read = Mock(return_value=response(text + "\nhttps://news.example.test/go?access=private\n"
                                                  "From: Previous Sender\nprivate-old-thread"))
                classifier = Mock(return_value={"target": expected})
                result = classify_with_body(mail(domain, subject), "key", "model", read, "token", classifier)
                self.assertEqual(result["target"], expected)
                read.assert_called_once()
                self.assertEqual(read.call_args.kwargs["headers"], IMMUTABLE_ID)
                classifier.assert_called_once()
                excerpt = classification_state(classifier.call_args.args[0])["body_excerpt"]
                for hidden in ("access=private", "ZQ8271", "private-old-thread"):
                    self.assertNotIn(hidden, excerpt)
                self.assertLessEqual(len(excerpt), 1200)

    def test_clear_preview_needs_no_body_uncertain_preview_gets_second_pass(self):
        message = mail("b.example.test", "An ordinary message", "A detailed and unambiguous description of the delivery.")
        read = Mock(return_value=response("The package you ordered was delivered to your door."))
        classifier = Mock(return_value={"target": "Inbox"})
        classify_with_body(message, "", "", read, "", classifier)
        read.assert_not_called()
        classifier.side_effect = [{"target": "review"}, {"target": "Inbox"}]
        self.assertEqual(classify_with_body(message, "", "", read, "", classifier)["target"], "Inbox")
        read.assert_called_once()
        self.assertNotEqual(decision_key(message, "m"), decision_key(dict(message, changeKey="v2"), "m"))

    def test_sensitive_stale_or_incomplete_body_stays_local(self):
        message = mail("b.example.test", "Program update", "\u034f\u200c" * 60)
        classifier = Mock()
        for bad in (response("Verification code: 000000"), dict(response("hello"), id="other"),
                    dict(response("hello"), parentFolderId="different"), response("")):
            self.assertEqual(classify_with_body(message, "", "", Mock(return_value=bad), "", classifier)["target"], "review")
        read = Mock(side_effect=PreviewError("timeout"))
        with self.assertRaises(PreviewError):
            classify_with_body(message, "", "", read, "", classifier)
        classifier.assert_not_called()

    def test_transaction_deep_in_body_vetoes_a_false_promotion(self):
        text = "Details about earning points. " * 90 + " Your invoice number A1834 is due tomorrow."
        classifier = Mock(return_value={"target": "Junk"})
        result = classify_with_body(mail("b.example.test", "Rewards"), "", "", Mock(return_value=response(text)), "", classifier)
        self.assertEqual(result["target"], "review")
        self.assertIn("invoice", classifier.call_args.args[0]["filing_excerpt"].lower())
