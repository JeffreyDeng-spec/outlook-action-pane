from unittest import TestCase
from unittest.mock import Mock,patch
from windows_jev_guard import make_guarded_classifier

def msg(subject="Weekend garden promotion",preview="Save 15% on a terracotta pot!"):
    return {"subject":subject,"bodyPreview":preview,
        "from":{"emailAddress":{"address":"promos@example.test"}},
        "internetMessageHeaders":[{"name":"List-Id","value":"generic"}]}

class GuardTests(TestCase):
    def test_sensitive_bypasses_provider(self):
        delegate=Mock(side_effect=AssertionError("Provider must not be called"))
        route=make_guarded_classifier(delegate)
        examples=[
            msg("Medical lab result","Synthetic report"),
            msg("体检报告","模拟"),
            msg("Your bank statement ready",""),
            msg("到账","银行卡：模拟"),
            msg("New sign-in; password reset",""),
            msg("验证","验证码 001001"),
            msg("Shipping confirmation","example: access_token=FAKE000"),
            msg("Account correspondence","Passport K000000"),
            msg("Other","999-99-9999"),
            msg("Other","请看身份证号码：１１１１１１１１１１１１１１１１１１"),
            msg("Note","See +1 555-555-0123"),
            msg("Normal looking subject","x"*400+" historical bank account number")
        ]
        with patch("socket.create_connection",side_effect=AssertionError("Offline guard test")):
            for mail in examples:
                with self.subTest(subject=mail["subject"]):
                    result=route(mail,"DUMMY","test")
                    self.assertEqual(result["target"],"review")
                    self.assertEqual(result["model"],"local-no-provider")
        delegate.assert_not_called()
    def test_malformed_bypasses_provider_and_ordinary_promos_keep_rules(self):
        selected={"target":"Junk","reason":"mock promo"}
        delegate=Mock(return_value=selected)
        route=make_guarded_classifier(delegate)
        malformed=[None,{**msg(),"bodyPreview":123},{**msg(),"subject":"bad\0control"},
                        {**msg(),"from":{"emailAddress":{"address":"not an email"}}},
                        {**msg(),"internetMessageHeaders":{"name":"bad"}}]
        for actual in malformed:
            self.assertEqual(route(actual,"DUMMY","test")["target"],"review")
        email=msg()
        self.assertIs(route(email,"DUMMY","test"),selected)
        delegate.assert_called_once_with(email,"DUMMY","test")
