"""Opt-in production candidate: obvious sensitive content stays in local review.

No integration is enabled by importing this module. It is a heuristic, not a
promise of complete detection or anonymization. No message text enters logs.
"""
from __future__ import annotations
import re
import unicodedata

RULES=(
    re.compile(r"\b(?:password\s*(?:reset|change)|reset\s+(?:your\s+)?password|verification\s*(?:code|number)|"
               r"security\s*(?:alert|code|challenge)|account\s*(?:recovery|compromised|takeover)|"
               r"one[- ]time\s*(?:password|code|passcode)|two[- ]factor|multifactor|"
               r"unauthori[sz]ed\s+(?:login|sign[- ]in|payment|transaction)|"
               r"(?:new|unusual)\s+(?:login|sign[- ]in)|otp|2fa|mfa|"
               r"(?:access|refresh|recovery|reset)[_ -]?token\s*[=:]|"
               r"(?:api[_ -]?key|client[_ -]?secret|recovery\s*phrase|private\s*key)\b)|"
               r"(?:验证码|动态口令|密码重置|重置密码|登录异常|安全警报|安全提醒|身份验证|恢复码)",re.I),
    re.compile(r"\b(?:medical|diagnosis|diagnostic|hospital|health\s*(?:record|care)|medical\s*record|"
               r"patient|prescription|blood\s*test|lab\s*(?:result|test)|clinic|medication|mental\s*health)\b|"
               r"(?:医疗|医院|就诊|诊断|病历|处方|化验|体检|检验报告|药物)",re.I),
    re.compile(r"\b(?:bank\s*(?:account|statement|transfer|payment)|"
               r"(?:credit|debit|payment)\s*card|credit\s*(?:score|report)|"
               r"wire\s*transfer|routing\s*number|iban|account\s*(?:balance|number)|"
               r"tax\s*(?:id|identifier|return|number|statement)|loan\s*(?:application|statement)|"
               r"payment\s*(?:reference|confirmation)|statement\s*balance|pay\s*stub)\b|"
               r"(?:银行账户|银行账号|银行卡|银行流水|银行账单|转账|付款凭证|支付凭证|信用卡|借记卡|还款|税号|工资单|报税|征信|社保号)",re.I),
    re.compile(r"\b(?:passport|social\s*security|ssn|sin\s*(?:number|identification)|"
               r"(?:national|government)\s*(?:identity|id)|taxpayer\s*identification|"
               r"driver(?:'s|’s|s)?\s*licen[cs]e|identification\s*(?:card|number))\b|"
               r"(?:护照|身份证|证件号码|公民身份号码|驾驶证|个人身份号码)",re.I),
    re.compile(r"(?<!\w)\d{3}[-\s]\d{2}[-\s]\d{4}(?!\w)|"
               r"(?<!\w)\d{3}[-\s]\d{3}[-\s]\d{3}(?!\w)|"
               r"(?<![0-9A-Z])[0-9]{17}[0-9X](?![0-9A-Z])|"
               r"(?<!\d)(?:\d[ -]?){12,18}\d(?!\d)|"
               r"(?:\+?1[- .]?)?\(?\d{3}\)?[- .]\d{3}[- .]\d{4}\b",re.I),
)
ADDRESS=re.compile(r"[^\s@<>]+@[^\s@<>]+[.][^\s@<>]+")

def must_review_locally(message):
    if not isinstance(message,dict):
        return "invalid"
    for name in ("subject","bodyPreview"):
        value=message.get(name)
        if not isinstance(value,str) or any(ord(c)<32 and c not in "\t\n\r" for c in value):
            return "invalid"
    sender=message.get("from")
    sender=sender.get("emailAddress") if isinstance(sender,dict) else None
    address=sender.get("address") if isinstance(sender,dict) else None
    if not isinstance(address,str) or not ADDRESS.fullmatch(address):
        return "invalid"
    headers=message.get("internetMessageHeaders",[])
    if not isinstance(headers,list) or any(not isinstance(h,dict) or not isinstance(h.get("name"),str)
                                          or not isinstance(h.get("value"),str) for h in headers):
        return "invalid"
    # Inspect original text before truncation or URL scrubbing. Raw headers are
    # deliberately not sent by the legacy classifier, nor used here as model data.
    content="\n".join((message["subject"],message["bodyPreview"],address.split("@",1)[-1]))
    text=unicodedata.normalize("NFKC",content)
    if any(pattern.search(text) for pattern in RULES):
        return "sensitive"
    return None

def local_review(message):
    why = must_review_locally(message)
    if why:
        reason = {
            "invalid": "邮件字段异常 · 仅本机待复核",
            "sensitive": "可能含敏感信息 · 仅本机待复核",
        }[why]
        return {"target": "review", "reason": reason, "scores": [],
                "signals": {}, "model": "local-no-provider"}
    return None


def make_guarded_classifier(delegate):
    """Use explicitly in the Windows rollout path; no change to legacy Mac code."""
    def guarded(message, key, model):
        return local_review(message) or delegate(message, key, model)
    return guarded
