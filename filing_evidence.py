"""Use a bounded excerpt of the original plain-text message for ambiguous filing."""
import re
import unicodedata
from urllib.parse import quote

from outlook_triage import GRAPH, IMMUTABLE_ID, PreviewError, scrub
from windows_jev_guard import local_review

PADDING = re.compile(r"[\u034f\u200b-\u200f\u2060\ufeff]")
PROGRAM = re.compile(r"会员|会籍|礼遇|积分|计划|升级|loyalty|rewards?|membership|program(?:me)?|benefits?|tier", re.I)
BOILERPLATE = re.compile(r"view (?:this )?(?:email )?in (?:a |your )?browser|learn more|see what.s new|important update|see details|了解(?:更多|最新)|查看详情|重要通知", re.I)
QUOTED = re.compile(r"^(?:from:|发件人[：:]|on .{2,180} wrote:|-----original message-----)", re.I | re.M)
PERSONAL = re.compile(
    r"(?:预订|订单|确认)(?:号|编号)|(?:您的|你的)(?:预订|账单|退款|会籍|会员等级|积分即将到期)"
    r"|(?:booking|reservation|invoice)\s*(?:number|code|confirmed|#)|amount due|your (?:booking|reservation|invoice)"
    r"|your (?:membership|tier|status) (?:has|will|expires)", re.I)
OPAQUE = re.compile(r"(?<![\w-])[A-Za-z0-9_~=/+-]{30,}(?![\w-])")
REFERENCE = re.compile(r"((?:membership|invoice|booking|reservation|confirmation)\s*(?:number|code|#)|会员号|预订号|订单号)\s*[:：#-]?\s*[A-Z0-9-]*\d[A-Z0-9-]*", re.I)
GREETING = re.compile(r"(?im)^(?:dear [^\n,]{1,65},|尊敬的[^\n，,:：]{1,40}(?:先生|女士)[^\n，,:：]{0,8}[，,:：]?)")
MAX_EXCERPT = 1200


def before_footer(raw):
    lines = []
    for line in raw.splitlines():
        if (re.match(r"^\s*(?:unsubscribe\b|取消订阅|退订此邮件|privacy policy\b|隐私政策)", line, re.I)
                or re.match(r"^\s*(?:为了?保障您的隐私.{0,20}请勿.{0,18}发送|for your (?:privacy|security).{0,35}do not send)", line, re.I)
                or (re.search(r"privacy|隐私", line, re.I)
                    and re.search(r"unsubscribe|contact us|terms of use|退订|联系我们|使用条款", line, re.I)
                    and len(re.findall(r"https?://", line, re.I)) >= 2)):
            break
        lines.append(line)
    return "\n".join(lines)


def clean(text):
    return PADDING.sub("", unicodedata.normalize("NFKC", text))


def weak_preview(message):
    raw = message.get("bodyPreview") or ""
    text = " ".join(clean(raw).split())
    return (not text or len(raw) - len(text) >= 60 or
            (len(text) < 100 and (BOILERPLATE.search(text) or PROGRAM.search(message["subject"] + " " + text))))


def hold(reason):
    # Empty scores also preserve folder placement during Junk-side reconciliation.
    return {"target": "review", "reason": reason, "model": "local-no-provider", "scores": [], "signals": {}}


def classify_with_body(message, key, model, read, token, classifier):
    initial = None
    if not weak_preview(message):
        initial = classifier(message, key, model)
        conflict = initial["target"] == "Junk" and PERSONAL.search(message["subject"] + " " + message["bodyPreview"])
        if initial["target"] != "review" and not conflict:
            return initial

    ident = message["id"]
    url = f"{GRAPH}/me/messages/{quote(ident, safe='')}?%24select=id,parentFolderId,uniqueBody"
    try:
        expanded = read(url, "Graph classification evidence", token=token, headers=IMMUTABLE_ID)
    except PreviewError as error:
        if error.http_status == 404:
            return hold("邮件已变化，保留当前位置")
        raise  # Leave queued for retry on temporary Graph failures.
    body = expanded.get("uniqueBody") or {}
    if (expanded.get("id") != ident or not message.get("parentFolderId")
            or expanded.get("parentFolderId") != message["parentFolderId"]
            or str(body.get("contentType")).lower() != "text" or not isinstance(body.get("content"), str)):
        return hold("无法验证当前邮件的纯文本正文")
    # Inspect the new material locally before it can be included in a provider request.
    raw = QUOTED.split(clean(body["content"]), maxsplit=1)[0]
    raw = "\n".join(line for line in raw.splitlines() if not line.lstrip().startswith(">"))
    if not raw.strip() or len(raw) > 100000:
        return hold("缺少可核实的正文")
    # Examine the substantive message, excluding conventional footer contact details.
    main = before_footer(raw)
    safe = local_review({**message, "bodyPreview": main})
    if safe:
        return safe
    text = GREETING.sub("", scrub(main))
    text = " ".join(OPAQUE.sub("[id]", REFERENCE.sub(r"\1 [id]", text)).split())
    if not text:
        return hold("缺少可核实的正文")
    excerpt = text[:MAX_EXCERPT]
    personal = PERSONAL.search(text)
    if personal and personal.end() > MAX_EXCERPT:
        excerpt = text[:850] + " … " + text[max(0, personal.start()-60):personal.start()+280]
    if initial and initial["target"] == "review" and text == " ".join(clean(scrub(message["bodyPreview"])).split()):
        return initial
    result = classifier({**message, "filing_excerpt": excerpt}, key, model)
    if result["target"] == "Junk" and (personal or PERSONAL.search(message["subject"]) or PERSONAL.search(raw[len(main):])):
        return hold("正文包含个人交易或会籍事项，保留当前位置")
    return result
