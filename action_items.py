"""Turn grounded requests into per-mailbox actions. Messages are never executed."""
from datetime import datetime, timedelta, timezone
import re
from urllib.parse import urlsplit
from uuid import uuid4

from outlook_triage import clears, domain, parse_received, utcstamp
from windows_jev_guard import must_review_locally


def outlook_link(value):
    try:
        p = urlsplit(value or "")
        if (p.scheme == "https" and p.netloc in {
            "outlook.live.com", "outlook.office.com", "outlook.office365.com", "outlook.cloud.microsoft"
        } and not any(c.isspace() for c in value)):
            return value
    except (TypeError, ValueError):
        pass
    return ""


VERIFY = re.compile(r"verification code|one[- ]time (?:password|passcode|code)|your (?:login|security|authentication) code|verify your (?:email|account)|验证码|动态口令|验证(?:你的|您的)?邮箱", re.I)
INTERVIEW = re.compile(r"\binterview\b|面试", re.I)
INVITE = re.compile(r"\binvit(?:e|ation)|\b(?:select|choose|schedule|book|pick|confirm|offer|available|availability|propose)\b|邀请|请选择|请确认|选择.{0,10}时[段间]|预约|哪些时间|是否方便", re.I)
ASSESS = re.compile(r"\b(?:assessment|test assignment|coding challenge)\b|测评|笔试|在线测试", re.I)
DOCUMENTS = re.compile(r"\b(?:transcript|documents?|portfolio|resume|references?)\b|成绩单|材料|作品集|简历")
REQUEST = re.compile(r"\b(?:please|kindly|you (?:must|need to|are required to)|can you|could you|we ask (?:that )?you)\b|请(?:您|你)?|须(?:要)?|需要(?:你|您)|务必", re.I)
TASK_VERB = re.compile(r"\b(?:complete|take|finish|submit|upload|provide|send us|return|fill out)\b|完成|上传|补交|补充|提供|提交|填写|参加", re.I)
REPLY = re.compile(r"\b(?:reply|respond|get back to (?:me|us)|confirm receipt|acknowledge receipt)\b|回复|回信|回覆|告知.{0,8}时间", re.I)
NEGATIVE = re.compile(r"\b(?:no (?:action|reply|response) (?:is )?(?:required|needed)|no need to|do not|don't|not required to|already (?:received|submitted|completed|confirmed))\b|无需|无须|不必|请勿|请不要|不用|已(?:经)?(?:完成|收到|提交|确认|过期|失效)", re.I)
EXPIRY = re.compile(r"(?:有效(?:期)?|失效|valid (?:for|within)|expires? (?:in|after)|within)\s*[:：为]?\s*(\d{1,4})\s*(分钟|小时|minutes?|hours?)|(\d{1,4})\s*(分钟|小时)\s*(?:内有效|后(?:失效|过期))", re.I)
DUE = re.compile(r"(?:本?周[一二三四五六日天]|星期[一二三四五六日天]|今天|明天|\d{1,2}月\d{1,2}日)(?:[上下午晚]午?)?\s*(?:\d{1,2}[:：]\d{2}|\d{1,2}点)?\s*(?:之?前|截止)|\bby (?:this |next )?(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|\d{4}-\d\d-\d\d)(?: at \d{1,2}(?::\d{2})?(?:\s?[ap]m)?)?", re.I)
OTHER = re.compile(r"\b(?:sign|approve|renew|review|update|choose|select|confirm|schedule|pay)\b|签署|批准|审核|调整|续费|缴费|选择|确认|预约", re.I)
SAFE_RECEIPT = re.compile(r"confirm (?:that you (?:have )?received|receipt)|acknowledge receipt|确认(?:已经?|已)?收到|回复.{0,4}收悉", re.I)
ASSESSMENT_REQUEST = re.compile(r"\b(?:complete|finish|take)\b.{0,48}\b(?:assessment|coding challenge)\b|(?:完成|参加).{0,16}(?:测评|笔试|在线测试)", re.I)
EXTERNAL_STEP = re.compile(r"https?[:/]|www\.|\b[a-z0-9-]+\.[a-z]{2,}\b|"
                           r"\b(?:link|click|download|install|attach(?:ed|ment)?|upload|password|pay|payment|transfer|gift card|log[- ]?in|sign[- ]?in)\b|"
                           r"链接|点击|下载|安装|附件|上传|密码|付款|支付|转账|缴费|登录|证件|银行卡", re.I)


def sentences(message):
    # uniqueBody is requested as text. Strip quotes and never store full bodies.
    body = message.get("uniqueBody") or {}
    text = body.get("content") if str(body.get("contentType", "")).lower() == "text" else None
    text = text or message.get("bodyPreview", "")
    text = re.split(r"(?im)^(?:from:|发件人[：:]|on .{2,180} wrote:|-----original message-----)", text, maxsplit=1)[0]
    text = "\n".join(line for line in text.splitlines() if not line.lstrip().startswith(">"))
    return [s.strip()[:750] for s in re.split(
        r"(?<=[。！？!?])|(?<=\.)\s+|[\r\n]+|[，,；;]\s*(?=(?:但|请|please\b|however\b))",
        text[:32768], flags=re.I) if s.strip()]


def eligible_inbox(row):
    scores = row.get("scores", [])
    return (row["source"] == "Inbox" and row["target"] != "Junk"
            and not (scores and (scores[0][0] == "marketing" or
                                 (scores[0][0] == "junk" and clears(scores[0], "junk", .90)))))


def needs_full_text(message, row):
    preview = message.get("bodyPreview", "")
    if not eligible_inbox(row) or not isinstance(preview, str) or len(preview) < 240:
        return False
    # Guarded mail never gets sent to the classifier. Keep the same narrow handling locally.
    guard = must_review_locally(message)
    return not guard or (guard == "sensitive" and bool(VERIFY.search(message.get("subject", "") + " " + preview)))


def action_for(message, row):
    if not eligible_inbox(row):
        return None
    subject = message.get("subject", "")
    parts = sentences(message)
    content = "\n".join([subject, *parts])
    sensitive = must_review_locally({**message, "bodyPreview": content})
    if sensitive == "invalid":
        return None
    eligible = [line for line in [subject, *parts] if not NEGATIVE.search(line)
                or (REQUEST.search(line) and SAFE_RECEIPT.search(line)
                    and not re.search(r"无需|不必|不用|请勿|no need|do not|don't", line, re.I))]
    if any(re.search(r"无需(?:操作|处理|行动|回复)|no (?:further )?(?:action|response|reply) "
                     r"(?:is )?(?:required|needed)", p, re.I) for p in parts):
        # A subject can remain unchanged across many replies in a thread.
        eligible = [p for p in eligible if p in parts]
    reply_subject = re.match(r"^\s*(?:re|fw|fwd|回复|转发)\s*[:：]", subject, re.I)
    code = any(VERIFY.search(line) for line in eligible if not reply_subject or line in parts)
    result = dict(source=domain(message), received=message["receivedDateTime"], due="", expires="", reply="")
    scores = row.get("scores", [])
    required = len(scores) > 1 and scores[1][0] == "required" and clears(scores[1], "required")
    if scores and scores[0][0] in {"junk", "unclear"}:
        # A weak nature vote can disagree with a concrete request. Allow only
        # a local reminder to check the user's existing portal, with no delegated side effect.
        request = next((line for line in eligible
                        if REQUEST.search(line) and ASSESSMENT_REQUEST.search(line)), "")
        if (request and (required or DUE.search(request)) and not sensitive
                and not code and not EXTERNAL_STEP.search(content)):
            due = DUE.search(content)
            return dict(result, kind="material", title="到申请平台核对在线测评",
                        due=due[0].strip() if due else "",
                        detail="请从你原先使用的申请平台查看是否有待完成的测评。")
        return None
    if code and not re.search(r"(?:code|passcode) (?:has |is )?expired|验证码已(?:经)?(?:过期|失效)", content, re.I):
        match = EXPIRY.search(content)
        if match:
            number, unit = match.groups()[:2] if match[1] else match.groups()[2:]
            minutes = int(number) * (60 if unit.lower().startswith(("h", "小时")) else 1)
            result["expires"] = utcstamp(parse_received(result["received"]) + timedelta(minutes=minutes))
            result["due"] = "收到邮件后 " + str(minutes) + " 分钟内"
        return dict(result, kind="verify", title="完成邮箱或登录验证", detail="如果你刚刚发起了这项操作，请打开原邮件，并在你使用的官方页面完成验证。")
    if sensitive:
        return None
    explicit = [line for line in eligible if REQUEST.search(line)]
    for line in eligible:
        if INTERVIEW.search(content) and INVITE.search(line) and (line in explicit or required or INTERVIEW.search(line)):
            kind, title = "interview", "选择或确认面试时间"
        elif (line in explicit or required) and TASK_VERB.search(line) and ASSESS.search(content):
            kind, title = "material", "完成在线测评"
        elif (line in explicit or required) and TASK_VERB.search(line) and DOCUMENTS.search(content):
            kind, title = "material", "提交" + ("成绩单" if re.search(r"transcript|成绩单", content, re.I) else "所需材料")
        elif line in explicit and REPLY.search(line):
            kind, title = "reply", "回复来信"
            draft_ok = (required and row["target"] == "Inbox" and scores[0][0] == "legitimate"
                        and clears(scores[0], "legitimate"))
            if draft_ok and SAFE_RECEIPT.search(line):
                result["reply"] = "您好，我已收到这封邮件。谢谢！" if re.search("[\u4e00-\u9fff]", line) else "Thank you for your email. I confirm receipt."
                title = "回复确认收悉"
            elif draft_ok:
                result["reply"] = ("您好，感谢来信。关于您提到的问题，[请在此填写答复]。"
                                   if re.search("[\u4e00-\u9fff]", line) else
                                   "Thank you for your email. In response to your question, [add your answer here].")
        elif required and line in explicit and (TASK_VERB.search(line) or OTHER.search(line)):
            kind, title = "action", line
        else:
            continue
        # Display mail excerpts as text. Strip links (including login tokens), never follow them.
        clean = lambda value: re.sub(r"https?://\S+", "（见原邮件）", value, flags=re.I)
        due = DUE.search(content)
        detail = clean(line)
        if kind == "interview":
            detail = clean(" ".join([line] + [s for s in parts if s != line and re.search(
                r"\d[:：]\d{2}|北京时间|时区|\b(?:UTC|GMT|EST|EDT|PST|PDT)\b", s, re.I)]))
            detail += " 请核对日历后选择时段。"
        return dict(result, kind=kind, title=clean(title)[:150], due=due[0].strip() if due else "", detail=detail[:1100])
    return None


def ensure_actions(db):
    db.execute("CREATE TABLE IF NOT EXISTS action_backfill(message TEXT PRIMARY KEY)")
    db.execute("""CREATE TABLE IF NOT EXISTS message_actions (
        message TEXT PRIMARY KEY, id TEXT UNIQUE NOT NULL, kind TEXT NOT NULL,
        title TEXT NOT NULL, source TEXT NOT NULL, received TEXT NOT NULL, due TEXT NOT NULL,
        expires TEXT NOT NULL, detail TEXT NOT NULL, mail_url TEXT NOT NULL,
        state TEXT NOT NULL DEFAULT 'open', reply TEXT NOT NULL,
        draft_state TEXT NOT NULL DEFAULT '', draft_url TEXT NOT NULL DEFAULT ''
    )""")


def record_action(db, message, row, *, historical=False):
    action = action_for(message, row)
    if action:
        db.execute("""INSERT INTO message_actions(message,id,kind,title,source,received,due,
            expires,detail,mail_url,reply) VALUES (?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(message) DO NOTHING""",
            (message["id"], uuid4().hex, action["kind"], action["title"], action["source"],
             action["received"], action["due"], action["expires"], action["detail"],
             outlook_link(message.get("webLink")), "" if historical else action["reply"]))


def open_actions(db):
    fields = "id kind title source received due detail mail_url draft_url".split()
    rows = db.execute("SELECT " + ",".join(fields) + ",draft_state FROM message_actions "
                      "WHERE state='open' AND (expires='' OR expires>?) "
                      "ORDER BY CASE WHEN expires<>'' THEN 0 WHEN due<>'' THEN 1 ELSE 2 END, received DESC",
                      (utcstamp(datetime.now(timezone.utc)),))
    items = []
    for row in rows:
        item = dict(zip(fields, row))
        if row[-1] == "creating":
            item["detail"] += " 草稿创建结果待核对，请先查看 Outlook 草稿箱。"
        items.append(item)
    return items


def update_action(db, ident, command):
    state = {"done": "done", "ignore": "ignored", "undo": "open"}[command]
    with db:
        if not db.execute("UPDATE message_actions SET state=? WHERE id=?", (state, ident)).rowcount:
            raise KeyError("Unknown action")


def prepare_replies(db, adapter, token):
    for message, reply in db.execute("""SELECT message,reply FROM message_actions
        WHERE state='open' AND draft_state='' AND reply<>'' ORDER BY received""").fetchall():
        with db:
            claimed = db.execute("""UPDATE message_actions SET draft_state='creating'
                WHERE message=? AND draft_state='' AND state='open'""", (message,)).rowcount
        if not claimed:
            continue
        # A request may have been accepted when the connection died. Keep 'creating' for human reconciliation.
        draft = adapter.create_reply(message, reply, token)
        if not (draft.get("isDraft") is True and draft.get("id") and outlook_link(draft.get("webLink"))):
            continue
        with db:
            db.execute("UPDATE message_actions SET draft_state='ready',draft_url=? WHERE message=?",
                       (outlook_link(draft["webLink"]), message))
