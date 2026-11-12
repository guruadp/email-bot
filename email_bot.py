import os
import json
import time
import re
import html
import base64
import tempfile
from datetime import datetime, timezone
import requests
import msal
from dotenv import load_dotenv
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from urllib.parse import quote, urlencode
from zoneinfo import ZoneInfo

load_dotenv()


def log(*args, **kwargs):
    kwargs.setdefault("flush", True)
    print(*args, **kwargs)


def get_required_env(name):
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def get_bool_env(name, default=False):
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def infer_addressing_name(mailbox_user):
    local_part = (mailbox_user or "").split("@", 1)[0].strip()
    if not local_part:
        return "Guru"
    first_part = re.split(r"[._-]+", local_part)[0].strip()
    if not first_part:
        return "Guru"
    return first_part[:1].upper() + first_part[1:]


TENANT_ID = get_required_env("TENANT_ID")  # Directory (tenant) ID
CLIENT_ID = get_required_env("CLIENT_ID")  # Application (client) ID
CLIENT_SECRET = get_required_env("CLIENT_SECRET")  # Application client secret
MAILBOX_USER = get_required_env("MAILBOX_USER")  # User principal name or user id
DIRECT_TO_ADDRESS = MAILBOX_USER.strip().lower()

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]
MAILBOX_USER_ENCODED = quote(MAILBOX_USER, safe="")

POLL_SECONDS = 15
SEEN_IDS_FILE = os.getenv("SEEN_IDS_FILE", "seen_ids.json")
SEEN_IDS_RETENTION_DAYS = 30
MAX_THREAD_MESSAGES = int(os.getenv("MAX_THREAD_MESSAGES", "5"))
MAX_THREAD_CHARS_PER_MESSAGE = int(os.getenv("MAX_THREAD_CHARS_PER_MESSAGE", "1200"))
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
TEAMS_CHANNEL_EMAIL = get_required_env("TEAMS_CHANNEL_EMAIL").strip().lower()
DEV_MODE = get_bool_env("DEV_MODE", default=False)
ADDRESSING_NAME = (os.getenv("ADDRESSING_NAME") or "").strip() or infer_addressing_name(MAILBOX_USER)
EMAIL_SIGNATURE = get_required_env("EMAIL_SIGNATURE").replace("\\n", "\n").strip()
PERSONA_DESCRIPTION = (os.getenv("PERSONA_DESCRIPTION") or "").strip() or (
    "Guru Nandhan, CEO of Ednex LLC, Ednex Automation, and Maker and Coder"
)
VIP_SENDERS = {s.strip().lower() for s in (os.getenv("VIP_SENDERS") or "").split(",") if s.strip()}
VIP_DOMAINS = {d.strip().lower() for d in (os.getenv("VIP_DOMAINS") or "").split(",") if d.strip()}
N8N_ACTION_WEBHOOK_URL = (os.getenv("N8N_ACTION_WEBHOOK_URL") or "").strip()
LINK_EXPIRY_DAYS = int(os.getenv("LINK_EXPIRY_DAYS", "3"))
FALLBACK_REPLY_TEXT = (
    "Hi,\n\n"
    "Thanks for your email. I have received it and will get back to you shortly.\n\n"
    f"{EMAIL_SIGNATURE}"
)


def raise_for_status_with_details(resp, context):
    try:
        resp.raise_for_status()
    except requests.HTTPError as exc:
        try:
            details = json.dumps(resp.json(), indent=2)
        except ValueError:
            details = (resp.text or "").strip()
        raise RuntimeError(
            f"{context} failed with HTTP {resp.status_code}.\nResponse body:\n{details}"
        ) from exc


def decode_jwt_payload(token):
    parts = token.split(".")
    if len(parts) < 2:
        return {}
    payload = parts[1]
    padding = "=" * (-len(payload) % 4)
    try:
        decoded = base64.urlsafe_b64decode(payload + padding).decode("utf-8")
        return json.loads(decoded)
    except Exception:
        return {}


def warn_if_missing_mail_role(token):
    payload = decode_jwt_payload(token)
    roles = payload.get("roles") or []

    if "Mail.ReadWrite" not in roles:
        log("Warning: token is missing app role 'Mail.ReadWrite'.")
        log("Check Graph API permissions (Application) and admin consent.")


def build_http_session():
    retry = Retry(
        total=5,
        connect=5,
        read=5,
        status=5,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "POST", "PATCH"],
    )
    adapter = HTTPAdapter(max_retries=retry)
    session = requests.Session()
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


def load_seen_ids(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return {str(k): float(v) for k, v in data.items()}
    except (FileNotFoundError, ValueError, OSError, TypeError):
        pass
    return {}


def save_seen_ids(path, seen_ids):
    cutoff = time.time() - SEEN_IDS_RETENTION_DAYS * 86400
    pruned = {mid: ts for mid, ts in seen_ids.items() if ts >= cutoff}
    seen_ids.clear()
    seen_ids.update(pruned)
    try:
        dir_name = os.path.dirname(os.path.abspath(path)) or "."
        fd, tmp_path = tempfile.mkstemp(dir=dir_name, prefix=".seen_ids_", suffix=".tmp")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(seen_ids, f)
        os.replace(tmp_path, path)
    except OSError as e:
        log(f"Warning: failed to persist seen_ids: {e}")


def get_token():
    session = build_http_session()
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        http_client=session,
    )
    result = app.acquire_token_for_client(scopes=SCOPES)

    if "access_token" not in result:
        raise RuntimeError(f"Token failed: {json.dumps(result, indent=2)}")

    return result["access_token"]


class GraphAuth:
    def __init__(self, refresh_buffer_seconds=120):
        self.refresh_buffer_seconds = refresh_buffer_seconds
        self._token = None
        self._expires_at = 0

    def _refresh_token(self):
        session = build_http_session()
        app = msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=AUTHORITY,
            client_credential=CLIENT_SECRET,
            http_client=session,
        )
        result = app.acquire_token_for_client(scopes=SCOPES)
        token = result.get("access_token")
        if not token:
            raise RuntimeError(f"Token failed: {json.dumps(result, indent=2)}")
        expires_in = int(result.get("expires_in") or 3600)
        self._token = token
        self._expires_at = time.time() + expires_in
        return token

    def get_valid_token(self):
        now = time.time()
        if (not self._token) or (now >= self._expires_at - self.refresh_buffer_seconds):
            return self._refresh_token()
        return self._token

    def get_headers(self):
        return {"Authorization": f"Bearer {self.get_valid_token()}"}

    def refresh_now(self):
        return self._refresh_token()


def graph_get_with_auth_retry(session, auth, url, params=None, timeout=30):
    headers = auth.get_headers()
    resp = session.get(url, headers=headers, params=params, timeout=timeout)
    if resp.status_code == 401:
        auth.refresh_now()
        headers = auth.get_headers()
        resp = session.get(url, headers=headers, params=params, timeout=timeout)
    return resp, headers


def poll_inbox(session, auth, url, timeout=30):
    try:
        resp, headers = graph_get_with_auth_retry(
            session=session,
            auth=auth,
            url=url,
            timeout=timeout,
        )
        raise_for_status_with_details(resp, "Inbox poll")
        return resp.json().get("value", []), headers
    except requests.RequestException as e:
        log(f"Inbox poll failed: {e}")
        return None, None


def strip_html(text):
    if not text:
        return ""
    no_tags = re.sub(r"<[^>]+>", " ", text)
    plain = html.unescape(no_tags)
    return re.sub(r"\s+", " ", plain).strip()


def format_received_abudhabi(received):
    if not received:
        return "(unknown)"
    try:
        dt = datetime.fromisoformat(received.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        abu_dhabi_time = dt.astimezone(ZoneInfo("Asia/Dubai"))
        return abu_dhabi_time.strftime("%Y-%m-%d %H:%M:%S %Z")
    except Exception:
        return received


def sanitize_reply_text(text):
    if not text:
        return ""
    lines = text.splitlines()
    while lines and re.match(r"^\s*(subject|from|to|cc)\s*:", lines[0], flags=re.IGNORECASE):
        lines.pop(0)
    return "\n".join(lines).strip()


def extract_text_from_graph_message(message):
    body = (message.get("body") or {}).get("content", "")
    body_type = ((message.get("body") or {}).get("contentType") or "").lower()
    if body_type == "html":
        body = strip_html(body)
    if not body:
        body = message.get("bodyPreview") or ""
    return re.sub(r"\s+", " ", body).strip()


def build_thread_context(session, headers, conversation_id, current_message_id=None):
    if not conversation_id:
        return ""

    escaped_conversation_id = conversation_id.replace("'", "''")
    url = f"https://graph.microsoft.com/v1.0/users/{MAILBOX_USER_ENCODED}/messages"
    params = {
        "$filter": f"conversationId eq '{escaped_conversation_id}'",
        "$orderby": "receivedDateTime desc",
        "$top": str(MAX_THREAD_MESSAGES),
        "$select": "id,subject,from,receivedDateTime,body,bodyPreview",
    }
    resp = session.get(url, headers=headers, params=params, timeout=30)
    if resp.status_code == 400:
        try:
            error_code = ((resp.json() or {}).get("error") or {}).get("code")
        except ValueError:
            error_code = None
        if error_code == "InefficientFilter":
            # Some tenants reject filter+orderby for conversation queries.
            fallback_params = {
                "$filter": f"conversationId eq '{escaped_conversation_id}'",
                "$top": str(MAX_THREAD_MESSAGES),
                "$select": "id,subject,from,receivedDateTime,body,bodyPreview",
            }
            resp = session.get(url, headers=headers, params=fallback_params, timeout=30)

    if resp.status_code >= 400:
        try:
            raise_for_status_with_details(resp, "Fetch thread messages")
        except RuntimeError as e:
            log(f"Thread context skipped: {e}")
            return ""

    items = resp.json().get("value", [])

    items = sorted(items, key=lambda item: item.get("receivedDateTime") or "")
    chunks = []
    for item in items:
        if item.get("id") == current_message_id:
            continue
        sender = (item.get("from") or {}).get("emailAddress", {}).get("address", "unknown")
        received = item.get("receivedDateTime") or "unknown-time"
        subject = item.get("subject") or ""
        text = extract_text_from_graph_message(item)[:MAX_THREAD_CHARS_PER_MESSAGE]
        if text:
            chunks.append(
                f"From: {sender}\nReceived: {received}\nSubject: {subject}\nBody: {text}"
            )
    return "\n\n---\n\n".join(chunks)


def get_full_message_body(session, headers, message_id):
    url = f"https://graph.microsoft.com/v1.0/users/{MAILBOX_USER_ENCODED}/messages/{message_id}?$select=body,bodyPreview"
    resp = session.get(url, headers=headers, timeout=30)
    raise_for_status_with_details(resp, "Fetch message body")
    data = resp.json()
    body = (data.get("body") or {}).get("content", "")
    body_type = ((data.get("body") or {}).get("contentType") or "").lower()
    if body_type == "html":
        body = strip_html(body)
    if not body:
        body = data.get("bodyPreview") or ""
    return body[:4000]


def generate_reply_with_llm(session, sender, subject, body_text, thread_context=""):
    if not OPENAI_API_KEY:
        log("LLM fallback: OPENAI_API_KEY is not set.")
        return FALLBACK_REPLY_TEXT

    prompt = (
        "Write a professional email reply draft.\n"
        "Role and voice:\n"
        f"- You are {PERSONA_DESCRIPTION}\n"
        "- Write with clear executive tone: confident, concise, strategic, and respectful\n"
        "- Sound like a CEO responding directly to business stakeholders\n"
        "Requirements:\n"
        "- Keep it concise (80-160 words)\n"
        "- Use the conversation context for continuity when provided\n"
        "- Acknowledge the sender's request\n"
        "- Do not include any email headers (no Subject/From/To/Cc lines)\n"
        "- End with exactly this signature block:\n"
        f"{EMAIL_SIGNATURE}\n"
        "- Return plain text only\n\n"
        f"Sender: {sender}\n"
        f"Subject: {subject}\n"
        f"Conversation context (older messages, if any):\n{thread_context or '(none)'}\n\n"
        f"Email body:\n{body_text}\n"
    )
    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": OPENAI_MODEL,
        "input": [
            {
                "role": "system",
                "content": (
                    "You draft clear, professional business email replies in an executive CEO voice."
                ),
            },
            {
                "role": "user",
                "content": prompt,
            },
        ],
    }
    try:
        resp = session.post(
            "https://api.openai.com/v1/responses",
            headers=headers,
            json=payload,
            timeout=45,
        )
        resp.raise_for_status()
        data = resp.json()
        text = (data.get("output_text") or "").strip()
        if not text:
            for item in data.get("output", []):
                for content in item.get("content", []):
                    if content.get("type") in ("output_text", "text"):
                        text = (content.get("text") or "").strip()
                        if text:
                            break
                if text:
                    break
        if not text:
            log("LLM fallback: model returned empty text.")
        return sanitize_reply_text(text) or FALLBACK_REPLY_TEXT
    except Exception as e:
        log(f"LLM fallback: reply generation failed: {e}")
        return FALLBACK_REPLY_TEXT


def generate_email_summary_with_llm(session, sender, subject, body_text, thread_context=""):
    fallback = (body_text or "").strip()[:300] or "No summary available."
    if not OPENAI_API_KEY:
        return fallback

    prompt = (
        "Summarize this email for a Teams channel notification.\n"
        "Requirements:\n"
        "- Keep it concise (30-80 words)\n"
        "- Highlight intent, key request, and urgency if any\n"
        "- Plain text only\n\n"
        f"Sender: {sender}\n"
        f"Subject: {subject}\n"
        f"Conversation context (older messages, if any):\n{thread_context or '(none)'}\n\n"
        f"Email body:\n{body_text}\n"
    )
    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": OPENAI_MODEL,
        "input": [
            {
                "role": "system",
                "content": "You write concise, accurate summaries for business emails.",
            },
            {
                "role": "user",
                "content": prompt,
            },
        ],
    }
    try:
        resp = session.post(
            "https://api.openai.com/v1/responses",
            headers=headers,
            json=payload,
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        text = (data.get("output_text") or "").strip()
        if not text:
            for item in data.get("output", []):
                for content in item.get("content", []):
                    if content.get("type") in ("output_text", "text"):
                        text = (content.get("text") or "").strip()
                        if text:
                            break
                if text:
                    break
        return text or fallback
    except Exception:
        return fallback


def build_reply_body_html(reply_text, existing_body):
    """Return an HTML body with reply_text on top of the draft's existing
    content (the quoted original email that createReplyAll puts in)."""
    reply_lines = html.escape((reply_text or "").replace("\r\n", "\n")).replace("\n", "<br>")
    reply_html = (
        '<div style="font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt;">'
        f"{reply_lines}</div><br>"
    )

    content = existing_body.get("content") or ""
    if (existing_body.get("contentType") or "").lower() != "html":
        content = html.escape(content.replace("\r\n", "\n")).replace("\n", "<br>")

    body_tag = re.search(r"<body[^>]*>", content, flags=re.IGNORECASE)
    if body_tag:
        return content[:body_tag.end()] + reply_html + content[body_tag.end():]
    return reply_html + content


def create_reply_draft(session, headers, original_message, reply_text):
    message_id = original_message.get("id")
    if not message_id:
        return None

    # Use Graph's native reply-all logic to preserve correct To/Cc recipients.
    # Ask for an HTML body so the draft comes back with the quoted original
    # email ("From: ... Sent: ...") already in it, as Outlook's Reply All does.
    html_headers = {**headers, "Prefer": 'outlook.body-content-type="html"'}
    create_url = f"https://graph.microsoft.com/v1.0/users/{MAILBOX_USER_ENCODED}/messages/{message_id}/createReplyAll"
    create_resp = session.post(create_url, headers=html_headers, timeout=30)
    raise_for_status_with_details(create_resp, "Create reply draft")
    draft = create_resp.json()
    draft_id = draft.get("id")
    if not draft_id:
        return None

    patch_url = f"https://graph.microsoft.com/v1.0/users/{MAILBOX_USER_ENCODED}/messages/{draft_id}"
    existing_body = draft.get("body") or {}
    if not existing_body.get("content"):
        get_resp = session.get(patch_url, headers=html_headers, params={"$select": "body"}, timeout=30)
        raise_for_status_with_details(get_resp, "Read reply draft body")
        existing_body = get_resp.json().get("body") or {}

    # Put the AI reply above the quoted original instead of replacing the
    # whole body, so the draft (and the sent reply) keeps the conversation.
    patch_body = {
        "body": {
            "contentType": "HTML",
            "content": build_reply_body_html(reply_text, existing_body),
        },
    }
    patch_headers = {**headers, "Content-Type": "application/json"}
    patch_resp = session.patch(patch_url, headers=patch_headers, json=patch_body, timeout=30)
    raise_for_status_with_details(patch_resp, "Update reply draft body")
    return draft_id


def get_to_addresses(message):
    recipients = message.get("toRecipients") or []
    addresses = []
    for recipient in recipients:
        address = (recipient.get("emailAddress") or {}).get("address")
        if address:
            addresses.append(address.strip().lower())
    return addresses


def is_automated_sender(sender):
    sender = (sender or "").lower()
    patterns = [
        "noreply",
        "no-reply",
        "donotreply",
        "notification",
        "alerts",
        "mailer-daemon",
    ]
    return any(p in sender for p in patterns)


def has_direct_greeting_for_name(text, name):
    if not text or not name:
        return False
    snippet = text[:600]
    escaped_name = re.escape(name.strip())
    patterns = [
        rf"\b(?:dear|hello|hi)\s+(?:mr\.?\s+)?{escaped_name}\b",
        rf"\bgood\s+(?:morning|afternoon|evening)\s*,?\s*(?:mr\.?\s+)?{escaped_name}\b",
    ]
    return any(re.search(pattern, snippet, flags=re.IGNORECASE) for pattern in patterns)


def is_vip_sender(sender):
    sender = (sender or "").strip().lower()
    if not sender:
        return False
    if sender in VIP_SENDERS:
        return True
    domain = sender.rsplit("@", 1)[-1] if "@" in sender else ""
    return domain in VIP_DOMAINS


def is_addressed_personally_llm(session, body_text, name):
    if not OPENAI_API_KEY:
        return False

    prompt = (
        f"Email (first 600 chars):\n{(body_text or '')[:600]}\n\n"
        f"Is this email personally addressed to {name} specifically — written as if "
        "speaking directly to them, not a mass email, newsletter, CC broadcast, or "
        "automated notice? Answer with exactly one word: yes or no."
    )
    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": OPENAI_MODEL,
        "input": [{"role": "user", "content": prompt}],
    }
    try:
        resp = session.post(
            "https://api.openai.com/v1/responses",
            headers=headers,
            json=payload,
            timeout=15,
        )
        resp.raise_for_status()
        data = resp.json()
        text = (data.get("output_text") or "").strip()
        if not text:
            for item in data.get("output", []):
                for content in item.get("content", []):
                    if content.get("type") in ("output_text", "text"):
                        text = (content.get("text") or "").strip()
                        if text:
                            break
                if text:
                    break
        return text.strip().lower().startswith("y")
    except Exception as e:
        log(f"Personal-address LLM check failed, falling back to regex-only: {e}")
        return False


def build_action_link(draft_id, action):
    if not N8N_ACTION_WEBHOOK_URL or not draft_id:
        return None
    expiry = int(time.time()) + LINK_EXPIRY_DAYS * 86400
    params = {
        "action": action,
        "mailbox": MAILBOX_USER,
        "draft_id": draft_id,
        "expiry": str(expiry),
    }
    return f"{N8N_ACTION_WEBHOOK_URL}?{urlencode(params)}"


def send_teams_channel_notification(session, headers, sender, subject, received, ai_summary, outlook_link="", draft_id=None):
    if not TEAMS_CHANNEL_EMAIL:
        return False

    safe_sender = html.escape(sender or "unknown")
    safe_subject = html.escape(subject or "(no subject)")
    safe_received = html.escape(format_received_abudhabi(received))
    safe_summary = html.escape((ai_summary or "No summary available.").strip()).replace("\n", "<br>")
    open_in_outlook_html = (
        f'<p><a href="{html.escape(outlook_link, quote=True)}">Open email</a></p>'
        if outlook_link
        else ""
    )
    send_link = build_action_link(draft_id, "send")
    discard_link = build_action_link(draft_id, "discard")
    action_buttons_html = ""
    if send_link and discard_link:
        button_style = (
            "display:inline-block;padding:8px 16px;margin-right:8px;"
            "border-radius:4px;text-decoration:none;font-weight:bold;"
        )
        action_buttons_html = (
            "<p>"
            f'<a href="{html.escape(send_link, quote=True)}" '
            f'style="{button_style}background:#107C10;color:#ffffff;">Send Reply</a>'
            f'<a href="{html.escape(discard_link, quote=True)}" '
            f'style="{button_style}background:#A80000;color:#ffffff;">Discard Draft</a>'
            "</p>"
        )
    url = f"https://graph.microsoft.com/v1.0/users/{MAILBOX_USER_ENCODED}/sendMail"
    payload = {
        "message": {
            "subject": f"New direct email: {subject or '(no subject)'}",
            "body": {
                "contentType": "HTML",
                "content": (
                    "<p><strong>New direct email received</strong></p>"
                    f"<p>From: {safe_sender}<br>"
                    f"Subject: {safe_subject}<br>"
                    f"Received: {safe_received}</p>"
                    f"{open_in_outlook_html}"
                    f"<p><strong>AI Summary:</strong><br>{safe_summary}</p>"
                    f"{action_buttons_html}"
                ),
            },
            "toRecipients": [{"emailAddress": {"address": TEAMS_CHANNEL_EMAIL}}],
        },
        "saveToSentItems": False,
    }
    post_headers = {**headers, "Content-Type": "application/json"}
    try:
        resp = session.post(url, headers=post_headers, json=payload, timeout=15)
        raise_for_status_with_details(resp, "Teams channel notification")
        return True
    except Exception as e:
        if DEV_MODE:
            log("Teams channel notification failed:", e)
        else:
            log("Teams message failed")
        return False


def main():
    auth = GraphAuth()
    warn_if_missing_mail_role(auth.get_valid_token())
    headers = auth.get_headers()
    session = build_http_session()
    log(f"Email bot started for {MAILBOX_USER}")
    log(f"Watching inbox every {POLL_SECONDS}s. Greeting name: {ADDRESSING_NAME}")

    url = (
        f"https://graph.microsoft.com/v1.0/users/{MAILBOX_USER_ENCODED}/mailFolders/Inbox/messages"
        "?$top=50&$orderby=receivedDateTime desc"
        "&$select=id,subject,from,toRecipients,ccRecipients,conversationId,receivedDateTime,bodyPreview,webLink"
    )

    seen_ids = load_seen_ids(SEEN_IDS_FILE)
    # If we loaded persisted state, we already know what's been evaluated
    # across restarts, so skip re-priming: any unseen item in the current
    # window gets evaluated immediately, which is what recovers messages
    # that arrived while the process was down.
    primed = bool(seen_ids)
    if primed:
        log(f"Resumed with {len(seen_ids)} previously-seen message ids from {SEEN_IDS_FILE}")

    try:
        while True:
            items, headers = poll_inbox(
                session=session,
                auth=auth,
                url=url,
                timeout=30,
            )
            if items is None:
                time.sleep(POLL_SECONDS)
                continue

            if not primed:
                # First-ever run (no persisted state): prime the seen set so
                # existing inbox items are not processed as new.
                for m in items:
                    mid = m.get("id")
                    if mid:
                        seen_ids[mid] = time.time()
                primed = True
                save_seen_ids(SEEN_IDS_FILE, seen_ids)
                time.sleep(POLL_SECONDS)
                continue

            # Show oldest-to-newest among unseen items for readable output order.
            new_items = [m for m in reversed(items) if m.get("id") and m["id"] not in seen_ids]
            for m in new_items:
                mid = m["id"]
                seen_ids[mid] = time.time()
                sender = (m.get("from") or {}).get("emailAddress", {}).get("address", "unknown")
                if is_automated_sender(sender):
                    log(f"Skipped automated email from {sender}")
                    continue
                to_addresses = get_to_addresses(m)
                if DIRECT_TO_ADDRESS not in to_addresses:
                    if DEV_MODE:
                        log(f"Skipped email not directly addressed to {DIRECT_TO_ADDRESS}")
                    continue

                try:
                    full_body = get_full_message_body(session, headers, mid)
                except (requests.RequestException, RuntimeError) as e:
                    log(f"Failed to fetch message body for {mid}: {e}")
                    continue

                if is_vip_sender(sender):
                    pass  # VIP override: always process regardless of greeting shape
                elif has_direct_greeting_for_name(full_body, ADDRESSING_NAME):
                    pass  # fast path, no LLM call needed
                elif is_addressed_personally_llm(session, full_body, ADDRESSING_NAME):
                    log(f"Personal-address check: LLM confirmed direct address to {ADDRESSING_NAME} (regex missed it)")
                else:
                    log(f"Skipped email without greeting for {ADDRESSING_NAME}")
                    continue

                subject = m.get("subject") or ""
                thread_context = build_thread_context(
                    session=session,
                    headers=headers,
                    conversation_id=m.get("conversationId"),
                    current_message_id=mid,
                )

                summary_text = generate_email_summary_with_llm(
                    session=session,
                    sender=sender,
                    subject=subject,
                    body_text=full_body,
                    thread_context=thread_context,
                )
                reply_text = generate_reply_with_llm(
                    session=session,
                    sender=sender,
                    subject=subject,
                    body_text=full_body,
                    thread_context=thread_context,
                )

                # Draft creation runs before the Teams notification so a
                # successful draft_id can be included as Send/Discard action
                # links. It's wrapped so a failure here still lets the Teams
                # alert fire below (draft_id=None -> notification just omits
                # the buttons instead of skipping the alert).
                try:
                    draft_id = create_reply_draft(session, headers, m, reply_text)
                    if draft_id:
                        log("Draft created")
                    else:
                        log("Draft skipped")
                except (requests.RequestException, RuntimeError) as e:
                    log(f"Draft creation failed for {mid}: {e}")
                    draft_id = None

                teams_sent = send_teams_channel_notification(
                    session=session,
                    headers=headers,
                    sender=sender,
                    subject=subject,
                    received=m.get("receivedDateTime") or "",
                    ai_summary=summary_text,
                    outlook_link=m.get("webLink") or "",
                    draft_id=draft_id,
                )
                if teams_sent:
                    log("Teams message sent")

            if new_items:
                save_seen_ids(SEEN_IDS_FILE, seen_ids)

            time.sleep(POLL_SECONDS)
    except KeyboardInterrupt:
        log("\nStopped inbox watcher.")

if __name__ == "__main__":
    main()
