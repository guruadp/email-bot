import os
import io
import sys
import json
import time
import re
import html
import base64
import tempfile
import zipfile
from datetime import datetime, timedelta, timezone
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
LOCAL_TZ = ZoneInfo("Asia/Dubai")
MAILBOX_DOMAIN = DIRECT_TO_ADDRESS.rsplit("@", 1)[-1]


def get_list_env(name, default=""):
    return {v.strip().lower() for v in (os.getenv(name) or default).split(",") if v.strip()}


# Skip marketing, newsletters, cold outreach and system notifications.
MARKETING_FILTER = get_bool_env("MARKETING_FILTER", default=True)
SKIP_SENDERS = get_list_env("SKIP_SENDERS")
SKIP_DOMAINS = get_list_env("SKIP_DOMAINS")

# Learn the mailbox owner's writing style from Sent Items.
STYLE_LEARNING = get_bool_env("STYLE_LEARNING", default=True)
STYLE_SAMPLE_SIZE = int(os.getenv("STYLE_SAMPLE_SIZE", "40"))
STYLE_REFRESH_HOURS = int(os.getenv("STYLE_REFRESH_HOURS", "24"))

# Read PDF, Word, Excel and text attachments.
ATTACHMENTS_ENABLED = get_bool_env("ATTACHMENTS_ENABLED", default=True)
ATTACHMENT_MAX_MB = float(os.getenv("ATTACHMENT_MAX_MB", "10"))
ATTACHMENT_MAX_CHARS = int(os.getenv("ATTACHMENT_MAX_CHARS", "6000"))
ATTACHMENT_MAX_FILES = 5

# Daily Teams digest of sent emails still waiting for a reply.
FOLLOWUP_ENABLED = get_bool_env("FOLLOWUP_ENABLED", default=True)
FOLLOWUP_AFTER_DAYS = int(os.getenv("FOLLOWUP_AFTER_DAYS", "3"))
FOLLOWUP_LOOKBACK_DAYS = int(os.getenv("FOLLOWUP_LOOKBACK_DAYS", "14"))
FOLLOWUP_DIGEST_HOUR = int(os.getenv("FOLLOWUP_DIGEST_HOUR", "9"))  # Dubai time
FOLLOWUP_IGNORE_DOMAINS = get_list_env("FOLLOWUP_IGNORE_DOMAINS", MAILBOX_DOMAIN)
FOLLOWUP_STATE_FILE = os.getenv("FOLLOWUP_STATE_FILE", "followup_state.json")
FOLLOWUP_MAX_ITEMS = 25

GRAPH_USER_URL = f"https://graph.microsoft.com/v1.0/users/{MAILBOX_USER_ENCODED}"
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


def get_message_details(session, headers, message_id):
    """Body text (first 4000 chars), internet headers and attachment flag for one message."""
    url = f"{GRAPH_USER_URL}/messages/{message_id}"
    params = {"$select": "body,bodyPreview,internetMessageHeaders,hasAttachments"}
    resp = session.get(url, headers=headers, params=params, timeout=30)
    raise_for_status_with_details(resp, "Fetch message body")
    data = resp.json()
    body = (data.get("body") or {}).get("content", "")
    body_type = ((data.get("body") or {}).get("contentType") or "").lower()
    if body_type == "html":
        body = strip_html(body)
    if not body:
        body = data.get("bodyPreview") or ""
    return {
        "body": body[:4000],
        "internet_headers": data.get("internetMessageHeaders") or [],
        "has_attachments": bool(data.get("hasAttachments")),
    }


def llm_text(session, messages, timeout=30):
    """Call the OpenAI Responses API and return its text output ("" if empty).
    Raises on HTTP or network errors; callers decide the fallback."""
    resp = session.post(
        "https://api.openai.com/v1/responses",
        headers={"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type": "application/json"},
        json={"model": OPENAI_MODEL, "input": messages},
        timeout=timeout,
    )
    resp.raise_for_status()
    data = resp.json()
    text = (data.get("output_text") or "").strip()
    if text:
        return text
    for item in data.get("output", []):
        for content in item.get("content", []):
            if content.get("type") in ("output_text", "text") and (content.get("text") or "").strip():
                return content["text"].strip()
    return ""


# ---------------------------------------------------------------------------
# Attachments
# ---------------------------------------------------------------------------

def extract_attachment_text(name, data):
    """Best-effort plain text from a PDF, Word (.docx), Excel (.xlsx) or text file.
    Returns "" for unsupported or unreadable files."""
    ext = os.path.splitext((name or "").lower())[1]
    try:
        if ext == ".pdf":
            from pypdf import PdfReader  # imported lazily so a missing package only disables PDFs

            reader = PdfReader(io.BytesIO(data))
            if reader.is_encrypted:
                return ""
            return "\n".join((page.extract_text() or "") for page in reader.pages[:20])
        if ext == ".docx":
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
            xml = re.sub(r"</w:p>|<w:br/>|<w:tab/>", "\n", xml)
            return html.unescape(re.sub(r"<[^>]+>", "", xml))
        if ext == ".xlsx":
            import openpyxl

            wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            lines = []
            for ws in wb.worksheets[:5]:
                lines.append(f"[Sheet: {ws.title}]")
                for row in ws.iter_rows(max_row=200, values_only=True):
                    cells = [str(c) for c in row if c is not None and str(c).strip()]
                    if cells:
                        lines.append(" | ".join(cells))
            return "\n".join(lines)
        if ext in (".txt", ".csv", ".md", ".json"):
            return data.decode("utf-8", errors="ignore")
    except Exception as e:
        log(f"Could not read attachment {name}: {type(e).__name__}")
    return ""


def get_attachments_text(session, headers, message_id):
    """Return (text_for_prompt, attachment_names). Reads at most ATTACHMENT_MAX_FILES
    non-inline files up to ATTACHMENT_MAX_MB each, ATTACHMENT_MAX_CHARS in total."""
    url = f"{GRAPH_USER_URL}/messages/{message_id}/attachments"
    resp = session.get(url, headers=headers, params={"$select": "id,name,size,isInline"}, timeout=30)
    raise_for_status_with_details(resp, "List attachments")
    attachments = [
        a for a in resp.json().get("value", [])
        if not a.get("isInline")
        and not a.get("@odata.type", "").endswith(("itemAttachment", "referenceAttachment"))
    ]
    names = [a.get("name") or "(unnamed)" for a in attachments]
    chunks, remaining = [], ATTACHMENT_MAX_CHARS
    for a in attachments[:ATTACHMENT_MAX_FILES]:
        if remaining <= 0:
            break
        if (a.get("size") or 0) > ATTACHMENT_MAX_MB * 1024 * 1024:
            continue
        name = a.get("name") or ""
        if os.path.splitext(name.lower())[1] not in (".pdf", ".docx", ".xlsx", ".txt", ".csv", ".md", ".json"):
            continue
        raw = session.get(f"{url}/{a['id']}/$value", headers=headers, timeout=60)
        if raw.status_code >= 400:
            log(f"Could not download attachment {name}: HTTP {raw.status_code}")
            continue
        text = re.sub(r"[ \t]+", " ", extract_attachment_text(name, raw.content))
        text = re.sub(r"\n\s*\n+", "\n", text).strip()
        if text:
            chunks.append(f"[Attachment: {name}]\n{text[:remaining]}")
            remaining -= len(text)
    return "\n\n".join(chunks), names


def generate_reply_with_llm(session, sender, subject, body_text, thread_context="", attachments_text="", style_text=""):
    if not OPENAI_API_KEY:
        log("LLM fallback: OPENAI_API_KEY is not set.")
        return FALLBACK_REPLY_TEXT

    if style_text:
        voice = (
            f"- You are {PERSONA_DESCRIPTION}\n"
            "- Write the way this person really writes. Match their style as described below\n"
            "  (greeting, length, tone, phrasing), but never copy facts from the examples:\n"
            f"{style_text}\n"
        )
    else:
        voice = (
            f"- You are {PERSONA_DESCRIPTION}\n"
            "- Write with clear executive tone: confident, concise, strategic, and respectful\n"
            "- Sound like a CEO responding directly to business stakeholders\n"
        )
    prompt = (
        "Write a professional email reply draft.\n"
        "Role and voice:\n"
        f"{voice}"
        "Requirements:\n"
        "- Keep it concise (80-160 words) unless the writing style above is clearly shorter\n"
        "- Use the conversation context for continuity when provided\n"
        "- Acknowledge the sender's request\n"
        "- If attachments are provided, refer to their actual contents where relevant "
        "(e.g. quote amounts, dates or items), and never invent details that are not in them\n"
        "- Do not include any email headers (no Subject/From/To/Cc lines)\n"
        "- End with exactly this signature block:\n"
        f"{EMAIL_SIGNATURE}\n"
        "- Return plain text only\n\n"
        f"Sender: {sender}\n"
        f"Subject: {subject}\n"
        f"Conversation context (older messages, if any):\n{thread_context or '(none)'}\n\n"
        f"Email body:\n{body_text}\n\n"
        f"Attachments (extracted text, if any):\n{attachments_text or '(none)'}\n"
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


def generate_email_summary_with_llm(session, sender, subject, body_text, thread_context="", attachments_text=""):
    fallback = (body_text or "").strip()[:300] or "No summary available."
    if not OPENAI_API_KEY:
        return fallback

    prompt = (
        "Summarize this email for a Teams channel notification.\n"
        "Requirements:\n"
        "- Keep it concise (30-80 words, up to 120 if there are attachments)\n"
        "- Highlight intent, key request, and urgency if any\n"
        "- If attachments are provided, add one line on what they contain "
        "(e.g. 'Attached quote: 3 items, total AED 45,000, valid until 30 Nov')\n"
        "- Plain text only\n\n"
        f"Sender: {sender}\n"
        f"Subject: {subject}\n"
        f"Conversation context (older messages, if any):\n{thread_context or '(none)'}\n\n"
        f"Email body:\n{body_text}\n\n"
        f"Attachments (extracted text, if any):\n{attachments_text or '(none)'}\n"
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


AUTOMATED_SENDER_PATTERNS = [
    "noreply",
    "no-reply",
    "donotreply",
    "do-not-reply",
    "do_not_reply",
    "notification",
    "alerts",
    "mailer",
    "systemgenerated",
    "system-generated",
    "newsletter",
    "marketing",
    "bounce",
    "postmaster",
]


def is_automated_sender(sender):
    sender = (sender or "").lower()
    return any(p in sender for p in AUTOMATED_SENDER_PATTERNS)


def is_skipped_sender(sender):
    """True for addresses or domains listed in SKIP_SENDERS / SKIP_DOMAINS."""
    sender = (sender or "").strip().lower()
    domain = sender.rsplit("@", 1)[-1] if "@" in sender else ""
    return sender in SKIP_SENDERS or domain in SKIP_DOMAINS


def bulk_header_reason(internet_headers):
    """Return why the email looks bulk or machine-sent based on its headers, or None.
    Newsletters and marketing tools add List-Unsubscribe/List-Id; auto-generated
    mail adds Auto-Submitted or Precedence. Feedback-ID is deliberately not used:
    some ordinary mail services add it to person-to-person email too."""
    h = {(x.get("name") or "").lower(): (x.get("value") or "").strip().lower() for x in internet_headers or []}
    if "list-unsubscribe" in h:
        return "List-Unsubscribe header (newsletter/marketing)"
    if "list-id" in h:
        return "List-Id header (mailing list)"
    if h.get("precedence") in ("bulk", "list", "junk"):
        return f"Precedence: {h['precedence']}"
    if h.get("auto-submitted", "no") != "no":
        return f"Auto-Submitted: {h['auto-submitted']}"
    if "x-campaign" in h or "x-campaignid" in h:
        return "campaign header (marketing)"
    return None


def classify_email_llm(session, sender, subject, body_text):
    """Return 'personal', 'marketing' or 'automated'. Fails open to 'personal'
    so an LLM outage never silently drops a real email."""
    if not OPENAI_API_KEY:
        return "personal"
    prompt = (
        f"Classify this email received by {ADDRESSING_NAME} ({PERSONA_DESCRIPTION}).\n"
        f"personal = a real person writing to {ADDRESSING_NAME} about an actual matter: clients, partners, "
        "suppliers handling an existing order or quote, colleagues, government, or a bank writing about a "
        f"specific case, where a reply from {ADDRESSING_NAME} may be needed.\n"
        "marketing = newsletters, promotions, discounts, event or webinar invitations sent to many people, "
        "and unsolicited cold sales outreach (someone pitching their own product or service), even when it "
        f"greets {ADDRESSING_NAME} by name.\n"
        "automated = system notifications, CRM or task alerts, receipts, confirmations, reports, or any "
        "machine-generated message.\n"
        "Answer with exactly one word: personal, marketing or automated.\n\n"
        f"Sender: {sender}\nSubject: {subject}\nEmail (first 1500 chars):\n{(body_text or '')[:1500]}"
    )
    try:
        answer = llm_text(session, [{"role": "user", "content": prompt}], timeout=20).lower()
    except Exception as e:
        log(f"Email classification failed, treating as personal: {e}")
        return "personal"
    for label in ("marketing", "automated", "personal"):
        if label in answer:
            return label
    return "personal"


def mailbox_replied_in_thread(session, headers, conversation_id):
    """True if the mailbox owner has sent a message in this conversation, i.e. it is
    an ongoing real exchange rather than a newsletter or cold pitch."""
    if not conversation_id:
        return False
    params = {
        "$filter": f"conversationId eq '{conversation_id.replace(chr(39), chr(39) * 2)}'",
        "$select": "from,isDraft",
        "$top": "50",
    }
    try:
        resp = session.get(f"{GRAPH_USER_URL}/messages", headers=headers, params=params, timeout=30)
        raise_for_status_with_details(resp, "Fetch thread participants")
    except (requests.RequestException, RuntimeError) as e:
        log(f"Thread participant check skipped: {e}")
        return False
    return any(
        not t.get("isDraft")
        and ((t.get("from") or {}).get("emailAddress") or {}).get("address", "").lower() == DIRECT_TO_ADDRESS
        for t in resp.json().get("value", [])
    )


def evaluate_email(session, headers, m):
    """Decide whether the bot should handle inbox message m.
    Returns (process, reason, details). details is the message body/headers dict
    when it was fetched, else None."""
    sender = (m.get("from") or {}).get("emailAddress", {}).get("address", "unknown")
    sender_lc = sender.strip().lower()
    if sender_lc == DIRECT_TO_ADDRESS:
        return False, "sent by the mailbox itself", None
    if is_skipped_sender(sender):
        return False, "sender is in SKIP_SENDERS/SKIP_DOMAINS", None
    if re.match(r"^(accepted|declined|tentative|canceled|cancelled)\s*:", m.get("subject") or "", re.I):
        return False, "meeting response", None
    vip = is_vip_sender(sender)
    if not vip and is_automated_sender(sender):
        return False, "automated sender address", None
    if DIRECT_TO_ADDRESS not in get_to_addresses(m):
        return False, f"not directly addressed to {DIRECT_TO_ADDRESS}", None

    details = get_message_details(session, headers, m["id"])
    if vip:
        return True, "VIP sender", details
    in_real_thread = MARKETING_FILTER and mailbox_replied_in_thread(session, headers, m.get("conversationId"))
    if MARKETING_FILTER and not in_real_thread:
        reason = bulk_header_reason(details["internet_headers"])
        if reason:
            return False, reason, details

    body = details["body"]
    if not has_direct_greeting_for_name(body, ADDRESSING_NAME):
        if not is_addressed_personally_llm(session, body, ADDRESSING_NAME):
            return False, f"no greeting for {ADDRESSING_NAME}", details
        log(f"Personal-address check: LLM confirmed direct address to {ADDRESSING_NAME} (regex missed it)")

    if in_real_thread:
        return True, f"ongoing thread {ADDRESSING_NAME} has replied in", details
    if MARKETING_FILTER:
        label = classify_email_llm(session, sender, m.get("subject") or "", body)
        if label != "personal":
            return False, f"classified as {label}", details
    return True, "personal email", details


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


# ---------------------------------------------------------------------------
# Writing style learned from Sent Items
# ---------------------------------------------------------------------------

QUOTE_START = re.compile(
    r"^\s*(?:-{3,}\s*Original Message\s*-{3,}|_{10,}|From:\s.+|On .+wrote:)\s*$",
    flags=re.IGNORECASE | re.MULTILINE,
)


def fetch_sent_samples(session, headers, count):
    """Return the text the mailbox owner wrote in their most recent sent emails
    (quoted history removed), skipping Teams notifications and very short ones."""
    params = {
        "$top": str(min(count, 100)),
        "$orderby": "sentDateTime desc",
        "$select": "subject,uniqueBody,toRecipients",
    }
    text_headers = {**headers, "Prefer": 'outlook.body-content-type="text"'}
    resp = session.get(f"{GRAPH_USER_URL}/mailFolders/sentitems/messages", headers=text_headers, params=params, timeout=30)
    raise_for_status_with_details(resp, "Fetch sent items")
    samples = []
    for m in resp.json().get("value", []):
        to = {(r.get("emailAddress") or {}).get("address", "").lower() for r in m.get("toRecipients") or []}
        if TEAMS_CHANNEL_EMAIL in to:
            continue
        text = ((m.get("uniqueBody") or {}).get("content") or "").replace("\r\n", "\n")
        quote = QUOTE_START.search(text)
        if quote:
            text = text[:quote.start()]
        text = re.sub(r"\n{3,}", "\n\n", text).strip()
        if len(text.split()) >= 12:
            samples.append(" ".join(text.split()[:250]) if len(text.split()) > 250 else text)
    return samples


class StyleGuide:
    """Writing-style description of the mailbox owner, rebuilt from Sent Items
    every STYLE_REFRESH_HOURS. Empty text means: use the default persona voice."""

    def __init__(self):
        self.text = ""
        self.refreshed_at = 0.0

    def maybe_refresh(self, session, headers):
        if not (STYLE_LEARNING and OPENAI_API_KEY):
            return
        if self.refreshed_at and time.time() - self.refreshed_at < STYLE_REFRESH_HOURS * 3600:
            return
        self.refreshed_at = time.time()  # also on failure, so errors don't retry every poll
        try:
            samples = fetch_sent_samples(session, headers, STYLE_SAMPLE_SIZE)
            if len(samples) < 5:
                log(f"Style learning skipped: only {len(samples)} usable sent emails.")
                return
            joined = "\n\n=====\n\n".join(samples[:30])
            profile = llm_text(session, [{"role": "user", "content": (
                f"Below are emails written by {PERSONA_DESCRIPTION}, separated by =====.\n"
                "Describe how this person writes, in 6-10 short bullet points: typical greeting, "
                "sign-off, length, tone and formality, sentence structure, recurring phrases, and how "
                "they handle requests or decline things. Describe style only; do not mention any "
                "names, companies, amounts or other facts from the emails.\n\n" + joined
            )}], timeout=60)
            # Two short real examples help the model more than the description alone.
            examples = sorted((s for s in samples if 25 <= len(s.split()) <= 120), key=len)[:2]
            parts = [profile.strip()] if profile.strip() else []
            for i, ex in enumerate(examples, 1):
                parts.append(f"Example {i} of their writing:\n\"\"\"\n{ex}\n\"\"\"")
            self.text = "\n\n".join(parts)
            log(f"Writing style learned from {len(samples)} sent emails.")
        except Exception as e:
            log(f"Style learning failed, using default voice: {e}")


# ---------------------------------------------------------------------------
# Follow-up reminders
# ---------------------------------------------------------------------------

def graph_get_all(session, headers, url, params, limit):
    items = []
    resp = session.get(url, headers=headers, params=params, timeout=30)
    while True:
        raise_for_status_with_details(resp, "Graph list")
        data = resp.json()
        items.extend(data.get("value", []))
        next_link = data.get("@odata.nextLink")
        if not next_link or len(items) >= limit:
            return items[:limit]
        resp = session.get(next_link, headers=headers, timeout=30)


def email_expects_reply_llm(session, subject, text):
    """True if a sent email asks something of the recipient. Fails open (True)."""
    if not OPENAI_API_KEY:
        return True
    try:
        answer = llm_text(session, [{"role": "user", "content": (
            "Here is an email someone sent. Does it ask a question, request something, or otherwise "
            "clearly expect a reply from the recipient? Emails that only say thanks, confirm, inform or "
            "close a topic do not. Answer with exactly one word: yes or no.\n\n"
            f"Subject: {subject}\nEmail:\n{(text or '')[:1500]}"
        )}], timeout=20)
        return not answer.lower().startswith("n")
    except Exception:
        return True


def find_waiting_followups(session, headers, expects_reply_cache=None):
    """Sent emails to outside recipients, FOLLOWUP_AFTER_DAYS..FOLLOWUP_LOOKBACK_DAYS old,
    that expect a reply and have had no reply (and no newer message) in their conversation."""
    cache = expects_reply_cache if expects_reply_cache is not None else {}
    now = datetime.now(timezone.utc)
    start = (now - timedelta(days=FOLLOWUP_LOOKBACK_DAYS)).strftime("%Y-%m-%dT%H:%M:%SZ")
    end = (now - timedelta(days=FOLLOWUP_AFTER_DAYS)).strftime("%Y-%m-%dT%H:%M:%SZ")
    sent = graph_get_all(
        session,
        {**headers, "Prefer": 'outlook.body-content-type="text"'},
        f"{GRAPH_USER_URL}/mailFolders/sentitems/messages",
        {
            "$filter": f"sentDateTime ge {start} and sentDateTime le {end}",
            "$orderby": "sentDateTime desc",
            "$select": "id,subject,toRecipients,ccRecipients,sentDateTime,conversationId,webLink,uniqueBody",
            "$top": "100",
        },
        limit=400,
    )
    latest_per_conversation = {}
    for m in sent:
        if m.get("@odata.type", "#microsoft.graph.message") != "#microsoft.graph.message":
            continue  # meeting responses and other item types
        if re.match(r"^(accepted|declined|tentative|canceled|cancelled)\s*:", m.get("subject") or "", re.I):
            continue
        conv = m.get("conversationId")
        if conv and conv not in latest_per_conversation:
            latest_per_conversation[conv] = m  # sorted newest first

    waiting = []
    for conv, m in latest_per_conversation.items():
        recipients = [
            r.get("emailAddress") or {}
            for r in (m.get("toRecipients") or []) + (m.get("ccRecipients") or [])
        ]
        external = [
            r for r in recipients
            if (r.get("address") or "").lower() not in (DIRECT_TO_ADDRESS, TEAMS_CHANNEL_EMAIL)
            and (r.get("address") or "").lower().rsplit("@", 1)[-1] not in FOLLOWUP_IGNORE_DOMAINS
            and not is_automated_sender(r.get("address"))
        ]
        if not external:
            continue
        sent_at = m["sentDateTime"]
        thread = graph_get_all(
            session, headers, f"{GRAPH_USER_URL}/messages",
            {
                "$filter": f"conversationId eq '{conv.replace(chr(39), chr(39) * 2)}'",
                "$select": "from,receivedDateTime,sentDateTime,isDraft",
                "$top": "50",
            },
            limit=200,
        )
        answered = False
        for t in thread:
            if t.get("isDraft"):
                continue
            from_addr = ((t.get("from") or {}).get("emailAddress") or {}).get("address", "").lower()
            if from_addr == DIRECT_TO_ADDRESS:
                if (t.get("sentDateTime") or "") > sent_at:
                    answered = True  # a newer message was sent; that one is tracked instead
            elif (t.get("receivedDateTime") or "") > sent_at:
                answered = True  # someone replied
            if answered:
                break
        if answered:
            continue
        if m["id"] not in cache:
            text = ((m.get("uniqueBody") or {}).get("content") or "")
            cache[m["id"]] = email_expects_reply_llm(session, m.get("subject") or "", text)
        if not cache[m["id"]]:
            continue
        sent_dt = datetime.fromisoformat(sent_at.replace("Z", "+00:00"))
        waiting.append({
            "to": ", ".join(r.get("name") or r.get("address") for r in external[:3]),
            "subject": m.get("subject") or "(no subject)",
            "sent": sent_dt.astimezone(LOCAL_TZ).strftime("%d %b"),
            "days": (now - sent_dt).days,
            "link": m.get("webLink") or "",
        })
    waiting.sort(key=lambda w: -w["days"])
    return waiting


def build_followup_digest_html(items):
    rows = []
    for w in items[:FOLLOWUP_MAX_ITEMS]:
        subject = html.escape(w["subject"])
        if w["link"]:
            subject = f'<a href="{html.escape(w["link"], quote=True)}">{subject}</a>'
        rows.append(
            f"<li><strong>{html.escape(w['to'])}</strong>: {subject}<br>"
            f"Sent {html.escape(w['sent'])} ({w['days']} days, no reply)</li>"
        )
    more = len(items) - FOLLOWUP_MAX_ITEMS
    extra = f"<p>…and {more} more.</p>" if more > 0 else ""
    return (
        f"<p><strong>{len(items)} sent email(s) still waiting for a reply</strong> "
        f"(sent {FOLLOWUP_AFTER_DAYS}–{FOLLOWUP_LOOKBACK_DAYS} days ago)</p>"
        f"<ul>{''.join(rows)}</ul>{extra}"
    )


def load_followup_state():
    try:
        with open(FOLLOWUP_STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, ValueError, OSError):
        return {}


def save_followup_state(state):
    try:
        with open(FOLLOWUP_STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f)
    except OSError as e:
        log(f"Warning: failed to save follow-up state: {e}")


def maybe_send_followup_digest(session, auth, state, cache):
    """Send the follow-up digest once a day at FOLLOWUP_DIGEST_HOUR (Dubai time)."""
    if not FOLLOWUP_ENABLED:
        return
    local_now = datetime.now(LOCAL_TZ)
    today = local_now.strftime("%Y-%m-%d")
    if local_now.hour != FOLLOWUP_DIGEST_HOUR or state.get("last_digest_date") == today:
        return
    state["last_digest_date"] = today  # set first so a failure doesn't retry every poll
    save_followup_state(state)
    try:
        headers = auth.get_headers()
        items = find_waiting_followups(session, headers, cache)
        if not items:
            log("Follow-up digest: nothing waiting for a reply.")
            return
        sent = send_teams_html(
            session, headers,
            f"Follow-ups: {len(items)} email(s) waiting for a reply",
            build_followup_digest_html(items),
        )
        log(f"Follow-up digest {'sent' if sent else 'failed'}: {len(items)} item(s).")
    except Exception as e:
        log(f"Follow-up digest failed: {e}")


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


def send_teams_html(session, headers, subject, html_content):
    """Email an HTML message to the Teams channel address. Returns True on success."""
    if not TEAMS_CHANNEL_EMAIL:
        return False
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_content},
            "toRecipients": [{"emailAddress": {"address": TEAMS_CHANNEL_EMAIL}}],
        },
        "saveToSentItems": False,
    }
    post_headers = {**headers, "Content-Type": "application/json"}
    try:
        resp = session.post(f"{GRAPH_USER_URL}/sendMail", headers=post_headers, json=payload, timeout=15)
        raise_for_status_with_details(resp, "Teams channel notification")
        return True
    except Exception as e:
        if DEV_MODE:
            log("Teams channel notification failed:", e)
        else:
            log("Teams message failed")
        return False


def send_teams_channel_notification(session, headers, sender, subject, received, ai_summary, outlook_link="", draft_id=None, attachment_names=None):
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
    attachments_html = (
        f"Attachments: {html.escape(', '.join(attachment_names))}<br>" if attachment_names else ""
    )
    return send_teams_html(
        session,
        headers,
        f"New direct email: {subject or '(no subject)'}",
        "<p><strong>New direct email received</strong></p>"
        f"<p>From: {safe_sender}<br>"
        f"Subject: {safe_subject}<br>"
        f"{attachments_html}"
        f"Received: {safe_received}</p>"
        f"{open_in_outlook_html}"
        f"<p><strong>AI Summary:</strong><br>{safe_summary}</p>"
        f"{action_buttons_html}",
    )


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

    style = StyleGuide()
    followup_state = load_followup_state()
    expects_reply_cache = {}
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
            maybe_send_followup_digest(session, auth, followup_state, expects_reply_cache)

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
                try:
                    process, reason, details = evaluate_email(session, headers, m)
                except (requests.RequestException, RuntimeError) as e:
                    log(f"Failed to fetch message body for {mid}: {e}")
                    continue
                if not process:
                    if DEV_MODE or details is not None:
                        log(f"Skipped email from {sender}: {reason}")
                    continue
                full_body = details["body"]

                attachments_text, attachment_names = "", []
                if ATTACHMENTS_ENABLED and details["has_attachments"]:
                    try:
                        attachments_text, attachment_names = get_attachments_text(session, headers, mid)
                        if attachments_text:
                            log(f"Read {len(attachment_names)} attachment(s): {', '.join(attachment_names)}")
                    except (requests.RequestException, RuntimeError) as e:
                        log(f"Attachments skipped for {mid}: {e}")

                style.maybe_refresh(session, headers)

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
                    attachments_text=attachments_text,
                )
                reply_text = generate_reply_with_llm(
                    session=session,
                    sender=sender,
                    subject=subject,
                    body_text=full_body,
                    thread_context=thread_context,
                    attachments_text=attachments_text,
                    style_text=style.text,
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
                    attachment_names=attachment_names,
                )
                if teams_sent:
                    log("Teams message sent")

            if new_items:
                save_seen_ids(SEEN_IDS_FILE, seen_ids)

            time.sleep(POLL_SECONDS)
    except KeyboardInterrupt:
        log("\nStopped inbox watcher.")

PREVIEW_USAGE = """Read-only previews (no drafts, no Teams messages; LLM calls are made):
  python3 email_bot.py --check-filter [N]        draft/skip decision for the last N inbox emails (default 30)
  python3 email_bot.py --preview-style           the writing style learned from Sent Items
  python3 email_bot.py --preview-followups       sent emails that would be in today's follow-up digest
  python3 email_bot.py --preview-attachments [N] text read from attachments in the last N inbox emails (default 30)"""


def run_preview(mode, count):
    auth = GraphAuth()
    session = build_http_session()
    headers = auth.get_headers()
    if mode == "--preview-style":
        style = StyleGuide()
        style.maybe_refresh(session, headers)
        print(style.text or "(no style learned; the default persona voice will be used)")
        return
    if mode == "--preview-followups":
        items = find_waiting_followups(session, headers)
        print(f"{len(items)} sent email(s) waiting for a reply:")
        for w in items:
            print(f"  {w['days']:>2} days | {w['sent']} | to {w['to']} | {w['subject']}")
        return
    inbox = session.get(
        f"{GRAPH_USER_URL}/mailFolders/Inbox/messages",
        headers=headers,
        params={
            "$top": str(count),
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,from,toRecipients,ccRecipients,conversationId,receivedDateTime,hasAttachments",
        },
        timeout=30,
    )
    raise_for_status_with_details(inbox, "Inbox fetch")
    for m in inbox.json().get("value", []):
        sender = (m.get("from") or {}).get("emailAddress", {}).get("address", "unknown")
        when = format_received_abudhabi(m.get("receivedDateTime"))[5:16]
        subject = (m.get("subject") or "")[:50]
        if mode == "--check-filter":
            process, reason, _ = evaluate_email(session, headers, m)
            print(f"{'DRAFT' if process else 'skip '} | {when} | {sender[:35]:35} | {subject:50} | {reason}")
        elif mode == "--preview-attachments" and m.get("hasAttachments"):
            text, names = get_attachments_text(session, headers, m["id"])
            print(f"{when} | {sender[:35]} | {subject}")
            print(f"    files: {names}")
            print(f"    text read: {len(text)} chars | starts: {text[:150].replace(chr(10), ' ')!r}")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        mode = sys.argv[1]
        if mode not in ("--check-filter", "--preview-style", "--preview-followups", "--preview-attachments"):
            sys.exit(PREVIEW_USAGE)
        run_preview(mode, int(sys.argv[2]) if len(sys.argv) > 2 else 30)
    else:
        main()
