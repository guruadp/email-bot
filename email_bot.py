import os
import json
import time
import re
import html
import base64
from datetime import datetime, timezone
import requests
import msal
from dotenv import load_dotenv
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from urllib.parse import quote
from zoneinfo import ZoneInfo

load_dotenv()

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


TENANT_ID = get_required_env("TENANT_ID")  # Directory (tenant) ID
CLIENT_ID = get_required_env("CLIENT_ID")  # Application (client) ID
CLIENT_SECRET = get_required_env("CLIENT_SECRET")  # Application client secret
MAILBOX_USER = get_required_env("MAILBOX_USER")  # User principal name or user id
DIRECT_TO_ADDRESS = MAILBOX_USER.strip().lower()

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]
MAILBOX_USER_ENCODED = quote(MAILBOX_USER, safe="")

POLL_SECONDS = 15
MAX_THREAD_MESSAGES = int(os.getenv("MAX_THREAD_MESSAGES", "5"))
MAX_THREAD_CHARS_PER_MESSAGE = int(os.getenv("MAX_THREAD_CHARS_PER_MESSAGE", "1200"))
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
TEAMS_CHANNEL_EMAIL = get_required_env("TEAMS_CHANNEL_EMAIL").strip().lower()
DEV_MODE = get_bool_env("DEV_MODE", default=False)
ADDRESSING_NAME = os.getenv("ADDRESSING_NAME", "Guru").strip()
EMAIL_SIGNATURE = get_required_env("EMAIL_SIGNATURE").replace("\\n", "\n").strip()
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


def print_token_diagnostics(token):
    payload = decode_jwt_payload(token)
    roles = payload.get("roles") or []
    scp = payload.get("scp")

    print("Token diagnostics:")
    print("  roles:", roles if roles else "(none)")
    print("  scp:", scp if scp else "(none)")

    if "Mail.ReadWrite" not in roles:
        print("Warning: token is missing app role 'Mail.ReadWrite'.")
        print("Check Graph API permissions (Application) and admin consent.")


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
            print(f"Thread context skipped: {e}")
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


def generate_reply_with_llm(sender, subject, body_text, thread_context=""):
    if not OPENAI_API_KEY:
        print("LLM fallback: OPENAI_API_KEY is not set.")
        return FALLBACK_REPLY_TEXT

    prompt = (
        "Write a professional email reply draft.\n"
        "Role and voice:\n"
        "- You are Guru Nandhan, CEO of Ednex LLC, Ednex Automation, and Maker and Coder\n"
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
    resp = requests.post("https://api.openai.com/v1/responses", headers=headers, json=payload, timeout=45)
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
        print("LLM fallback: model returned empty text.")
    return sanitize_reply_text(text) or FALLBACK_REPLY_TEXT


def generate_email_summary_with_llm(sender, subject, body_text, thread_context=""):
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
        resp = requests.post(
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


def create_reply_draft(session, headers, original_message, reply_text):
    message_id = original_message.get("id")
    if not message_id:
        return None

    create_url = f"https://graph.microsoft.com/v1.0/users/{MAILBOX_USER_ENCODED}/messages/{message_id}/createReply"
    create_resp = session.post(create_url, headers=headers, timeout=30)
    raise_for_status_with_details(create_resp, "Create reply draft")
    draft = create_resp.json()
    draft_id = draft.get("id")
    if not draft_id:
        return None

    patch_url = f"https://graph.microsoft.com/v1.0/users/{MAILBOX_USER_ENCODED}/messages/{draft_id}"
    to_emails = []
    sender = (original_message.get("from") or {}).get("emailAddress", {}).get("address")
    if sender:
        to_emails.append(sender.strip().lower())

    for recipient in original_message.get("toRecipients") or []:
        address = (recipient.get("emailAddress") or {}).get("address")
        if address:
            to_emails.append(address.strip().lower())

    # Keep order but remove duplicates.
    to_unique = list(dict.fromkeys(to_emails))

    cc_emails = []
    for recipient in original_message.get("ccRecipients") or []:
        address = (recipient.get("emailAddress") or {}).get("address")
        if address:
            cc_emails.append(address.strip().lower())
    cc_unique = list(dict.fromkeys(cc_emails))

    patch_body = {
        "body": {
            "contentType": "Text",
            "content": reply_text,
        },
        "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_unique],
        "ccRecipients": [{"emailAddress": {"address": addr}} for addr in cc_unique],
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


def send_teams_channel_notification(session, headers, sender, subject, received, ai_summary, outlook_link=""):
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
            print("Teams channel notification failed:", e)
        else:
            print("Teams message failed")
        return False


def main():
    token = get_token()
    print_token_diagnostics(token)
    headers = {"Authorization": f"Bearer {token}"}
    session = build_http_session()
    print("LLM mode:", "enabled" if OPENAI_API_KEY else "disabled")
    print("Mailbox mode: app-only auth for", MAILBOX_USER)
    print("Teams channel notifications:", "enabled" if TEAMS_CHANNEL_EMAIL else "disabled")
    print("Dev mode:", "enabled" if DEV_MODE else "disabled")

    url = (
        f"https://graph.microsoft.com/v1.0/users/{MAILBOX_USER_ENCODED}/mailFolders/Inbox/messages"
        "?$top=50&$orderby=receivedDateTime desc"
        "&$select=id,subject,from,toRecipients,ccRecipients,conversationId,receivedDateTime,bodyPreview,webLink"
    )

    seen_ids = set()

    try:
        # Prime the seen set so existing inbox items are not printed as "new".
        first = session.get(url, headers=headers, timeout=30)
        raise_for_status_with_details(first, "Initial inbox poll")
        for m in first.json().get("value", []):
            mid = m.get("id")
            if mid:
                seen_ids.add(mid)

        print(f"Watching inbox for new mail (poll every {POLL_SECONDS}s). Press Ctrl+C to stop.")

        while True:
            r = session.get(url, headers=headers, timeout=30)
            raise_for_status_with_details(r, "Inbox poll")
            items = r.json().get("value", [])

            # Show oldest-to-newest among unseen items for readable output order.
            new_items = [m for m in reversed(items) if m.get("id") and m["id"] not in seen_ids]
            for m in new_items:
                mid = m["id"]
                seen_ids.add(mid)
                sender = (m.get("from") or {}).get("emailAddress", {}).get("address", "unknown")
                if DEV_MODE:
                    print("\n--- New Mail ---")
                    print("From:", sender)
                    print("Subject:", m.get("subject"))
                    print("Received:", m.get("receivedDateTime"))
                    print("Preview:", (m.get("bodyPreview") or "")[:140])
                else:
                    print("Mail received")
                if is_automated_sender(sender):
                    if DEV_MODE:
                        print("Skipped: automated email")
                    continue
                to_addresses = get_to_addresses(m)
                if DIRECT_TO_ADDRESS not in to_addresses:
                    if DEV_MODE:
                        print(
                            f"Skipped: not directly addressed to {DIRECT_TO_ADDRESS}. "
                            f"To={to_addresses or ['(none)']}"
                        )
                    continue
                try:
                    full_body = get_full_message_body(session, headers, mid)
                    if not has_direct_greeting_for_name(full_body, ADDRESSING_NAME):
                        if DEV_MODE:
                            print(
                                f"Skipped: no direct greeting found for '{ADDRESSING_NAME}'."
                            )
                        continue
                    thread_context = build_thread_context(
                        session=session,
                        headers=headers,
                        conversation_id=m.get("conversationId"),
                        current_message_id=mid,
                    )
                    reply_text = generate_reply_with_llm(
                        sender=sender,
                        subject=m.get("subject") or "",
                        body_text=full_body,
                        thread_context=thread_context,
                    )
                    summary_text = generate_email_summary_with_llm(
                        sender=sender,
                        subject=m.get("subject") or "",
                        body_text=full_body,
                        thread_context=thread_context,
                    )
                    teams_sent = send_teams_channel_notification(
                        session=session,
                        headers=headers,
                        sender=sender,
                        subject=m.get("subject") or "",
                        received=m.get("receivedDateTime") or "",
                        ai_summary=summary_text,
                        outlook_link=m.get("webLink") or "",
                    )
                    if teams_sent:
                        print("Teams message sent")
                    draft_id = create_reply_draft(session, headers, m, reply_text)
                    if draft_id:
                        if DEV_MODE:
                            print("Draft reply created:", draft_id)
                        else:
                            print("Draft created")
                    else:
                        if DEV_MODE:
                            print("Draft reply skipped: missing draft id.")
                        else:
                            print("Draft skipped")
                except (requests.RequestException, RuntimeError) as e:
                    print("Draft reply failed:", e)

            time.sleep(POLL_SECONDS)
    except KeyboardInterrupt:
        print("\nStopped inbox watcher.")

if __name__ == "__main__":
    main()
