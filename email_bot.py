import os
import json
import time
import re
import html
import base64
import requests
import msal
from dotenv import load_dotenv
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from urllib.parse import quote

load_dotenv()

def get_required_env(name):
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value

TENANT_ID = get_required_env("TENANT_ID")  # Directory (tenant) ID
CLIENT_ID = get_required_env("CLIENT_ID")  # Application (client) ID
CLIENT_SECRET = get_required_env("CLIENT_SECRET")  # Application client secret
MAILBOX_USER = get_required_env("MAILBOX_USER")  # User principal name or user id
DIRECT_TO_ADDRESS = MAILBOX_USER.strip().lower()

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]
MAILBOX_USER_ENCODED = quote(MAILBOX_USER, safe="")

POLL_SECONDS = 15
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
FALLBACK_REPLY_TEXT = (
    "Hi,\n\n"
    "Thanks for your email. I have received it and will get back to you shortly.\n\n"
    "Best regards,\n"
    "Guru\n"
    "Robotics system Integration Engineer"
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


def sanitize_reply_text(text):
    if not text:
        return ""
    lines = text.splitlines()
    while lines and re.match(r"^\s*(subject|from|to|cc)\s*:", lines[0], flags=re.IGNORECASE):
        lines.pop(0)
    return "\n".join(lines).strip()


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


def generate_reply_with_llm(sender, subject, body_text):
    if not OPENAI_API_KEY:
        print("LLM fallback: OPENAI_API_KEY is not set.")
        return FALLBACK_REPLY_TEXT

    prompt = (
        "Write a professional email reply draft.\n"
        "Requirements:\n"
        "- Keep it concise (80-160 words)\n"
        "- Acknowledge the sender's request\n"
        "- Ask one clarifying question if needed\n"
        "- Do not include any email headers (no Subject/From/To/Cc lines)\n"
        "- End with exactly this signature block:\n"
        "Best regards,\n"
        "Guru\n"
        "Robotics system Integration Engineer\n"
        "- Return plain text only\n\n"
        f"Sender: {sender}\n"
        f"Subject: {subject}\n"
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
                "content": "You draft clear, professional email replies.",
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
    patch_body = {
        "body": {
            "contentType": "Text",
            "content": reply_text,
        }
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

def main():
    token = get_token()
    print_token_diagnostics(token)
    headers = {"Authorization": f"Bearer {token}"}
    session = build_http_session()
    print("LLM mode:", "enabled" if OPENAI_API_KEY else "disabled")
    print("Mailbox mode: app-only auth for", MAILBOX_USER)

    url = (
        f"https://graph.microsoft.com/v1.0/users/{MAILBOX_USER_ENCODED}/mailFolders/Inbox/messages"
        "?$top=50&$orderby=receivedDateTime desc"
        "&$select=id,subject,from,toRecipients,receivedDateTime,bodyPreview"
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
                print("\n--- New Mail ---")
                print("From:", sender)
                print("Subject:", m.get("subject"))
                print("Received:", m.get("receivedDateTime"))
                print("Preview:", (m.get("bodyPreview") or "")[:140])
                to_addresses = get_to_addresses(m)
                if DIRECT_TO_ADDRESS not in to_addresses:
                    print(
                        f"Draft reply skipped: not directly addressed to {DIRECT_TO_ADDRESS}. "
                        f"To={to_addresses or ['(none)']}"
                    )
                    continue
                try:
                    full_body = get_full_message_body(session, headers, mid)
                    reply_text = generate_reply_with_llm(
                        sender=sender,
                        subject=m.get("subject") or "",
                        body_text=full_body,
                    )
                    draft_id = create_reply_draft(session, headers, m, reply_text)
                    if draft_id:
                        print("Draft reply created:", draft_id)
                    else:
                        print("Draft reply skipped: missing draft id.")
                except requests.RequestException as e:
                    print("Draft reply failed:", e)

            time.sleep(POLL_SECONDS)
    except KeyboardInterrupt:
        print("\nStopped inbox watcher.")

if __name__ == "__main__":
    main()
