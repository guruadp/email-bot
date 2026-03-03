import os
import json
import time
import re
import html
import requests
import msal
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

def get_required_env(name):
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value

TENANT_ID = get_required_env("TENANT_ID")  # Directory (tenant) ID
CLIENT_ID = get_required_env("CLIENT_ID")  # Application (client) ID

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Mail.Read", "Mail.ReadWrite"]

TOKEN_CACHE_FILE = "token_cache.bin"
POLL_SECONDS = 15
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
FALLBACK_REPLY_TEXT = (
    "Hi,\n\n"
    "Thanks for your email. I have received it and will get back to you shortly.\n\n"
    "Best regards,"
)


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

def load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        cache.deserialize(open(TOKEN_CACHE_FILE, "r").read())
    return cache

def save_cache(cache):
    if cache.has_state_changed:
        open(TOKEN_CACHE_FILE, "w").write(cache.serialize())

def get_token():
    cache = load_cache()
    session = build_http_session()
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
        http_client=session,
    )
    accounts = app.get_accounts()
    result = None

    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Device flow init failed: {json.dumps(flow, indent=2)}")

        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    save_cache(cache)

    if "access_token" not in result:
        raise RuntimeError(f"Token failed: {json.dumps(result, indent=2)}")

    return result["access_token"]


def strip_html(text):
    if not text:
        return ""
    no_tags = re.sub(r"<[^>]+>", " ", text)
    plain = html.unescape(no_tags)
    return re.sub(r"\s+", " ", plain).strip()


def get_full_message_body(session, headers, message_id):
    url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}?$select=body,bodyPreview"
    resp = session.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
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
        "- End with a polite sign-off\n"
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
    return text or FALLBACK_REPLY_TEXT


def create_reply_draft(session, headers, original_message, reply_text):
    message_id = original_message.get("id")
    if not message_id:
        return None

    create_url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/createReply"
    create_resp = session.post(create_url, headers=headers, timeout=30)
    create_resp.raise_for_status()
    draft = create_resp.json()
    draft_id = draft.get("id")
    if not draft_id:
        return None

    patch_url = f"https://graph.microsoft.com/v1.0/me/messages/{draft_id}"
    patch_body = {
        "body": {
            "contentType": "Text",
            "content": reply_text,
        }
    }
    patch_headers = {**headers, "Content-Type": "application/json"}
    patch_resp = session.patch(patch_url, headers=patch_headers, json=patch_body, timeout=30)
    patch_resp.raise_for_status()
    return draft_id

def main():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    session = build_http_session()
    print("LLM mode:", "enabled" if OPENAI_API_KEY else "disabled")

    url = (
        "https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages"
        "?$top=50&$orderby=receivedDateTime desc"
        "&$select=id,subject,from,receivedDateTime,bodyPreview"
    )

    seen_ids = set()

    try:
        # Prime the seen set so existing inbox items are not printed as "new".
        first = session.get(url, headers=headers, timeout=30)
        first.raise_for_status()
        for m in first.json().get("value", []):
            mid = m.get("id")
            if mid:
                seen_ids.add(mid)

        print(f"Watching inbox for new mail (poll every {POLL_SECONDS}s). Press Ctrl+C to stop.")

        while True:
            r = session.get(url, headers=headers, timeout=30)
            r.raise_for_status()
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
