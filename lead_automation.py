#!/usr/bin/env python3
"""Read the Sales mailbox via Microsoft Graph, classify lead emails with
OpenAI, extract structured lead details, and append them to an Excel tracker.

Runs one cycle and exits; start.py invokes it on a schedule alongside
email_bot.py. Can also be run standalone: python3 lead_automation.py

Pipeline per email:
  1. Sender filter: whitelisted senders are always processed; unknown senders
     are processed only if classify_unknown_senders is true in lead_config.json.
  2. One LLM call classifies (is it a lead?) and extracts all fields at once.
     If the client states no specific product, the model proposes best-fit
     products from the Ednex catalog, marked in a separate column.
  3. Leads are appended to the Excel tracker and the email is marked as read.
     Non-lead emails are left unread. Message IDs prevent duplicates, and
     repeat inquiries from a known client email are flagged.

Environment (LEAD_-prefixed to avoid clashing with email_bot's app):
  LEAD_CLIENT_ID, LEAD_CLIENT_SECRET  — the Lead-Email-Automation app
  LEAD_TENANT_ID                      — falls back to TENANT_ID
  OPENAI_API_KEY                      — shared with email_bot
  LEAD_CONFIG_FILE / LEAD_STATE_FILE / LEAD_TOKEN_CACHE — optional path overrides

Storage ("storage" in lead_config.json):
  "local"      — a leads.xlsx file next to this script (openpyxl).
  "onedrive"   — a workbook in the signed-in user's OneDrive via the Graph
                 Excel API (delegated auth; one-time device-code sign-in,
                 cached in token_cache_lead.json). Requires "Allow public
                 client flows" = Yes on the app registration.
  "sharepoint" — a workbook on a SharePoint site (app-only, Sites.Selected).
                 Requires IT to grant this app access to the site.
  "google_sheets" — a Google Sheet (owned by ednexautomation@gmail.com),
                 written via a Google service account (gspread). Share the
                 sheet with the service account's email as Editor. Key from
                 GOOGLE_SERVICE_ACCOUNT_JSON (env, JSON text) or the file
                 GOOGLE_SERVICE_ACCOUNT_FILE (default google_service_account.json).
                 The last column, "Zoho Lead ID", is left blank here and filled
                 by the n8n "Lead Tracker → Zoho CRM" workflow.
"""

import json
import os
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import msal
import requests
from openai import OpenAI
from openpyxl import Workbook, load_workbook

BASE_DIR = Path(__file__).resolve().parent
GRAPH = "https://graph.microsoft.com/v1.0"
TABLE_NAME = "LeadsTable"

CONFIG_PATH = BASE_DIR / os.getenv("LEAD_CONFIG_FILE", "lead_config.json")
STATE_PATH = BASE_DIR / os.getenv("LEAD_STATE_FILE", "lead_state.json")
TOKEN_CACHE_PATH = BASE_DIR / os.getenv("LEAD_TOKEN_CACHE", "token_cache_lead.json")

EXCEL_HEADERS = [
    "Logged At",
    "Lead Type",
    "Client Name",
    "Client Email",
    "Phone",
    "Company",
    "Industry",
    "Products / Requirements",
    "Proposed Products",
    "From Email",
    "Time Received",
    "Subject",
    "Notes",
    "Message ID",
]
COL_CLIENT_EMAIL = EXCEL_HEADERS.index("Client Email")
COL_MESSAGE_ID = EXCEL_HEADERS.index("Message ID")
# Google Sheet only: one extra column at the end, filled in by n8n after the
# lead is created in Zoho CRM. This bot never writes to it.
SHEET_HEADERS = EXCEL_HEADERS + ["Zoho Lead ID"]

EXTRACTION_SCHEMA = {
    "name": "lead_extraction",
    "strict": True,
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "is_lead": {
                "type": "boolean",
                "description": "True if this email is a sales lead / business inquiry about products or services.",
            },
            "confidence": {
                "type": "number",
                "description": "Confidence that this is a lead, 0.0 to 1.0.",
            },
            "client_name": {"type": "string", "description": "Name of the person inquiring. Empty string if unknown."},
            "client_email": {"type": "string", "description": "The client's own email address. Empty string if unknown."},
            "phone": {"type": "string", "description": "Client phone number. Empty string if not present."},
            "company": {"type": "string", "description": "Client's company name. Empty string if unknown."},
            "industry": {
                "type": "string",
                "description": "Client's industry (e.g. logistics, healthcare, F&B, manufacturing, retail, education, oil & gas), inferred from email content, company name, or email domain. Empty string if it cannot be reasonably determined.",
            },
            "products": {
                "type": "array",
                "description": "One entry per product the client explicitly asked about. Empty if no specific product was named.",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "product": {"type": "string", "description": "Product the client named."},
                        "requirement": {"type": "string", "description": "Their requirement for this product: quantity, specs, timeline, use case."},
                    },
                    "required": ["product", "requirement"],
                },
            },
            "proposed_products": {
                "type": "array",
                "description": "Filled ONLY when the client stated a requirement but no specific product: best-fit products chosen from the provided catalog. Empty when the client named products.",
                "items": {"type": "string"},
            },
            "general_requirement": {
                "type": "string",
                "description": "Overall requirement summary in one or two sentences. Empty string if none stated.",
            },
            "notes": {"type": "string", "description": "Any other useful detail (budget, urgency, location, competitors mentioned). Empty string if none."},
        },
        "required": [
            "is_lead", "confidence", "client_name", "client_email", "phone",
            "company", "industry", "products", "proposed_products",
            "general_requirement", "notes",
        ],
    },
}


def load_env():
    env_path = BASE_DIR / os.getenv("LEAD_ENV_FILE", ".env")
    if env_path.exists():
        for line in env_path.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, _, value = line.partition("=")
                os.environ.setdefault(key.strip(), value.strip().strip("'\""))
    os.environ.setdefault("LEAD_TENANT_ID", os.environ.get("TENANT_ID", ""))
    missing = [k for k in ("LEAD_TENANT_ID", "LEAD_CLIENT_ID", "LEAD_CLIENT_SECRET", "OPENAI_API_KEY")
               if not os.environ.get(k) or os.environ[k].startswith("PASTE_")]
    if missing:
        sys.exit(f"Missing environment values: {', '.join(missing)}")


def get_token():
    app = msal.ConfidentialClientApplication(
        os.environ["LEAD_CLIENT_ID"],
        authority=f"https://login.microsoftonline.com/{os.environ['LEAD_TENANT_ID']}",
        client_credential=os.environ["LEAD_CLIENT_SECRET"],
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        sys.exit(f"Auth failed: {result.get('error')}: {result.get('error_description')}")
    return result["access_token"]


def get_delegated_token():
    """Sign in as the user (device code, one time) for OneDrive access.

    The token is cached in token_cache_lead.json and refreshed silently on
    later runs, so the sign-in prompt only appears the first time. On Render,
    seed this file as a Secret File from a local sign-in.
    """
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        cache.deserialize(TOKEN_CACHE_PATH.read_text())
    app = msal.PublicClientApplication(
        os.environ["LEAD_CLIENT_ID"],
        authority=f"https://login.microsoftonline.com/{os.environ['LEAD_TENANT_ID']}",
        token_cache=cache,
    )
    scopes = ["Files.ReadWrite"]
    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
    if not result:
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            sys.exit(
                "Could not start device sign-in. In Entra, open the app registration >\n"
                "Authentication > Advanced settings and set 'Allow public client flows' to Yes.\n"
                f"Details: {flow.get('error_description', flow)}"
            )
        print(f"\n>>> One-time sign-in required for OneDrive access:\n>>> {flow['message']}\n", flush=True)
        result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        sys.exit(f"OneDrive sign-in failed: {result.get('error')}: {result.get('error_description')}")
    if cache.has_state_changed:
        try:
            TOKEN_CACHE_PATH.write_text(cache.serialize())
            os.chmod(TOKEN_CACHE_PATH, 0o600)
        except OSError as exc:
            print(f"Warning: could not persist token cache: {exc}", flush=True)
    return result["access_token"]


# ---------------------------------------------------------------- mailbox ----

def fetch_messages(token, mailbox, since_iso):
    """Return all messages received after since_iso, oldest first."""
    url = (
        f"{GRAPH}/users/{mailbox}/messages"
        f"?$filter=receivedDateTime ge {since_iso}"
        "&$select=id,subject,from,receivedDateTime,body,bodyPreview,isRead"
        "&$orderby=receivedDateTime asc&$top=50"
    )
    headers = {"Authorization": f"Bearer {token}"}
    messages = []
    while url:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        messages.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return messages


def mark_read(token, mailbox, message_id):
    requests.patch(
        f"{GRAPH}/users/{mailbox}/messages/{message_id}",
        headers={"Authorization": f"Bearer {token}"},
        json={"isRead": True},
        timeout=30,
    ).raise_for_status()


# ------------------------------------------------------------- extraction ----

def strip_html(html):
    text = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", html, flags=re.S | re.I)
    text = re.sub(r"<br\s*/?>|</p>|</div>|</tr>", "\n", text, flags=re.I)
    text = re.sub(r"</t[dh]>", ": ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = text.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    return "\n".join(line.strip() for line in text.splitlines() if line.strip())


def analyze_email(client, model, catalog, from_email, subject, body_text):
    """One LLM call: classify lead vs not, and extract all fields."""
    catalog_text = "\n".join(f"- {p}" for p in catalog)
    prompt = (
        "You process inbound emails for Ednex, a robotics and automation company. "
        "Decide whether the email below is a sales lead (a business inquiry about "
        "products, solutions, or services) and extract the lead details.\n\n"
        "Rules:\n"
        "- Newsletters, invoices, marketing blasts, job applications, and internal mail are NOT leads.\n"
        "- If the client names specific products, list each with its requirement in `products` "
        "and leave `proposed_products` empty.\n"
        "- If the client only describes a need without naming a product, leave `products` empty "
        "and pick the best-fit product(s) for that need from this Ednex catalog in `proposed_products`:\n"
        f"{catalog_text}\n"
        "- Infer `industry` from the content, company name, or sender domain; leave empty if unclear.\n"
        "- Never invent contact details; use empty strings for anything not present.\n\n"
        f"From: {from_email}\n"
        f"Subject: {subject}\n\n"
        f"Email body:\n{body_text[:8000]}"
    )
    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_schema", "json_schema": EXTRACTION_SCHEMA},
        temperature=0,
    )
    return json.loads(response.choices[0].message.content)


# ---------------------------------------------------------------- storage ----

class LocalExcelStore:
    """leads.xlsx next to the script, via openpyxl."""

    def __init__(self, path):
        self.path = path
        if path.exists():
            self.wb = load_workbook(path)
            self.ws = self.wb.active
        else:
            self.wb = Workbook()
            self.ws = self.wb.active
            self.ws.title = "Leads"
            self.ws.append(EXCEL_HEADERS)
        rows = [[c.value for c in row] for row in self.ws.iter_rows(min_row=2)]
        self.message_ids = {r[COL_MESSAGE_ID] for r in rows if r[COL_MESSAGE_ID]}
        self.client_emails = {str(r[COL_CLIENT_EMAIL]).lower() for r in rows if r[COL_CLIENT_EMAIL]}
        self.dirty = False

    def append(self, row):
        self.ws.append(row)
        self.dirty = True

    def flush(self):
        if self.dirty:
            self.wb.save(self.path)

    def describe(self):
        return self.path.name


class GraphExcelStore:
    """Workbook in OneDrive/SharePoint, written row-by-row via the Graph Excel API."""

    def __init__(self, token, drive_url, file_path, denied_hint, label):
        self.h = {"Authorization": f"Bearer {token}"}
        self.drive_url = drive_url
        self.file_path = "/" + file_path.strip("/")
        self.denied_hint = denied_hint
        self.label = label
        self.item_id = self._ensure_file()
        self.table_url = f"{self.drive_url}/items/{self.item_id}/workbook/tables/{TABLE_NAME}"
        self._ensure_table()

        rows = []
        url = f"{self.table_url}/rows?$select=values"
        while url:
            data = self._get(url, hint="reading existing rows")
            rows.extend(r["values"][0] for r in data.get("value", []))
            url = data.get("@odata.nextLink")
        self.message_ids = {r[COL_MESSAGE_ID] for r in rows if r[COL_MESSAGE_ID]}
        self.client_emails = {str(r[COL_CLIENT_EMAIL]).lower() for r in rows if r[COL_CLIENT_EMAIL]}

    def _fail_if_denied(self, resp, hint):
        if resp.status_code in (401, 403):
            sys.exit(
                f"Access denied while accessing {hint} (HTTP {resp.status_code}).\n"
                + self.denied_hint
            )

    def _get(self, url, hint=""):
        resp = requests.get(url, headers=self.h, timeout=30)
        self._fail_if_denied(resp, hint)
        resp.raise_for_status()
        return resp.json()

    def _ensure_file(self):
        resp = requests.get(f"{self.drive_url}/root:{self.file_path}", headers=self.h, timeout=30)
        self._fail_if_denied(resp, f"file {self.file_path}")
        if resp.status_code == 200:
            return resp.json()["id"]
        if resp.status_code != 404:
            resp.raise_for_status()
        # Create the workbook locally with headers, then upload it.
        import io
        wb = Workbook()
        ws = wb.active
        ws.title = "Leads"
        ws.append(EXCEL_HEADERS)
        buf = io.BytesIO()
        wb.save(buf)
        up = requests.put(
            f"{self.drive_url}/root:{self.file_path}:/content",
            headers={**self.h, "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
            data=buf.getvalue(), timeout=60,
        )
        self._fail_if_denied(up, f"creating {self.file_path}")
        up.raise_for_status()
        print(f"Created {self.file_path}.", flush=True)
        return up.json()["id"]

    def _ensure_table(self):
        resp = requests.get(self.table_url, headers=self.h, timeout=30)
        if resp.status_code == 200:
            return
        last_col = chr(ord("A") + len(EXCEL_HEADERS) - 1)
        created = requests.post(
            f"{self.drive_url}/items/{self.item_id}/workbook/tables/add",
            headers=self.h,
            json={"address": f"Leads!A1:{last_col}1", "hasHeaders": True},
            timeout=30,
        )
        self._fail_if_denied(created, "creating table")
        created.raise_for_status()
        requests.patch(
            f"{self.drive_url}/items/{self.item_id}/workbook/tables/{created.json()['id']}",
            headers=self.h, json={"name": TABLE_NAME}, timeout=30,
        ).raise_for_status()
        self._hide_message_id_column()

    def _hide_message_id_column(self):
        col_letter = chr(ord("A") + COL_MESSAGE_ID)
        requests.patch(
            f"{self.drive_url}/items/{self.item_id}/workbook/worksheets('Leads')/range(address='{col_letter}1')",
            headers=self.h, json={"columnHidden": True}, timeout=30,
        ).raise_for_status()

    def append(self, row):
        resp = requests.post(
            f"{self.table_url}/rows",
            headers=self.h,
            json={"values": [[str(v) if v is not None else "" for v in row]]},
            timeout=30,
        )
        resp.raise_for_status()

    def flush(self):
        pass  # rows are written immediately

    def describe(self):
        return f"{self.label}:{self.file_path}"


class GoogleSheetStore:
    """A Google Sheet written via a Google service account (gspread).

    The sheet lives in ednexautomation@gmail.com's Drive and must be shared
    with the service account's client_email as Editor. Rows are appended
    immediately (RAW, so email text is never interpreted as a formula).
    """

    def __init__(self, cfg):
        try:
            import gspread
        except ImportError:
            sys.exit("gspread is not installed. Run: pip install -r requirements.txt")
        self._gspread = gspread

        raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
        if raw:
            info = json.loads(raw)
            gc = gspread.service_account_from_dict(info)
        else:
            key_path = BASE_DIR / os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "google_service_account.json")
            if not key_path.exists():
                sys.exit(
                    f"Google service account key not found ({key_path.name}).\n"
                    "Set GOOGLE_SERVICE_ACCOUNT_JSON, or add the key file (on Render: a Secret File)."
                )
            info = json.loads(key_path.read_text())
            gc = gspread.service_account(filename=str(key_path))
        sa_email = info.get("client_email", "the service account")

        sheet_id = (cfg.get("spreadsheet_id") or "").strip()
        if not sheet_id or sheet_id.startswith("PASTE_"):
            sys.exit("Set google_sheets.spreadsheet_id in lead_config.json (the ID from the sheet's URL).")
        try:
            sh = gc.open_by_key(sheet_id)
        except (gspread.exceptions.SpreadsheetNotFound, gspread.exceptions.APIError) as exc:
            sys.exit(
                f"Cannot open Google Sheet {sheet_id}: {exc}\n"
                f"Share the sheet with {sa_email} as Editor, and check the Google Sheets API\n"
                "is enabled in the service account's Google Cloud project."
            )

        name = (cfg.get("worksheet") or "").strip()
        if name:
            try:
                ws = sh.worksheet(name)
            except gspread.exceptions.WorksheetNotFound:
                ws = sh.add_worksheet(title=name, rows=1000, cols=len(SHEET_HEADERS))
        else:
            ws = sh.sheet1
        self.ws = ws
        self.label = f"Google Sheet:{sh.title}/{ws.title}"

        values = ws.get_all_values()
        if not values or not any(cell.strip() for cell in values[0]):
            ws.update(values=[SHEET_HEADERS], range_name="A1")
            ws.freeze(rows=1)
            print(f"Wrote header row to {self.label}.", flush=True)
            values = [SHEET_HEADERS]
        elif values[0][:len(EXCEL_HEADERS)] != EXCEL_HEADERS:
            print(f"Warning: header row in {self.label} differs from the expected columns; "
                  "rows are still appended in the standard order.", flush=True)

        rows = values[1:]
        self.message_ids = {r[COL_MESSAGE_ID] for r in rows
                            if len(r) > COL_MESSAGE_ID and r[COL_MESSAGE_ID]}
        self.client_emails = {r[COL_CLIENT_EMAIL].lower() for r in rows
                              if len(r) > COL_CLIENT_EMAIL and r[COL_CLIENT_EMAIL]}

    def append(self, row):
        self.ws.append_row(
            [str(v) if v is not None else "" for v in row],
            value_input_option="RAW",
            insert_data_option="INSERT_ROWS",
            table_range="A1",
        )

    def flush(self):
        pass  # rows are written immediately

    def describe(self):
        return self.label


def make_onedrive_store(cfg):
    """Workbook in the signed-in user's own OneDrive (delegated auth, no IT grant)."""
    token = get_delegated_token()
    hint = (
        "The signed-in account could not access its OneDrive. Delete the token cache\n"
        "file and re-run to sign in again, and check the account has a OneDrive\n"
        "provisioned (open onedrive.com as that user once)."
    )
    drive = requests.get(f"{GRAPH}/me/drive", headers={"Authorization": f"Bearer {token}"}, timeout=30)
    if drive.status_code in (401, 403):
        sys.exit(hint)
    drive.raise_for_status()
    return GraphExcelStore(token, f"{GRAPH}/drives/{drive.json()['id']}",
                           cfg["file_path"], hint, "OneDrive")


def make_sharepoint_store(token, cfg):
    """Workbook on a SharePoint site (app-only, needs Sites.Selected site grant)."""
    hint = (
        "Sites.Selected grants access to NO site until an admin grants this app\n"
        "permission on the specific site. Ask IT to run (Graph Explorer or script):\n"
        "  POST /v1.0/sites/{site-id}/permissions\n"
        '  {"roles": ["write"], "grantedToIdentities": [{"application":\n'
        f'    {{"id": "{os.environ["LEAD_CLIENT_ID"]}", "displayName": "Lead-Email-Automation"}}}}]}}'
    )
    h = {"Authorization": f"Bearer {token}"}
    hostname = cfg["hostname"]
    site_path = cfg["site_path"].rstrip("/")
    site = requests.get(f"{GRAPH}/sites/{hostname}:{site_path}", headers=h, timeout=30)
    if site.status_code in (401, 403):
        sys.exit(f"Access denied for site {hostname}{site_path}.\n" + hint)
    site.raise_for_status()
    drive = requests.get(f"{GRAPH}/sites/{site.json()['id']}/drive", headers=h, timeout=30)
    drive.raise_for_status()
    return GraphExcelStore(token, f"{GRAPH}/drives/{drive.json()['id']}",
                           cfg["file_path"], hint, "SharePoint")


# ------------------------------------------------------------------- main ----

def run_once():
    load_env()
    config = json.loads(CONFIG_PATH.read_text())
    mailbox = config["mailbox"]
    senders = {s.lower() for s in config["lead_senders"]}
    classify_unknown = config.get("classify_unknown_senders", False)
    min_confidence = config.get("min_lead_confidence", 0.6)

    if STATE_PATH.exists():
        since_iso = json.loads(STATE_PATH.read_text())["last_checked"]
    else:
        lookback = config.get("lookback_days_first_run", 7)
        since_iso = (datetime.now(timezone.utc) - timedelta(days=lookback)).strftime("%Y-%m-%dT%H:%M:%SZ")

    token = get_token()
    openai_client = OpenAI()

    storage = config.get("storage", "local")
    if storage == "onedrive":
        store = make_onedrive_store(config["onedrive"])
    elif storage == "sharepoint":
        store = make_sharepoint_store(token, config["sharepoint"])
    elif storage == "google_sheets":
        store = GoogleSheetStore(config["google_sheets"])
    else:
        store = LocalExcelStore(BASE_DIR / config["excel_file"])

    print(f"Checking {mailbox} for messages since {since_iso} ...", flush=True)
    messages = fetch_messages(token, mailbox, since_iso)
    print(f"Fetched {len(messages)} message(s). Tracker: {store.describe()}", flush=True)

    new_leads = 0
    last_received = since_iso

    for msg in messages:
        received = msg["receivedDateTime"]
        last_received = max(last_received, received)
        if msg["id"] in store.message_ids:
            continue
        from_email = (msg.get("from", {}).get("emailAddress", {}).get("address") or "").lower()
        if from_email == mailbox.lower():
            continue  # our own sent/looped mail
        if from_email not in senders and not classify_unknown:
            continue

        subject = msg.get("subject", "")
        body = msg.get("body", {})
        body_text = strip_html(body["content"]) if body.get("contentType") == "html" else body.get("content", "")

        try:
            lead = analyze_email(
                openai_client, config.get("openai_model", "gpt-4o-mini"),
                config.get("product_catalog", []), from_email, subject, body_text,
            )
        except Exception as exc:
            print(f"  ! Extraction failed for '{subject}' ({from_email}): {exc}", flush=True)
            continue  # not marked read, retried next run

        if not lead["is_lead"] or lead["confidence"] < min_confidence:
            continue  # not a lead: leave unread

        client_email = lead["client_email"].lower() or from_email
        lead_type = "Repeat inquiry" if client_email in store.client_emails else "New lead"
        store.client_emails.add(client_email)

        if lead["products"]:
            products_cell = "\n".join(
                f"{p['product']}: {p['requirement']}" if p["requirement"] else p["product"]
                for p in lead["products"]
            )
        else:
            products_cell = lead["general_requirement"]
        proposed_cell = "\n".join(lead["proposed_products"])

        store.append([
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            lead_type,
            lead["client_name"],
            lead["client_email"] or from_email,
            lead["phone"],
            lead["company"],
            lead["industry"],
            products_cell,
            proposed_cell,
            from_email,
            received,
            subject,
            lead["notes"],
            msg["id"],
        ])
        new_leads += 1
        label = lead["client_name"] or lead["company"] or from_email
        print(f"  + {lead_type}: {label} — {products_cell.splitlines()[0][:60] if products_cell else subject}", flush=True)

        if config.get("mark_as_read") and not msg.get("isRead"):
            mark_read(token, mailbox, msg["id"])

    store.flush()
    STATE_PATH.write_text(json.dumps({"last_checked": last_received}))
    print(f"Done. {new_leads} new lead(s) added.", flush=True)


if __name__ == "__main__":
    run_once()
