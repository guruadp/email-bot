# Email Bot

Python inbox watcher for Microsoft 365 that:
- polls a mailbox Inbox via Microsoft Graph,
- sends a Teams channel email notification with an AI summary,
- creates a reply draft in Outlook using an LLM-generated response.

## Behavior

For each new inbox message, the bot will process it only when all checks pass:
1. sender is not likely automated (`noreply`, `notification`, etc.),
2. your mailbox address is directly in `To:`,
3. the sender is on the VIP list (`VIP_SENDERS`/`VIP_DOMAINS`, always processed), **or** the email
   body includes a direct greeting to `ADDRESSING_NAME` (examples: `Dear Guru`, `Hi Guru`, `Hello Guru`,
   `Dear Mr. Guru`, `Good morning Guru`), **or**, when the greeting regex doesn't match and an
   `OPENAI_API_KEY` is configured, an LLM call confirms the email is genuinely addressed to
   `ADDRESSING_NAME` personally (catches phrasing like `Senthil, following up on...` or a threaded
   reply with no greeting at all, which the regex alone would miss).

If checks pass, it sends Teams notification and creates a reply draft.

## Files

| File | What it does |
|---|---|
| `email_bot.py` | The inbox watcher: polls the mailbox every 15 s, sends Teams notifications with AI summaries, creates LLM reply drafts. |
| `lead_automation.py` | The lead tracker: classifies Sales-mailbox emails with an LLM, extracts lead details, appends them to the Excel tracker in OneDrive. Runs one cycle and exits. See "Lead automation" below. |
| `start.py` | Production entry point: runs `email_bot.py` continuously (auto-restart on crash) and a `lead_automation.py` cycle every `LEAD_POLL_MINUTES` (default 15). This is the Render start command. |
| `lead_config.json` | Lead automation settings: mailbox, storage mode, lead-sender whitelist, confidence threshold, Ednex product catalog. |
| `requirements.txt` | Python dependencies for both bots. |
| `.env.example` | Template for the local `.env` — copy and fill in. |
| `.env` | Your local secrets (gitignored, never commit). On Render, use environment variables instead. |
| `email-bot.service.example` | systemd unit template for running on a Linux server instead of Render. |
| `application-access-policy.example.ps1` | Exchange PowerShell runbook for IT: restricts each app registration to only its own mailbox. |
| `seen_ids.json` | Generated state: message IDs email_bot has already processed. |
| `lead_state.json` | Generated state: last time lead_automation checked the mailbox. |
| `token_cache.bin` / `token_cache_lead.json` | Generated auth token caches (gitignored). `token_cache_lead.json` holds the OneDrive sign-in — see "Lead automation". |
| `leads.xlsx` | The lead tracker — only present locally when `storage` is `local`; normally it lives in OneDrive. |

## Requirements

- Python 3.10+
- Microsoft Graph app with **application** permissions and admin consent (including `Mail.ReadWrite`)
- Mailbox user access in your tenant
- (Optional) OpenAI API key for AI summary/reply generation

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Copy env template and fill values:
```bash
cp .env.example .env
```

3. Required `.env` keys:
- `TENANT_ID`
- `CLIENT_ID`
- `CLIENT_SECRET`
- `MAILBOX_USER`
- `TEAMS_CHANNEL_EMAIL`
- `EMAIL_SIGNATURE` (use `\n` for line breaks)

4. Common optional keys:
- `OPENAI_API_KEY`
- `OPENAI_MODEL` (default: `gpt-4o-mini`)
- `ADDRESSING_NAME` (default in code: `Guru`)
- `PERSONA_DESCRIPTION` — who the LLM writes as, e.g. `Senthil Kugan, CEO of Ednex LLC, Ednex Automation` (default in code is Guru's; **always set this explicitly per mailbox**, otherwise every drafted reply introduces itself as the default persona regardless of whose inbox is running)
- `MAX_THREAD_MESSAGES` (default: `5`)
- `MAX_THREAD_CHARS_PER_MESSAGE` (default: `1200`)
- `DEV_MODE` (`true`/`false`)
- `SEEN_IDS_FILE` (default: `seen_ids.json`) — where processed-message state is persisted; set this to a distinct path per mailbox if running multiple instances from the same working directory
- `VIP_SENDERS` — comma-separated exact sender addresses that always bypass the greeting check (e.g. `investor@fund.com,client@bigco.com`)
- `VIP_DOMAINS` — comma-separated sender domains that always bypass the greeting check (e.g. `bigco.com`)
- `N8N_ACTION_WEBHOOK_URL` — if set, the Teams notification includes Send/Discard buttons for the drafted reply (see below). Leave unset to keep the old passive-notification behavior.
- `LINK_EXPIRY_DAYS` (default: `3`) — how long a Send/Discard link stays valid before it shows "expired" instead of acting

5. Lead automation keys (only if you run `lead_automation.py` / `start.py`):
- `LEAD_CLIENT_ID` / `LEAD_CLIENT_SECRET` — the **Lead-Email-Automation** app registration (a separate app from the email bot's, hence the prefix)
- `LEAD_TENANT_ID` — optional, falls back to `TENANT_ID`
- `LEAD_POLL_MINUTES` — optional, minutes between lead cycles (default: `15`)
- `OPENAI_API_KEY` — required for lead classification (shared with the email bot)

## Run

```bash
python3 email_bot.py
```

The bot will keep polling every 15 seconds until stopped (`Ctrl+C`).

### Running continuously (recommended for production)

The bot is a long-running foreground process with no built-in supervision — if it
crashes, nothing restarts it. Use `email-bot.service.example` as a template for a
systemd unit with `Restart=always`:

```bash
sudo cp email-bot.service.example /etc/systemd/system/email-bot-senthil.service
# edit the placeholders inside, then:
sudo systemctl daemon-reload
sudo systemctl enable --now email-bot-senthil
journalctl -u email-bot-senthil -f   # tail logs
```

## Deploying on Render

The repo runs as a single **Background Worker** hosting both bots.

1. **Service settings** (Render dashboard → the worker → Settings):
   - Build command: `pip install -r requirements.txt`
   - Start command: `python3 start.py`
2. **Environment variables** (Render dashboard → Environment). There is no
   `.env` on Render — every key goes here instead:

   | Key | For | Notes |
   |---|---|---|
   | `TENANT_ID` | both | Entra tenant ID (shared) |
   | `CLIENT_ID` / `CLIENT_SECRET` | email bot | email bot's app registration |
   | `MAILBOX_USER` | email bot | mailbox the email bot watches |
   | `TEAMS_CHANNEL_EMAIL` | email bot | Teams channel address for notifications |
   | `EMAIL_SIGNATURE` / `ADDRESSING_NAME` / `PERSONA_DESCRIPTION` | email bot | per-mailbox identity |
   | `OPENAI_API_KEY` | both | summaries, reply drafts, lead classification |
   | `LEAD_CLIENT_ID` / `LEAD_CLIENT_SECRET` | lead automation | Lead-Email-Automation app registration |
   | `LEAD_POLL_MINUTES` | lead automation | optional, default 15 |
   | `VIP_SENDERS` / `VIP_DOMAINS` / `N8N_ACTION_WEBHOOK_URL` / `LINK_EXPIRY_DAYS` / `DEV_MODE` | email bot | optional, as in Setup above |

3. **Secret file** (Render dashboard → Environment → Secret Files): after the
   one-time local OneDrive sign-in (see "Lead automation" below), add
   `token_cache_lead.json` with the file's contents. Filename exactly
   `token_cache_lead.json` so it lands in the repo root at deploy.
4. Push to `main` — Render builds and redeploys automatically. Check the
   worker's **Logs** tab: you should see `[start] email_bot started` and a
   `running lead_automation cycle` line every poll interval.

Note on state: Render's disk is ephemeral. `seen_ids.json` and
`lead_state.json` reset on every deploy; both bots are designed to recover
from that without double-processing anything visible (see "Reliability notes"
and "Lead-automation state on redeploys").

## Interactive Teams approval

When `N8N_ACTION_WEBHOOK_URL` is set, each Teams notification includes **Send Reply**
and **Discard Draft** buttons instead of requiring you to leave Teams and go find the
draft in Outlook. Clicking either calls an n8n workflow
(`Email Bot - Approve/Discard Action`, https://ednexautomation.app.n8n.cloud/workflow/y8dFbKS1dz1Wm5mi)
that sends or deletes the exact draft via Microsoft Graph and shows a plain
confirmation page.

**One-time setup required in n8n** — the workflow was created with no credentials
wired in on purpose (secrets are never generated or handled outside n8n's own
credential store):

1. Open the workflow in n8n. Each of the four HTTP Request nodes ("Guru: Send
   Draft", "Guru: Discard Draft", "Senthil: Send Draft", "Senthil: Discard Draft")
   needs its `oAuth2Api` credential configured:
   - Grant Type: **Client Credentials**
   - Access Token URL: `https://login.microsoftonline.com/424eb54f-7164-4789-b988-551caa4351c6/oauth2/v2.0/token`
   - Client ID / Client Secret: the matching app's values from `.env_guru` or `.env_senthil`
   - Scope: `https://graph.microsoft.com/.default`
   - Authentication: **Send as Body** (Microsoft's token endpoint expects
     client_secret in the POST body, not a Basic Auth header)
2. Save each credential, then re-save/activate the workflow.

Security model: each Send/Discard link embeds the mailbox, the Graph draft ID
(itself a long unguessable Graph object ID), an action, and an expiry timestamp —
there's no separate secret to leak, and a link stops working after `LINK_EXPIRY_DAYS`.
Anyone who can see the Teams channel message can already read the full email and AI
summary today; the buttons don't widen that trust boundary, they just act within it.
If the Application Access Policy work below is applied, each app can only touch its
own mailbox even with a valid link for the wrong one.

## Lead automation

`lead_automation.py` is a second bot in this repo: it reads the Sales mailbox
(`Sales@ednex.me`), uses an LLM (`gpt-4o-mini`) to decide whether each email is a
sales lead, extracts the details (client name/email/phone, company, industry,
products with per-product requirements — or proposed products from the Ednex
catalog when the client only describes a need), and appends new leads to an
Excel tracker (`leads.xlsx`) in the Sales account's OneDrive. Leads are marked
as read in the mailbox; non-lead emails are left unread. Message IDs stored in
the sheet prevent duplicate rows, and repeat inquiries from a known client
email are flagged as `Repeat inquiry` instead of `New lead`.

Both bots run in one Render Background Worker via `start.py` (set the Render
start command to `python3 start.py`): it keeps `email_bot.py` running,
restarts it if it crashes, and runs one `lead_automation.py` cycle every
`LEAD_POLL_MINUTES` (default 15).

### Configuration

- `lead_config.json` — mailbox, storage mode (`local` / `onedrive` /
  `sharepoint`), lead-source sender whitelist, `classify_unknown_senders`
  (when `true`, every unknown sender is classified by the LLM instead of
  skipped), confidence threshold, and the Ednex product catalog used for
  proposed products. Edit the catalog to match the real portfolio — proposals
  are only as good as this list.
- Env keys (see `.env.example`): `LEAD_CLIENT_ID` / `LEAD_CLIENT_SECRET` are
  the **Lead-Email-Automation** app registration (separate from the email-bot
  app, hence the prefix); `LEAD_TENANT_ID` falls back to `TENANT_ID`;
  `OPENAI_API_KEY` is shared.

### OneDrive sign-in (`token_cache_lead.json`)

Mail reading uses app-only auth, but writing the Excel to OneDrive uses
**delegated** auth as the Sales user (`Files.ReadWrite` via device-code flow)
so no extra IT/admin consent is needed. The app registration must have
"Allow public client flows" enabled (Authentication > Advanced settings).

Setup is one-time: run `python3 lead_automation.py` locally, follow the
printed device-code prompt (sign in as the Sales account), then upload the
resulting `token_cache_lead.json` as a Render **Secret File** (same filename,
repo root). The cached refresh token renews itself silently on every run;
Render restarts re-seed from the secret file.

Re-doing the sign-in is only needed when the refresh token is invalidated:

- the Sales account password is changed (most common trigger),
- an admin revokes the account's sessions or a Conditional Access policy
  forces periodic re-authentication,
- the bot is offline for 90+ days (inactivity expiry).

You'll know it's needed when the Render logs show the
`>>> One-time sign-in required` prompt again and no leads get logged. Fix:
delete the local `token_cache_lead.json`, run the script locally, sign in,
upload the new file to Render's Secret Files, redeploy.

### Lead-automation state on redeploys

`lead_state.json` (last-checked timestamp) lives on Render's ephemeral disk
and resets on every deploy; the script then re-scans the
`lookback_days_first_run` window. Message IDs already in the Excel prevent
duplicates, so a redeploy only costs a few redundant LLM calls on emails that
were previously classified as non-leads.

## Reliability notes

- Processed message IDs are persisted to `SEEN_IDS_FILE` after every poll (atomic
  write). On restart, the bot resumes from that state instead of blindly marking
  the current inbox snapshot as seen — so a crash or redeploy no longer causes
  in-flight mail to be silently skipped. Only the very first run (no state file
  yet) primes without processing, to avoid replaying your whole existing inbox.
- Reply and summary generation both fall back to plain text on any LLM failure
  and reuse the Graph HTTP session's retry/backoff, so a transient OpenAI error
  no longer drops the entire message (Teams notification still fires even if
  drafting fails).
- Still a known gap: `SEEN_IDS_FILE` is not distributed state — running the same
  mailbox from two machines/processes at once will double-process. Run one
  instance per mailbox.
- The LLM personal-address check only runs when the greeting regex doesn't already
  match, so it's a small number of extra `gpt-4o-mini` calls per day, not one per
  email. If that call fails or `OPENAI_API_KEY` is unset, it fails safe to "skip"
  (same as today's regex-only behavior) — it can only catch more real emails than
  before, never fewer.

## Security hardening

Both app registrations currently hold **tenant-wide** `Mail.ReadWrite`/`Mail.Send`
application permissions — a leaked `CLIENT_SECRET` grants access to every mailbox
in the tenant, not just the one it's meant for. `application-access-policy.example.ps1`
is a ready-to-run Exchange Online PowerShell runbook that restricts each app to only
its own mailbox via `New-ApplicationAccessPolicy`. Requires Exchange admin rights;
not something this bot can apply on its own.

## Notes

- If `OPENAI_API_KEY` is missing, draft replies use fallback text.
- `EMAIL_SIGNATURE` is required and appended exactly as provided in the prompt/fallback.
- Keep `.env` private; do not commit secrets.
