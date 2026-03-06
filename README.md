# Email Bot

Python inbox watcher for Microsoft 365 that:
- polls a mailbox Inbox via Microsoft Graph,
- sends a Teams channel email notification with an AI summary,
- creates a reply draft in Outlook using an LLM-generated response.

## Behavior

For each new inbox message, the bot will process it only when all checks pass:
1. sender is not likely automated (`noreply`, `notification`, etc.),
2. your mailbox address is directly in `To:`,
3. email body includes a direct greeting to `ADDRESSING_NAME` (examples: `Dear Guru`, `Hi Guru`, `Hello Guru`, `Dear Mr. Guru`, `Good morning Guru`).

If checks pass, it sends Teams notification and creates a reply draft.

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
- `MAX_THREAD_MESSAGES` (default: `5`)
- `MAX_THREAD_CHARS_PER_MESSAGE` (default: `1200`)
- `DEV_MODE` (`true`/`false`)

## Run

```bash
python3 email_bot.py
```

The bot will keep polling every 15 seconds until stopped (`Ctrl+C`).

## Notes

- If `OPENAI_API_KEY` is missing, draft replies use fallback text.
- `EMAIL_SIGNATURE` is required and appended exactly as provided in the prompt/fallback.
- Keep `.env` private; do not commit secrets.
