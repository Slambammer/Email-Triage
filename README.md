# Email Triage MVP (Outlook -> Azure OpenAI -> Slack)

This MVP polls Outlook every minute for unread emails that do **not** have the `triaged` category, sends sender/subject/body to Azure OpenAI, posts the triage output to Slack, then marks the email with `triaged` (and optionally read).

## 1) Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

Fill `.env` values.

## 2) Azure / Microsoft prerequisites

- Azure AD app registration with delegated permissions:
  - `Mail.Read`
  - `Mail.ReadWrite`
  - `User.Read`
- First run uses device-code login through `msal`.
- Mailbox UPN in `.env` must be accessible by signed-in user.

## 3) Run

```bash
python main.py
```

The worker loop:
- polls every `POLL_SECONDS` (default 60),
- fetches unread Inbox emails,
- excludes any with category `triaged`,
- calls Azure OpenAI,
- sends result to Slack webhook,
- PATCHes categories to include `triaged`.

## 4) Next refinements

- Swap Azure OpenAI adapter for Copilot Studio Direct Line adapter.
- Add local state / idempotency checks.
- Add sender allowlist / redact rules.
- Add structured JSON output enforcement from LLM.
