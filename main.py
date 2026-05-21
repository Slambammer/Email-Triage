import json
import os
import time
from typing import Dict, List

import msal
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv

load_dotenv()

GRAPH_TENANT_ID = os.environ["GRAPH_TENANT_ID"]
GRAPH_CLIENT_ID = os.environ["GRAPH_CLIENT_ID"]
MAILBOX_UPN = os.environ["MAILBOX_USER_PRINCIPAL_NAME"]
SLACK_WEBHOOK_URL = os.environ["SLACK_WEBHOOK_URL"]

TRIAGED_CATEGORY = os.getenv("TRIAGED_CATEGORY", "triaged")
POLL_SECONDS = int(os.getenv("POLL_SECONDS", "60"))
MARK_READ = os.getenv("MARK_READ", "true").lower() == "true"
MAX_BODY_CHARS = int(os.getenv("MAX_BODY_CHARS", "8000"))

AZURE_OPENAI_ENDPOINT = os.environ["AZURE_OPENAI_ENDPOINT"].rstrip("/")
AZURE_OPENAI_API_KEY = os.environ["AZURE_OPENAI_API_KEY"]
AZURE_OPENAI_DEPLOYMENT = os.environ["AZURE_OPENAI_DEPLOYMENT"]
AZURE_OPENAI_API_VERSION = os.getenv("AZURE_OPENAI_API_VERSION", "2025-01-01-preview")

SCOPES = ["Mail.ReadWrite", "Mail.Read", "User.Read"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def get_graph_token() -> str:
    authority = f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}"
    app = msal.PublicClientApplication(client_id=GRAPH_CLIENT_ID, authority=authority)

    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(scopes=SCOPES, account=accounts[0])
    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError("Failed to start device code flow")
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"Could not obtain Graph token: {result}")
    return result["access_token"]


def graph_get_unread(token: str) -> List[Dict]:
    headers = {"Authorization": f"Bearer {token}"}
    params = {
        "$select": "id,subject,from,body,categories,isRead,receivedDateTime,webLink",
        "$filter": "isRead eq false",
        "$orderby": "receivedDateTime desc",
        "$top": "25",
    }
    url = f"{GRAPH_BASE}/users/{MAILBOX_UPN}/mailFolders/Inbox/messages"
    resp = requests.get(url, headers=headers, params=params, timeout=30)
    resp.raise_for_status()
    messages = resp.json().get("value", [])
    return [m for m in messages if TRIAGED_CATEGORY.lower() not in [c.lower() for c in m.get("categories", [])]]


def html_to_text(html: str) -> str:
    soup = BeautifulSoup(html or "", "html.parser")
    return soup.get_text("\n", strip=True)


def call_azure_openai(sender: str, subject: str, body: str) -> str:
    prompt = (
        "You are an email triage assistant. Return concise markdown with:\n"
        "1) urgency (low/medium/high),\n"
        "2) 2-3 bullet summary,\n"
        "3) recommended action,\n"
        "4) draft reply (short)."
    )
    payload = {
        "messages": [
            {"role": "system", "content": prompt},
            {
                "role": "user",
                "content": f"Sender: {sender}\nSubject: {subject}\nBody:\n{body[:MAX_BODY_CHARS]}",
            },
        ],
        "temperature": 0.2,
    }
    url = (
        f"{AZURE_OPENAI_ENDPOINT}/openai/deployments/{AZURE_OPENAI_DEPLOYMENT}/chat/completions"
        f"?api-version={AZURE_OPENAI_API_VERSION}"
    )
    headers = {"api-key": AZURE_OPENAI_API_KEY, "Content-Type": "application/json"}
    resp = requests.post(url, headers=headers, data=json.dumps(payload), timeout=60)
    resp.raise_for_status()
    data = resp.json()
    return data["choices"][0]["message"]["content"]


def post_to_slack(message: str) -> None:
    resp = requests.post(SLACK_WEBHOOK_URL, json={"text": message}, timeout=30)
    resp.raise_for_status()


def mark_triaged(token: str, message_id: str, existing_categories: List[str]) -> None:
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    categories = list(dict.fromkeys(existing_categories + [TRIAGED_CATEGORY]))
    body = {"categories": categories}
    if MARK_READ:
        body["isRead"] = True

    url = f"{GRAPH_BASE}/users/{MAILBOX_UPN}/messages/{message_id}"
    resp = requests.patch(url, headers=headers, data=json.dumps(body), timeout=30)
    resp.raise_for_status()


def process_once(token: str) -> int:
    emails = graph_get_unread(token)
    for e in emails:
        sender = (((e.get("from") or {}).get("emailAddress") or {}).get("address") or "unknown")
        subject = e.get("subject") or "(no subject)"
        body_text = html_to_text((e.get("body") or {}).get("content") or "")

        triage = call_azure_openai(sender=sender, subject=subject, body=body_text)
        slack_text = (
            f"*New triaged email*\n"
            f"*From:* {sender}\n"
            f"*Subject:* {subject}\n"
            f"*Outlook Link:* {e.get('webLink','')}\n\n"
            f"{triage}"
        )
        post_to_slack(slack_text)
        mark_triaged(token, e["id"], e.get("categories", []))
        print(f"Processed: {subject} ({e['id']})")
    return len(emails)


def main() -> None:
    token = get_graph_token()
    print("Email triage worker started.")
    while True:
        try:
            count = process_once(token)
            print(f"Cycle complete, processed={count}")
        except Exception as exc:
            print(f"Error in cycle: {exc}")
            token = get_graph_token()
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
