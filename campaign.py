"""
CKC Capital IR Campaign Runner
Usage:
  python campaign.py --subject "Q2 2026 Investor Update" --body body.txt
  python campaign.py --subject "Q2 2026 Investor Update" --body body.txt --send
  python campaign.py --send-only
"""

import argparse
import csv
import json
import os
import sys
import time
import httpx

# ── Config ────────────────────────────────────────────────
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
AGENT_ID          = "agent_01QdK4mGu4Ccrkvk9Bnm9cM7"
ENVIRONMENT_ID    = "env_01ED74h2MAhHFn43jdfV7QZy"
CONTACTS_FILE     = "contacts.csv"
DRAFTS_FILE       = "drafts.json"
POLL_INTERVAL     = 5
MAX_WAIT          = 120

HEADERS = {
    "x-api-key":        ANTHROPIC_API_KEY,
    "anthropic-version": "2023-06-01",
    "anthropic-beta":   "managed-agents-2026-04-01",
    "content-type":     "application/json"
}
BASE = "https://api.anthropic.com/v1"


def api(method: str, path: str, **kwargs) -> dict:
    with httpx.Client(timeout=30.0) as client:
        resp = client.request(method, f"{BASE}{path}", headers=HEADERS, **kwargs)
        resp.raise_for_status()
        return resp.json()


def load_contacts(filepath: str) -> list[dict]:
    contacts = []
    with open(filepath, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            contacts.append({"name": row["Name"].strip(), "email": row["Email"].strip()})
    return contacts


def create_session(title: str) -> str:
    data = api("POST", "/sessions", json={
        "agent": AGENT_ID,
        "environment_id": ENVIRONMENT_ID,
        "title": title
    })
    return data["id"]


def send_message(session_id: str, text: str):
    api("POST", f"/sessions/{session_id}/events", json={
        "events": [{"type": "user.message", "content": [{"type": "text", "text": text}]}]
    })


def wait_for_response(session_id: str) -> str:
    elapsed = 0
    last_message = ""
    while elapsed < MAX_WAIT:
        time.sleep(POLL_INTERVAL)
        elapsed += POLL_INTERVAL
        data = api("GET", f"/sessions/{session_id}/events")
        status = "running"
        for event in data.get("data", []):
            if event.get("type") == "session.status_idle":
                status = "idle"
            if event.get("type") == "agent.message":
                for block in event.get("content", []):
                    if block.get("type") == "text":
                        last_message = block["text"]
        if status == "idle":
            break
    return last_message


def generate_draft(contact: dict, subject: str, body: str) -> dict:
    print(f"  → Creating session for {contact['name']}...")
    session_id = create_session(f"IR Campaign - {contact['name']}")

    prompt = f"""You are drafting a personalized IR outreach email for CKC Capital.

Contact: {contact['name']} ({contact['email']})

Step 1: Use search_email_history to retrieve all prior emails exchanged with {contact['email']}.
        Read the subjects AND previews carefully to understand:
        - What topics have been discussed before
        - What stage of the relationship this is (first contact, ongoing, etc.)
        - Any specific interests, concerns, or questions they have raised
        - Any commitments or follow-ups that were made

Step 2: Write a personalized email that feels like a natural continuation of the relationship.
        - Start with Hi {contact['name']},
        - If prior history exists:
          * Reference specific topics or discussions naturally e.g. "Following up on our conversation about X..."
          * Build on where the conversation left off
          * Do NOT just say "as we discussed previously" — be specific
          * The email should feel like it was written by someone who knows this person
        - If no prior history:
          * Write a warm, professional first introduction
          * No references to prior conversations
        - Incorporate the base marketing content below naturally into the email
        - Do NOT just paste the base copy — weave it into the context of the relationship
        - Format the email with proper paragraphs, not a wall of text
        - Use HTML formatting: wrap paragraphs in <p> tags
        - Sign off as: The IR Team, CKC Capital

Base email copy to incorporate:
{body}

Return ONLY the final email body as HTML. No subject line. No explanations. No markdown."""

    print(f"  → Generating personalized draft...")
    send_message(session_id, prompt)
    draft_body = wait_for_response(session_id)

    return {
        "name": contact["name"],
        "email": contact["email"],
        "subject": subject,
        "body": draft_body,
        "session_id": session_id,
        "status": "draft"
    }


def send_draft(draft: dict) -> dict:
    session_id = create_session(f"IR Send - {draft['name']}")
    prompt = f"""Send this email using the send_email tool. Preserve the HTML formatting exactly.
To: {draft['email']}
Subject: {draft['subject']}

{draft['body']}

Use send_email tool with the body exactly as provided above. Do not modify the content."""
    send_message(session_id, prompt)
    wait_for_response(session_id)
    draft["status"] = "sent"
    print(f"  ✓ Sent to {draft['email']}")
    return draft


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--subject", help="Email subject line")
    parser.add_argument("--body", help="Path to base email body text file")
    parser.add_argument("--send", action="store_true", help="Send after generating")
    parser.add_argument("--send-only", action="store_true", help="Send existing drafts")
    args = parser.parse_args()

    if args.send_only:
        if not os.path.exists(DRAFTS_FILE):
            print(f"No drafts file found at {DRAFTS_FILE}")
            sys.exit(1)
        with open(DRAFTS_FILE) as f:
            drafts = json.load(f)
        pending = [d for d in drafts if d["status"] == "draft"]
        print(f"\nSending {len(pending)} drafts...\n")
        for draft in pending:
            send_draft(draft)
        with open(DRAFTS_FILE, "w") as f:
            json.dump(drafts, f, indent=2)
        print(f"\n✓ Done.")
        return

    if not args.subject or not args.body:
        print("--subject and --body are required")
        sys.exit(1)

    if not os.path.exists(args.body):
        print(f"Body file not found: {args.body}")
        sys.exit(1)

    with open(args.body, encoding="utf-8") as f:
        body = f.read().strip()

    contacts = load_contacts(CONTACTS_FILE)
    print(f"\nCKC Capital IR Campaign")
    print(f"Subject: {args.subject}")
    print(f"Contacts: {len(contacts)}")
    print(f"{'─' * 40}\n")

    drafts = []
    for i, contact in enumerate(contacts, 1):
        print(f"[{i}/{len(contacts)}] {contact['name']} <{contact['email']}>")
        try:
            draft = generate_draft(contact, args.subject, body)
            drafts.append(draft)
            print(f"  ✓ Draft ready\n")
        except Exception as e:
            print(f"  ✗ Error: {e}\n")
            drafts.append({
                "name": contact["name"],
                "email": contact["email"],
                "subject": args.subject,
                "body": "",
                "status": "error",
                "error": str(e)
            })
        time.sleep(2)

    with open(DRAFTS_FILE, "w") as f:
        json.dump(drafts, f, indent=2)

    print(f"{'─' * 40}")
    print(f"✓ {len([d for d in drafts if d['status'] == 'draft'])} drafts saved to {DRAFTS_FILE}")
    print(f"✗ {len([d for d in drafts if d['status'] == 'error'])} errors\n")

    print("DRAFT PREVIEW")
    print(f"{'─' * 40}")
    for draft in drafts:
        if draft["status"] == "draft":
            print(f"\nTo: {draft['name']} <{draft['email']}>")
            print(f"Subject: {draft['subject']}")
            print(f"{'─' * 40}")
            print(draft["body"][:500] + "..." if len(draft["body"]) > 500 else draft["body"])
            print()

    if args.send:
        confirm = input("Send all drafts? (yes/no): ").strip().lower()
        if confirm == "yes":
            print(f"\nSending...\n")
            for draft in drafts:
                if draft["status"] == "draft":
                    send_draft(draft)
            with open(DRAFTS_FILE, "w") as f:
                json.dump(drafts, f, indent=2)
            print(f"\n✓ Campaign complete.")
        else:
            print("\nNot sent. Run with --send-only when ready.")
    else:
        print(f"Review drafts.json then run with --send-only to send.")


if __name__ == "__main__":
    main()
