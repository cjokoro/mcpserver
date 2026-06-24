import httpx
import os
import json
import csv
import re
import uvicorn
from mcp.server.fastmcp import FastMCP
from mcp.server.transport_security import TransportSecuritySettings
from starlette.requests import Request
from starlette.responses import PlainTextResponse, JSONResponse
from starlette.routing import Route, Mount
from starlette.applications import Starlette

# ── Config ────────────────────────────────────────────────
CLIENT_ID      = os.environ["AZURE_CLIENT_ID"]
TENANT_ID      = os.environ["AZURE_TENANT_ID"]
CLIENT_SECRET  = os.environ["AZURE_CLIENT_SECRET"]
GRAPH_BASE     = "https://graph.microsoft.com/v1.0"
IR_EMAIL       = os.environ.get("IR_EMAIL", "ops@ckccapital.com")
FROM_EMAIL     = os.environ.get("FROM_EMAIL", "cokoro@ckccapital.com")
ANTHROPIC_KEY  = os.environ.get("ANTHROPIC_API_KEY", "")
AGENT_ID       = os.environ.get("AGENT_ID", "agent_01QdK4mGu4Ccrkvk9Bnm9cM7")
ENVIRONMENT_ID = os.environ.get("ENVIRONMENT_ID", "env_01ED74h2MAhHFn43jdfV7QZy")
PORT           = int(os.environ.get("PORT", 8000))
CONTACTS_FILE  = os.environ.get("CONTACTS_FILE", "/opt/render/project/src/contacts.csv")

# ── Load approved contacts ────────────────────────────────
def load_contacts() -> dict:
    """Load contacts.csv and return dict of {email: name}."""
    contacts = {}
    try:
        with open(CONTACTS_FILE, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                email = row.get("Email", "").strip().lower()
                name  = row.get("Name", "").strip()
                if email:
                    contacts[email] = name
    except Exception as e:
        print(f"[CONTACTS] Could not load contacts file: {e}")
    return contacts

# ── State file for pending scheduling approvals ───────────
STATE_FILE = "/tmp/scheduling_state.json"

def load_state() -> dict:
    try:
        with open(STATE_FILE) as f:
            return json.load(f)
    except:
        return {}

def save_state(state: dict):
    with open(STATE_FILE, "w") as f:
        json.dump(state, f, indent=2)

# ── FastMCP server ────────────────────────────────────────
mcp = FastMCP(
    "CKC IR Server",
    transport_security=TransportSecuritySettings(
        enable_dns_rebinding_protection=False
    )
)

# ── Token management ──────────────────────────────────────
_token_cache = {"token": None}

async def get_token():
    if _token_cache["token"]:
        return _token_cache["token"]
    async with httpx.AsyncClient(timeout=30.0) as client:
        resp = await client.post(
            f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
            data={
                "client_id":     CLIENT_ID,
                "client_secret": CLIENT_SECRET,
                "scope":         "https://graph.microsoft.com/.default",
                "grant_type":    "client_credentials"
            }
        )
        resp.raise_for_status()
        _token_cache["token"] = resp.json()["access_token"]
        return _token_cache["token"]

async def graph(method: str, path: str, **kwargs):
    token = await get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    async with httpx.AsyncClient(timeout=30.0) as client:
        resp = await client.request(
            method, f"{GRAPH_BASE}{path}", headers=headers, **kwargs
        )
        resp.raise_for_status()
        return resp.json() if resp.content else {}

# ── MCP Tools ─────────────────────────────────────────────

@mcp.tool()
async def search_email_history(email: str, limit: int = 5) -> str:
    """Search Outlook inbox for prior correspondence with a contact."""
    data = await graph(
        "GET",
        f"/users/{FROM_EMAIL}/messages"
        f"?$filter=from/emailAddress/address eq '{email}'"
        f"&$top={limit}"
        f"&$select=subject,from,receivedDateTime,bodyPreview,id"
    )
    msgs = data.get("value", [])
    if not msgs:
        return f"No prior correspondence found with {email}"
    lines = [f"Found {len(msgs)} prior email(s) with {email}:"]
    for m in msgs:
        lines.append(f"- [{m['receivedDateTime'][:10]}] {m['subject']} | {m['bodyPreview'][:100]}")
    return "\n".join(lines)


@mcp.tool()
async def send_email(to: str, subject: str, body: str) -> str:
    """Send an email from cokoro@ckccapital.com via Microsoft Graph."""
    await graph("POST", f"/users/{FROM_EMAIL}/sendMail", json={
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body},
            "toRecipients": [{"emailAddress": {"address": to}}]
        }
    })
    return f"Email sent to {to}"


@mcp.tool()
async def forward_email(message_id: str, classification: str) -> str:
    """Forward a reply to ops@ckccapital.com with Claude classification prepended."""
    comment = f"[CKC IR CLASSIFICATION]\n{classification}\n\n[ORIGINAL MESSAGE]"
    await graph(
        "POST",
        f"/users/{FROM_EMAIL}/messages/{message_id}/forward",
        json={
            "comment": comment,
            "toRecipients": [{"emailAddress": {"address": IR_EMAIL}}]
        }
    )
    return f"Forwarded to {IR_EMAIL}"


@mcp.tool()
async def check_availability(emails: list[str], start_time: str, end_time: str) -> str:
    """Check free/busy availability for a list of attendees."""
    data = await graph("POST", f"/users/{FROM_EMAIL}/calendar/getSchedule", json={
        "schedules":   emails,
        "startTime":   {"dateTime": start_time, "timeZone": "Eastern Standard Time"},
        "endTime":     {"dateTime": end_time,   "timeZone": "Eastern Standard Time"},
        "availabilityViewInterval": 30
    })
    return str(data)


@mcp.tool()
async def create_calendar_event(
    title: str, attendees: list[str], start: str, end: str, body: str = ""
) -> str:
    """Create a calendar event and send invites to all attendees."""
    attendee_list = [
        {"emailAddress": {"address": e}, "type": "required"}
        for e in attendees
    ]
    await graph("POST", f"/users/{FROM_EMAIL}/events", json={
        "subject":   title,
        "body":      {"contentType": "HTML", "content": body},
        "start":     {"dateTime": start, "timeZone": "Eastern Standard Time"},
        "end":       {"dateTime": end,   "timeZone": "Eastern Standard Time"},
        "attendees": attendee_list
    })
    return f"Calendar event created: {title}"


# ── Agent session helper ──────────────────────────────────

async def run_agent_session(title: str, prompt: str) -> str:
    headers = {
        "x-api-key":         ANTHROPIC_KEY,
        "anthropic-version": "2023-06-01",
        "anthropic-beta":    "managed-agents-2026-04-01",
        "content-type":      "application/json"
    }
    async with httpx.AsyncClient(timeout=60.0) as client:
        resp = await client.post(
            "https://api.anthropic.com/v1/sessions",
            headers=headers,
            json={"agent": AGENT_ID, "environment_id": ENVIRONMENT_ID, "title": title}
        )
        session_id = resp.json()["id"]
        await client.post(
            f"https://api.anthropic.com/v1/sessions/{session_id}/events",
            headers=headers,
            json={"events": [{"type": "user.message", "content": [{"type": "text", "text": prompt}]}]}
        )
    print(f"[AGENT] Started session {session_id}: {title}")
    return session_id


# ── Webhook logic ─────────────────────────────────────────

async def handle_approve_reply(email_body: str, sender_email: str):
    """Handle APPROVE reply — create calendar invite.
    
    Expected format: APPROVE investor@email.com Thursday June 26 12pm
    The date/time after the email address becomes time_preferences_exact.
    """
    state = load_state()

    # Extract: APPROVE email@domain.com <everything after = the exact time>
    approve_match = re.search(
        r'APPROVE\s+([\w.+-]+@[\w.-]+)\s*(.+)?',
        email_body,
        re.IGNORECASE
    )
    investor_email = None
    approved_time  = None

    if approve_match:
        investor_email = approve_match.group(1).lower()
        approved_time  = approve_match.group(2).strip() if approve_match.group(2) else None
    else:
        for email, data in state.items():
            if data.get("status") == "awaiting_approval":
                investor_email = email
                break

    if not investor_email or investor_email not in state:
        print(f"[APPROVE] Could not find pending investor")
        return

    contact = state[investor_email]
    name = contact.get("name", investor_email)

    # Save approved time as time_preferences_exact
    if approved_time:
        contact["time_preferences_exact"] = [approved_time]
        print(f"[APPROVE] Exact time set: {approved_time}")

    time_prefs_exact    = contact.get("time_preferences_exact", [])
    time_prefs_estimate = contact.get("time_preferences_estimate", [])

    print(f"[APPROVE] Creating calendar invite for {investor_email}")

    prompt = f"""Create a calendar meeting invite for this approved investor meeting:

Investor: {name} ({investor_email})
Exact times requested: {', '.join(time_prefs_exact) if time_prefs_exact else 'None'}
Estimated times requested: {', '.join(time_prefs_estimate) if time_prefs_estimate else 'None'}
CKC Capital host: {FROM_EMAIL}

Instructions:
- Prefer exact times if available, otherwise use estimates
- Use check_availability tool to find a free 30-minute slot
- Create calendar event using create_calendar_event tool:
  Title: CKC Capital IR Meeting - {name}
  Attendees: {investor_email} and {FROM_EMAIL}
  Duration: 30 minutes
  Body: Thank you for your interest in CKC Capital. We look forward to discussing our credit and fixed income strategy.
- Send a confirmation email to {investor_email} letting them know the invite has been sent"""

    await run_agent_session(f"Calendar - {name}", prompt)

    state[investor_email]["status"] = "scheduled"
    save_state(state)


async def handle_investor_reply(email: dict, contact_name: str):
    """Process an incoming investor reply."""
    sender     = email.get("from", {}).get("emailAddress", {})
    sender_email = sender.get("address", "unknown").lower()
    subject    = email.get("subject", "")
    body_text  = email.get("body", {}).get("content", "")
    message_id = email.get("id", "")

    print(f"[WEBHOOK] Processing reply from {sender_email} ({contact_name})")

    prompt = f"""Process this investor reply to a CKC Capital IR email:

From: {contact_name} <{sender_email}>
Subject: {subject}
Body: {body_text[:1500]}

Instructions:
1. Classify the reply. Return JSON with these exact keys:
   - category: one of meeting_request | info_request | unsubscribe | positive_interest | escalate
   - sentiment: one of positive | neutral | negative
   - summary: one sentence
   - time_preferences_exact: array of specific times with day AND time e.g. ["Thursday June 26 at 10:00 AM"] — only if investor gave a specific time
   - time_preferences_estimate: array of general time ranges e.g. ["Thursday morning", "Friday afternoon"] — if investor gave a range but no specific time
   - requires_human: true if sensitive, legal, or compliance content detected

2. Based on classification:

   If category is NOT meeting_request:
   - Forward the email (message_id: {message_id}) to ops using forward_email tool

   If category is meeting_request AND both time_preferences_exact and time_preferences_estimate are EMPTY:
   - Send an email directly to {sender_email} asking what time works for them
   - Subject: Re: {subject}
   - Brief and professional, sign as The IR Team, CKC Capital
   - Do NOT forward to ops

   If category is meeting_request AND time_preferences_estimate is NOT EMPTY (range given, no exact time):
   - Forward email to ops using forward_email tool
   - Include in classification: "REPLY APPROVE {sender_email} to confirm and auto-create calendar invite"

   If category is meeting_request AND time_preferences_exact is NOT EMPTY (specific time given):
   - Forward email to ops using forward_email tool
   - Include in classification: "REPLY APPROVE {sender_email} to confirm and auto-create calendar invite"
   - Note the exact times clearly"""

    session_id = await run_agent_session(f"IR Reply - {contact_name}", prompt)

    # Save to state for approval tracking
    state = load_state()
    state[sender_email] = {
        "name": contact_name,
        "status": "awaiting_approval",
        "time_preferences_exact": [],
        "time_preferences_estimate": [],
        "message_id": message_id,
        "session_id": session_id
    }
    save_state(state)


async def webhook(request: Request):
    """Handle incoming Graph API webhook notifications."""
    validation_token = request.query_params.get("validationToken")
    if validation_token:
        print(f"[WEBHOOK] Validation handshake OK")
        return PlainTextResponse(validation_token, status_code=200)

    try:
        payload = await request.json()
        contacts = load_contacts()

        for notification in payload.get("value", []):
            resource = notification.get("resource", "")
            if "/Messages/" not in resource:
                continue

            message_id = resource.split("/Messages/")[-1]
            email = await graph(
                "GET",
                f"/users/{FROM_EMAIL}/messages/{message_id}"
                f"?$select=subject,from,body,receivedDateTime,id,toRecipients"
            )

            sender_email  = email.get("from", {}).get("emailAddress", {}).get("address", "").lower()
            subject       = email.get("subject", "")
            body_content  = email.get("body", {}).get("content", "")
            to_recipients = [
                r.get("emailAddress", {}).get("address", "").lower()
                for r in email.get("toRecipients", [])
            ]

            # Only process emails directly addressed to cokoro
            if FROM_EMAIL.lower() not in to_recipients:
                print(f"[WEBHOOK] Skipping - not directly addressed to {FROM_EMAIL}")
                continue

            # Skip classification emails to prevent loops
            if "[CKC IR CLASSIFICATION]" in body_content or "[CKC IR CLASSIFICATION]" in subject:
                print(f"[WEBHOOK] Skipping classification email")
                continue

            # Route: APPROVE reply → create calendar invite
            if "APPROVE" in body_content[:200].upper() and sender_email in [IR_EMAIL.lower(), FROM_EMAIL.lower()]:
                print(f"[WEBHOOK] APPROVE detected")
                await handle_approve_reply(body_content, sender_email)
                continue

            # Only process emails from contacts in contacts.csv
            if sender_email not in contacts:
                print(f"[WEBHOOK] Skipping - {sender_email} not in contacts list")
                continue

            contact_name = contacts[sender_email]
            await handle_investor_reply(email, contact_name)

    except Exception as e:
        print(f"[WEBHOOK] Error: {e}")

    return JSONResponse({"status": "ok"})


# ── Build combined app ────────────────────────────────────
sse_app = mcp.sse_app()
app = Starlette(routes=[
    Route("/webhook", endpoint=webhook, methods=["GET", "POST"]),
    Mount("/", app=sse_app)
])

if __name__ == "__main__":
    print(f"Starting CKC IR MCP Server on port {PORT}...")
    uvicorn.run(
        "ckc_mcp_server:app",
        host="0.0.0.0",
        port=PORT,
        proxy_headers=True,
        forwarded_allow_ips="*"
    )
