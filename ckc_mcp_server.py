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
from urllib.parse import quote

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

# ── State file ────────────────────────────────────────────
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
async def forward_email(message_id: str, classification: str, approval_link: str = "") -> str:
    """Forward a reply to ops@ckccapital.com with Claude classification and approval link."""
    comment = f"[CKC IR CLASSIFICATION]\n{classification}"
    if approval_link:
        comment += f"\n\n{approval_link}"
    comment += "\n\n[ORIGINAL MESSAGE]"
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


# ── Approval link generator ───────────────────────────────
def make_approval_link(investor_email: str, suggested_time: str) -> str:
    """Generate a mailto link that pre-fills an APPROVE email to cokoro."""
    subject = f"APPROVE {investor_email} {suggested_time}"
    encoded_subject = quote(subject)
    mailto = f"mailto:{FROM_EMAIL}?subject={encoded_subject}"
    return (
        f"► CONFIRM MEETING\n"
        f"Click the link below. Edit the meeting time in the subject if needed, then send.\n"
        f"{mailto}\n\n"
        f"Pre-filled subject: {subject}\n"
        f"(Change the date/time in the subject before sending if needed)"
    )


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


# ── Handle APPROVE email ──────────────────────────────────
async def handle_approve_reply(subject: str):
    """
    Handles APPROVE emails sent to cokoro@ckccapital.com.
    Expected subject: APPROVE investor@email.com Thursday June 26 10am
    """
    approve_match = re.search(
        r'APPROVE\s+([\w.+-]+@[\w.-]+)\s+(.+)',
        subject,
        re.IGNORECASE
    )
    if not approve_match:
        print(f"[APPROVE] Could not parse APPROVE subject: {subject}")
        return

    investor_email = approve_match.group(1).lower()
    exact_time     = approve_match.group(2).strip()

    print(f"[APPROVE] Confirmed: {investor_email} at {exact_time}")

    state = load_state()
    contact = state.get(investor_email, {})
    name    = contact.get("name", investor_email)

    # Save exact time to state
    contact["time_preferences_exact"] = [exact_time]
    contact["status"] = "approved"
    state[investor_email] = contact
    save_state(state)

    prompt = f"""Create a calendar meeting invite for this approved investor meeting:

Investor: {name} ({investor_email})
Confirmed time: {exact_time}
CKC Capital host: {FROM_EMAIL}

Instructions:
1. Use check_availability tool to verify the slot is free around {exact_time}
2. Create calendar event using create_calendar_event tool:
   - Title: CKC Capital IR Meeting - {name}
   - Attendees: {investor_email} and {FROM_EMAIL}
   - Duration: 30 minutes
   - Parse "{exact_time}" to ISO 8601 datetime for start/end
   - Body: Thank you for your interest in CKC Capital. We look forward to discussing our credit and fixed income strategy with you.
3. Send a confirmation email to {investor_email} letting them know the invite has been sent"""

    await run_agent_session(f"Calendar - {name}", prompt)

    state[investor_email]["status"] = "scheduled"
    save_state(state)


# ── Handle investor reply ─────────────────────────────────
async def handle_investor_reply(email: dict, contact_name: str):
    sender       = email.get("from", {}).get("emailAddress", {})
    sender_email = sender.get("address", "unknown").lower()
    subject      = email.get("subject", "")
    body_text    = email.get("body", {}).get("content", "")
    message_id   = email.get("id", "")

    print(f"[WEBHOOK] Processing reply from {sender_email} ({contact_name})")

    # Pre-generate approval link with a suggested time placeholder
    approval_link = make_approval_link(sender_email, "REPLACE_WITH_DATE_AND_TIME")

    prompt = (
        f"Process this investor reply to a CKC Capital IR email:\n\n"
        f"From: {contact_name} <{sender_email}>\n"
        f"Subject: {subject}\n"
        f"Body: {body_text[:1500]}\n\n"
        "Instructions:\n"
        "1. Classify the reply internally. Extract:\n"
        "   - category: meeting_request | info_request | unsubscribe | positive_interest | escalate\n"
        "   - sentiment: positive | neutral | negative\n"
        "   - summary: one sentence\n"
        "   - time_preferences_exact: specific times e.g. Thursday June 26 at 10:00 AM\n"
        "   - time_preferences_estimate: general ranges e.g. Thursday morning\n"
        "   - requires_human: true if sensitive/legal/compliance content\n\n"
        "2. Based on classification:\n\n"
        "   If NOT meeting_request:\n"
        f"   - Use forward_email tool. classification parameter:\n"
        f"     Hi Ops,\n\n\n"
        f"     [one sentence summary]\n\n"
        f"     Category: [category] | Sentiment: [sentiment] | Requires Human Review: [yes/no]\n\n"
        "   If meeting_request AND both time arrays EMPTY:\n"
        f"   - Use send_email to reply to {sender_email}\n"
        f"   - Subject: Re: {subject}\n"
        "   - Brief professional note asking what time works\n"
        "   - Sign as: The IR Team, CKC Capital\n"
        "   - Do NOT forward to ops\n\n"
        "   If meeting_request AND time preferences exist:\n"
        "   - Use forward_email tool. Set classification parameter to EXACTLY this format:\n\n"
        f"Hi Ops,\n\n\n"
        f"[Your one sentence summary of the investor reply]\n\n"
        f"------------------------------------------\n"
        f"CLASSIFICATION\n"
        f"------------------------------------------\n"
        f"Category: [category]\n"
        f"Sentiment: [sentiment]\n"
        f"Summary: [one sentence]\n"
        f"Requires Human Review: [Yes/No]\n\n"
        f"TIME PREFERENCES\n"
        f"- Exact times given: [list or None]\n"
        f"- Ranges given: [list or None]\n\n"
        f"------------------------------------------\n"
        f"APPROVAL LINK\n"
        f"------------------------------------------\n"
        f"Click the mailto link below. Edit the time in the subject if needed, then send.\n\n"
        f"mailto:{FROM_EMAIL}?subject=APPROVE {sender_email} [suggested date and time]\n\n"
        f"Pre-filled subject: APPROVE {sender_email} [suggested date and time]\n"
        f"(Edit the date/time in the subject before sending if needed)\n\n"
        f"------------------------------------------\n"
        f"ORIGINAL MESSAGE\n"
        f"------------------------------------------\n\n"
        "   - ONLY use mailto: links. No HTTP or HTTPS links.\n"
        "   - Do NOT output raw JSON anywhere.\n"
        "   - Replace [suggested date and time] with the best time from their preferences.\n"
        "   - If estimate only, suggest a specific date and time e.g. Thursday June 26 2026 10:00 AM"
    )

    session_id = await run_agent_session(f"IR Reply - {contact_name}", prompt)

    # Save to state
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


# ── Webhook ───────────────────────────────────────────────
async def webhook(request: Request):
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

            # Route: APPROVE email (subject starts with APPROVE)
            if subject.strip().upper().startswith("APPROVE"):
                print(f"[WEBHOOK] APPROVE detected: {subject}")
                await handle_approve_reply(subject)
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


# ── Build app ─────────────────────────────────────────────
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
