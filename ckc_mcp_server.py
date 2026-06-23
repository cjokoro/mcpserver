import httpx
import os
import json
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
    """Handle APPROVE reply from ops — create calendar invite."""
    state = load_state()

    # Extract investor email from APPROVE message
    # Expected format: "APPROVE investor@email.com" or just "APPROVE" with state lookup
    approve_match = re.search(r'APPROVE\s+([\w.+-]+@[\w.-]+)', email_body, re.IGNORECASE)

    investor_email = None
    if approve_match:
        investor_email = approve_match.group(1).lower()
    else:
        # Find first pending contact in state
        for email, data in state.items():
            if data.get("status") == "awaiting_approval":
                investor_email = email
                break

    if not investor_email or investor_email not in state:
        print(f"[APPROVE] Could not find pending investor for approval")
        return

    contact = state[investor_email]
    time_prefs = contact.get("time_preferences", [])
    name = contact.get("name", investor_email)

    print(f"[APPROVE] Creating calendar invite for {investor_email} - times: {time_prefs}")

    prompt = f"""Create a calendar meeting invite for this approved investor meeting:

Investor: {name} ({investor_email})
Requested times: {', '.join(time_prefs)}
CKC Capital host: {FROM_EMAIL}

Please:
1. Use check_availability tool to find a free 30-minute slot matching their preferences
2. Create a calendar event using create_calendar_event tool
   - Title: CKC Capital IR Meeting - {name}
   - Attendees: {investor_email} and {FROM_EMAIL}
   - Duration: 30 minutes
   - Body: Thank you for your interest in CKC Capital. We look forward to discussing our credit and fixed income strategy with you.
3. Send a confirmation email to {investor_email} letting them know the invite has been sent"""

    await run_agent_session(f"Calendar - {name}", prompt)

    # Update state
    state[investor_email]["status"] = "scheduled"
    save_state(state)


async def handle_investor_reply(email: dict):
    """Process an incoming investor reply."""
    sender     = email.get("from", {}).get("emailAddress", {})
    sender_email = sender.get("address", "unknown").lower()
    sender_name  = sender.get("name", sender_email)
    subject    = email.get("subject", "")
    body_text  = email.get("body", {}).get("content", "")
    message_id = email.get("id", "")

    # Skip ops replies that are APPROVE commands — handled separately
    if sender_email == IR_EMAIL.lower() or "APPROVE" in body_text[:50].upper():
        return

    print(f"[WEBHOOK] Processing reply from {sender_email}")

    prompt = f"""Process this investor reply to a CKC Capital IR email:

From: {sender_name} <{sender_email}>
Subject: {subject}
Body: {body_text[:1500]}

Instructions:
1. Classify the reply. Return JSON with: category, sentiment, summary, time_preferences, requires_human

2. Based on classification:

   If category is NOT meeting_request:
   - Forward the email (message_id: {message_id}) to ops using forward_email tool

   If category is meeting_request AND time_preferences is EMPTY (investor wants to meet but gave no specific time):
   - Send an email to {sender_email} asking what time works for them
   - Subject: Re: {subject}
   - Keep it brief and professional, sign as The IR Team, CKC Capital
   - Do NOT forward to ops yet

   If category is meeting_request AND time_preferences is NOT EMPTY (investor gave specific times):
   - Forward the email to ops using forward_email tool
   - Include in the classification comment: "REPLY APPROVE {sender_email} to confirm and auto-create calendar invite"
   - Save the contact info for scheduling"""

    session_id = await run_agent_session(f"IR Reply - {sender_name}", prompt)

    # For meeting requests with times — save to state for approval
    # We'll update state after the agent runs (simplified — agent handles the logic)
    state = load_state()
    state[sender_email] = {
        "name": sender_name,
        "status": "awaiting_approval",
        "time_preferences": [],  # agent will extract these
        "message_id": message_id,
        "session_id": session_id
    }
    save_state(state)


async def webhook(request: Request):
    """Handle incoming Graph API webhook notifications."""
    # Graph validation handshake
    validation_token = request.query_params.get("validationToken")
    if validation_token:
        print(f"[WEBHOOK] Validation handshake OK")
        return PlainTextResponse(validation_token, status_code=200)

    try:
        payload = await request.json()
        for notification in payload.get("value", []):
            resource = notification.get("resource", "")
            if "/Messages/" not in resource:
                continue

            message_id = resource.split("/Messages/")[-1]
            email = await graph(
                "GET",
                f"/users/{FROM_EMAIL}/messages/{message_id}"
                f"?$select=subject,from,body,receivedDateTime,id"
            )

            sender_email = email.get("from", {}).get("emailAddress", {}).get("address", "").lower()
            body_preview = email.get("body", {}).get("content", "")[:100].upper()

            # Route: APPROVE reply from ops → create calendar invite
            if sender_email == IR_EMAIL.lower() or (
                sender_email == FROM_EMAIL.lower() and "APPROVE" in body_preview
            ):
                print(f"[WEBHOOK] APPROVE detected from {sender_email}")
                await handle_approve_reply(
                    email.get("body", {}).get("content", ""),
                    sender_email
                )
            else:
                # Route: investor reply → classify and handle
                await handle_investor_reply(email)

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
