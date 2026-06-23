import httpx
import os
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
async def flag_email(message_id: str) -> str:
    """Flag an email as replied/processed in Outlook."""
    await graph(
        "PATCH",
        f"/users/{FROM_EMAIL}/messages/{message_id}",
        json={"flag": {"flagStatus": "flagged"}}
    )
    return "Email flagged"


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


# ── Webhook handler ───────────────────────────────────────

async def trigger_agent(email: dict):
    sender     = email.get("from", {}).get("emailAddress", {})
    sender_email = sender.get("address", "unknown").lower()
    sender_name  = sender.get("name", "unknown")
    subject    = email.get("subject", "")
    body       = email.get("body", {}).get("content", "")
    message_id = email.get("id", "")

    prompt = f"""A new email reply has arrived in the CKC Capital IR inbox. Process it:

From: {sender_name} <{sender_email}>
Subject: {subject}
Message ID: {message_id}
Body:
{body[:2000]}

Please:
1. Classify this reply (return JSON with category, sentiment, summary, time_preferences, requires_human)
2. Flag the email using flag_email tool
3. If category is meeting_request:
   - Extract the time preferences from the reply
   - Check availability using check_availability tool for the requested times
   - Create a calendar event using create_calendar_event tool
   - Attendees should include {sender_email} and {FROM_EMAIL}
   - Title should be: CKC Capital IR Meeting - {sender_name}
   - Duration: 30 minutes
   - Body: Follow-up meeting to discuss CKC Capital investment strategy

Note: Forwarding to ops is disabled during testing."""

    headers = {
        "x-api-key":         ANTHROPIC_KEY,
        "anthropic-version": "2023-06-01",
        "anthropic-beta":    "managed-agents-2026-04-01",
        "content-type":      "application/json"
    }
    async with httpx.AsyncClient(timeout=30.0) as client:
        resp = await client.post(
            "https://api.anthropic.com/v1/sessions",
            headers=headers,
            json={"agent": AGENT_ID, "environment_id": ENVIRONMENT_ID,
                  "title": f"IR Reply - {sender_name}"}
        )
        session_id = resp.json()["id"]
        await client.post(
            f"https://api.anthropic.com/v1/sessions/{session_id}/events",
            headers=headers,
            json={"events": [{"type": "user.message",
                               "content": [{"type": "text", "text": prompt}]}]}
        )
    print(f"[WEBHOOK] Started session {session_id} for reply from {sender_email}")


async def webhook(request: Request):
    # Graph validation handshake
    validation_token = request.query_params.get("validationToken")
    if validation_token:
        print(f"[WEBHOOK] Validation handshake OK")
        return PlainTextResponse(validation_token, status_code=200)

    # Process notification
    try:
        body = await request.json()
        for notification in body.get("value", []):
            resource = notification.get("resource", "")
            if "/Messages/" in resource:
                message_id = resource.split("/Messages/")[-1]
                email = await graph(
                    "GET",
                    f"/users/{FROM_EMAIL}/messages/{message_id}"
                    f"?$select=subject,from,body,receivedDateTime,id"
                )
                await trigger_agent(email)
    except Exception as e:
        print(f"[WEBHOOK] Error: {e}")

    return JSONResponse({"status": "ok"})


# ── Build combined Starlette app ──────────────────────────
sse_app = mcp.sse_app()
app = Starlette(routes=[
    Route("/webhook", endpoint=webhook, methods=["GET", "POST"]),
    Mount("/", app=sse_app)
])

# ── Run ───────────────────────────────────────────────────
if __name__ == "__main__":
    print(f"Starting CKC IR MCP Server on port {PORT}...")
    uvicorn.run(
        "ckc_mcp_server:app",
        host="0.0.0.0",
        port=PORT,
        proxy_headers=True,
        forwarded_allow_ips="*"
    )
