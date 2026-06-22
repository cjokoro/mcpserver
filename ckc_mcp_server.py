import httpx
import os
import uvicorn
from mcp.server.fastmcp import FastMCP
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request

# ── Allow all host headers ────────────────────────────────
os.environ["MCP_ALLOW_ALL_HOSTS"] = "1"

# ── Config — read from environment variables ──────────────
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
GRAPH_BASE    = "https://graph.microsoft.com/v1.0"
IR_EMAIL      = os.environ.get("IR_EMAIL", "ops@ckccapital.com")
FROM_EMAIL    = os.environ.get("FROM_EMAIL", "cokoro@ckccapital.com")

# ── FastMCP server ────────────────────────────────────────
mcp = FastMCP("CKC IR Server")

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
            method,
            f"{GRAPH_BASE}{path}",
            headers=headers,
            **kwargs
        )
        resp.raise_for_status()
        return resp.json() if resp.content else {}

# ── Tools ─────────────────────────────────────────────────

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
    title: str,
    attendees: list[str],
    start: str,
    end: str,
    body: str = ""
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


# ── Middleware to strip host validation ───────────────────
class AllowAllHostsMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        # Override host header to bypass MCP host validation
        request.scope["headers"] = [
            (k, v) for k, v in request.scope["headers"]
            if k != b"host"
        ] + [(b"host", b"localhost")]
        return await call_next(request)


# ── Run ───────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    print(f"Starting CKC IR MCP Server on port {port}...")
    app = mcp.sse_app()
    from starlette.applications import Starlette
    from starlette.routing import Mount
    wrapped = Starlette(routes=[Mount("/", app=app)])
    wrapped.add_middleware(AllowAllHostsMiddleware)
    uvicorn.run(
        wrapped,
        host="0.0.0.0",
        port=port,
        proxy_headers=True,
        forwarded_allow_ips="*"
    )
