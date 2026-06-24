import httpx
import os
import json
import csv
import re
import uvicorn
from mcp.server.fastmcp import FastMCP
from mcp.server.transport_security import TransportSecuritySettings
from starlette.requests import Request
from starlette.responses import PlainTextResponse, JSONResponse, HTMLResponse
from starlette.routing import Route, Mount
from starlette.applications import Starlette
from urllib.parse import quote, urlencode

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
SERVER_URL     = os.environ.get("SERVER_URL", "https://mcpserver-kgkf.onrender.com")

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
async def forward_email(message_id: str, classification: str) -> str:
    """Forward a reply to ops@ckccapital.com with classification and approval link."""
    await graph(
        "POST",
        f"/users/{FROM_EMAIL}/messages/{message_id}/forward",
        json={
            "comment": classification,
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


# ── Approval URL builder ──────────────────────────────────
def make_approval_url(investor_email: str, investor_name: str,
                      time_exact: list, time_estimate: list) -> str:
    """Build a real URL to the approval web page on this server."""
    # Pick best suggested datetime for pre-fill
    suggested = ""
    if time_exact:
        suggested = time_exact[0]
    elif time_estimate:
        # Convert estimate to a suggested specific time
        # Agent will handle this in the prompt
        suggested = time_estimate[0]

    params = {
        "investor": investor_email,
        "name":     investor_name,
        "suggested": suggested,
        "estimate":  " | ".join(time_estimate) if time_estimate else "",
    }
    return f"{SERVER_URL}/approve?{urlencode(params)}"


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


# ── Approval web page ─────────────────────────────────────
async def approve_page(request: Request):
    """Serve the meeting approval form."""
    investor = request.query_params.get("investor", "")
    name     = request.query_params.get("name", "Investor")
    suggested = request.query_params.get("suggested", "")
    estimate  = request.query_params.get("estimate", "")

    # Parse suggested time for date/time input pre-fill
    # Default to next business day 10am if not parseable
    suggested_date = ""
    suggested_time = "10:00"

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CKC Capital — Meeting Approval</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
         background: #f5f5f5; min-height: 100vh; display: flex;
         align-items: center; justify-content: center; padding: 20px; }}
  .card {{ background: white; border-radius: 12px; padding: 32px;
           max-width: 480px; width: 100%; box-shadow: 0 2px 16px rgba(0,0,0,0.1); }}
  .logo {{ font-size: 13px; font-weight: 600; color: #666; letter-spacing: 0.05em;
           text-transform: uppercase; margin-bottom: 24px; }}
  h1 {{ font-size: 20px; font-weight: 600; margin-bottom: 6px; color: #111; }}
  .subtitle {{ font-size: 14px; color: #666; margin-bottom: 24px; }}
  .investor-box {{ background: #f8f8f8; border-radius: 8px; padding: 14px 16px;
                   margin-bottom: 24px; }}
  .investor-box .label {{ font-size: 11px; color: #999; text-transform: uppercase;
                          letter-spacing: 0.05em; margin-bottom: 4px; }}
  .investor-box .name {{ font-size: 15px; font-weight: 500; color: #111; }}
  .investor-box .email {{ font-size: 13px; color: #666; }}
  .requested {{ font-size: 13px; color: #555; margin-top: 8px; }}
  .field {{ margin-bottom: 16px; }}
  label {{ display: block; font-size: 13px; font-weight: 500; color: #333;
           margin-bottom: 6px; }}
  input, select {{ width: 100%; padding: 10px 12px; border: 1px solid #ddd;
                   border-radius: 8px; font-size: 14px; color: #111;
                   outline: none; transition: border-color 0.15s; }}
  input:focus, select:focus {{ border-color: #000; }}
  .row {{ display: flex; gap: 12px; }}
  .row .field {{ flex: 1; }}
  button {{ width: 100%; padding: 12px; background: #111; color: white;
            border: none; border-radius: 8px; font-size: 15px; font-weight: 500;
            cursor: pointer; margin-top: 8px; transition: opacity 0.15s; }}
  button:hover {{ opacity: 0.85; }}
  .divider {{ height: 1px; background: #eee; margin: 20px 0; }}
  .success {{ display: none; text-align: center; padding: 24px 0; }}
  .success .icon {{ font-size: 40px; margin-bottom: 12px; }}
  .success h2 {{ font-size: 18px; font-weight: 600; margin-bottom: 8px; }}
  .success p {{ font-size: 14px; color: #666; }}
</style>
</head>
<body>
<div class="card">
  <div class="logo">CKC Capital IR</div>
  <h1>Meeting Approval</h1>
  <p class="subtitle">Review and confirm the meeting details below.</p>

  <div class="investor-box">
    <div class="label">Investor</div>
    <div class="name">{name}</div>
    <div class="email">{investor}</div>
    {f'<div class="requested">Requested: {estimate}</div>' if estimate else ''}
  </div>

  <form id="approvalForm">
    <div class="row">
      <div class="field">
        <label>Meeting Date</label>
        <input type="date" id="meetingDate" name="date" required
               value="{suggested_date}" />
      </div>
      <div class="field">
        <label>Meeting Time</label>
        <input type="time" id="meetingTime" name="time" required
               value="{suggested_time}" />
      </div>
    </div>
    <div class="field">
      <label>Duration</label>
      <select id="duration" name="duration">
        <option value="30">30 minutes</option>
        <option value="45">45 minutes</option>
        <option value="60">60 minutes</option>
      </select>
    </div>
    <div class="field">
      <label>Meeting Title</label>
      <input type="text" id="title" name="title"
             value="CKC Capital IR Meeting - {name}" />
    </div>
    <div class="divider"></div>
    <button type="submit">Confirm &amp; Send Calendar Invite</button>
  </form>

  <div class="success" id="successMsg">
    <div class="icon">✓</div>
    <h2>Invite Sent</h2>
    <p>Calendar invite sent to {name} at {investor}.</p>
  </div>
</div>

<script>
document.getElementById('approvalForm').addEventListener('submit', async function(e) {{
  e.preventDefault();
  const btn = this.querySelector('button');
  btn.textContent = 'Sending...';
  btn.disabled = true;

  const payload = {{
    investor_email: '{investor}',
    investor_name:  '{name}',
    date:     document.getElementById('meetingDate').value,
    time:     document.getElementById('meetingTime').value,
    duration: document.getElementById('duration').value,
    title:    document.getElementById('title').value
  }};

  try {{
    const resp = await fetch('/confirm', {{
      method: 'POST',
      headers: {{'Content-Type': 'application/json'}},
      body: JSON.stringify(payload)
    }});
    const data = await resp.json();
    if (data.status === 'ok') {{
      document.getElementById('approvalForm').style.display = 'none';
      document.getElementById('successMsg').style.display = 'block';
    }} else {{
      btn.textContent = 'Error — Try Again';
      btn.disabled = false;
    }}
  }} catch(err) {{
    btn.textContent = 'Error — Try Again';
    btn.disabled = false;
  }}
}});
</script>
</body>
</html>"""
    return HTMLResponse(html)


# ── Confirm endpoint ──────────────────────────────────────
async def confirm_meeting(request: Request):
    """Handle form submission — create calendar invite via agent."""
    try:
        data = await request.json()
        investor_email = data.get("investor_email", "")
        investor_name  = data.get("investor_name", "Investor")
        date           = data.get("date", "")       # YYYY-MM-DD
        time           = data.get("time", "10:00")  # HH:MM
        duration       = int(data.get("duration", 30))
        title          = data.get("title", f"CKC Capital IR Meeting - {investor_name}")

        # Build ISO 8601 start/end
        start_dt = f"{date}T{time}:00"
        # Calculate end time
        h, m = map(int, time.split(":"))
        total_minutes = h * 60 + m + duration
        end_h = total_minutes // 60
        end_m = total_minutes % 60
        end_dt = f"{date}T{end_h:02d}:{end_m:02d}:00"

        print(f"[CONFIRM] Creating invite for {investor_email} at {start_dt}")

        prompt = (
            f"Create a calendar meeting invite:\n\n"
            f"Investor: {investor_name} ({investor_email})\n"
            f"Title: {title}\n"
            f"Start: {start_dt} Eastern Time\n"
            f"End: {end_dt} Eastern Time\n"
            f"Attendees: {investor_email} and {FROM_EMAIL}\n\n"
            f"Instructions:\n"
            f"1. Use create_calendar_event tool with these exact details\n"
            f"2. Send a confirmation email to {investor_email}:\n"
            f"   Subject: Meeting Confirmed - {title}\n"
            f"   Body: Thank you for your interest in CKC Capital. "
            f"A calendar invite has been sent for {date} at {time} Eastern Time. "
            f"We look forward to speaking with you.\n"
            f"   Sign as: The IR Team, CKC Capital"
        )

        await run_agent_session(f"Calendar - {investor_name}", prompt)

        # Update state
        state = load_state()
        if investor_email in state:
            state[investor_email]["status"] = "scheduled"
            state[investor_email]["time_preferences_exact"] = [f"{date} {time}"]
            save_state(state)

        return JSONResponse({"status": "ok"})

    except Exception as e:
        print(f"[CONFIRM] Error: {e}")
        return JSONResponse({"status": "error", "message": str(e)}, status_code=500)


# ── Webhook ───────────────────────────────────────────────
async def handle_investor_reply(email: dict, contact_name: str):
    sender_email = email.get("from", {}).get("emailAddress", {}).get("address", "").lower()
    subject      = email.get("subject", "")
    body_text    = email.get("body", {}).get("content", "")
    message_id   = email.get("id", "")

    print(f"[WEBHOOK] Processing reply from {sender_email} ({contact_name})")

    # Build approval URL — agent will fill in the time preferences
    approval_url_base = f"{SERVER_URL}/approve?investor={quote(sender_email)}&name={quote(contact_name)}"

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
        "   - Use forward_email tool with classification:\n"
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
        "   - Build the approval URL by appending suggested and estimate params:\n"
        f"     Base URL: {approval_url_base}\n"
        "     Add: &suggested=[suggested ISO date e.g. 2026-06-26T10:00]\n"
        "     Add: &estimate=[their time ranges joined by pipe e.g. Thursday morning|Friday afternoon]\n"
        "   - Use forward_email tool with classification formatted EXACTLY like this:\n\n"
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
        f"Click the link below to open the meeting approval form:\n\n"
        f"[FULL APPROVAL URL HERE]\n\n"
        f"------------------------------------------\n"
        f"ORIGINAL MESSAGE\n"
        f"------------------------------------------\n\n"
        "   - Replace [FULL APPROVAL URL HERE] with the complete URL you built above.\n"
        "   - Do NOT use mailto links. Use only the https URL.\n"
        "   - Do NOT output raw JSON anywhere."
    )

    session_id = await run_agent_session(f"IR Reply - {contact_name}", prompt)

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
    Route("/approve", endpoint=approve_page, methods=["GET"]),
    Route("/confirm", endpoint=confirm_meeting, methods=["POST"]),
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
