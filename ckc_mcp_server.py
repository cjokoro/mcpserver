"""
Register a Microsoft Graph webhook subscription to watch for new emails.
Run once to set up — subscription lasts 3 days, then needs renewal.
Usage: python register_webhook.py
"""

import httpx
import os
from datetime import datetime, timezone, timedelta

CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
FROM_EMAIL    = os.environ.get("FROM_EMAIL", "cokoro@ckccapital.com")
WEBHOOK_URL   = "https://mcpserver-kgkf.onrender.com/webhook"


def get_token() -> str:
    resp = httpx.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={
            "client_id":     CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope":         "https://graph.microsoft.com/.default",
            "grant_type":    "client_credentials"
        }
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def register_subscription(token: str) -> dict:
    expiry = (datetime.now(timezone.utc) + timedelta(days=3)).strftime("%Y-%m-%dT%H:%M:%SZ")
    resp = httpx.post(
        "https://graph.microsoft.com/v1.0/subscriptions",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={
            "changeType":         "created",
            "notificationUrl":    WEBHOOK_URL,
            "resource":           f"/users/{FROM_EMAIL}/mailFolders/inbox/messages",
            "expirationDateTime": expiry,
            "clientState":        "ckc-ir-webhook-secret"
        }
    )
    resp.raise_for_status()
    return resp.json()


def list_subscriptions(token: str) -> list:
    resp = httpx.get(
        "https://graph.microsoft.com/v1.0/subscriptions",
        headers={"Authorization": f"Bearer {token}"}
    )
    resp.raise_for_status()
    return resp.json().get("value", [])


if __name__ == "__main__":
    print("Registering Graph API webhook subscription...")
    token = get_token()
    print("Token acquired")

    # Check existing subscriptions
    existing = list_subscriptions(token)
    if existing:
        print(f"Found {len(existing)} existing subscription(s):")
        for s in existing:
            print(f"  - {s['id']} expires {s['expirationDateTime']}")

    # Register new subscription
    sub = register_subscription(token)
    print(f"\n✓ Subscription registered:")
    print(f"  ID:      {sub['id']}")
    print(f"  Resource: {sub['resource']}")
    print(f"  Expires: {sub['expirationDateTime']}")
    print(f"\nWebhook URL: {WEBHOOK_URL}")
    print(f"\nNote: Subscription expires in 3 days. Re-run this script to renew.")
