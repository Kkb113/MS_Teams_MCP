"""
Microsoft Teams MCP Server
---------------------------
A Model Context Protocol server that exposes Microsoft Teams + Calendar
operations so any LLM client can drive them via simple tool calls.

Auth: Device-code OAuth flow (opens browser, user signs in with Microsoft account).
Hosting: Designed for Render (uses SSE transport over HTTP).
"""

import os
import json
import logging
import asyncio
import webbrowser
from datetime import datetime, timezone
from typing import Any

import httpx
import msal
from mcp.server.fastmcp import FastMCP

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

CLIENT_ID = os.environ.get("MS_CLIENT_ID", "")
TENANT_ID = os.environ.get("MS_TENANT_ID", "common")
SCOPES = [
    "User.Read",
    "Chat.ReadWrite",
    "Channel.ReadBasic.All",
    "ChannelMessage.Send",
    "Team.ReadBasic.All",
    "Calendars.ReadWrite",
    "OnlineMeetings.ReadWrite",
]

GRAPH = "https://graph.microsoft.com/v1.0"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
TOKEN_CACHE_FILE = "token_cache.json"

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("teams-mcp")

# ---------------------------------------------------------------------------
# Token cache (persists across restarts)
# ---------------------------------------------------------------------------

_cache = msal.SerializableTokenCache()
if os.path.exists(TOKEN_CACHE_FILE):
    _cache.deserialize(open(TOKEN_CACHE_FILE).read())


def _save_cache():
    if _cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(_cache.serialize())


def _build_app() -> msal.PublicClientApplication:
    return msal.PublicClientApplication(
        CLIENT_ID, authority=AUTHORITY, token_cache=_cache
    )


async def _get_token() -> str:
    """Return a valid access token, refreshing silently or starting device-code flow."""
    app = _build_app()

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache()
            return result["access_token"]

    # Device-code flow — prints URL + code, opens browser
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Device flow failed: {json.dumps(flow, indent=2)}")

    log.info(flow["message"])
    # Try to open browser automatically
    try:
        webbrowser.open(flow["verification_uri"])
    except Exception:
        pass

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description', result)}")

    _save_cache()
    return result["access_token"]


# ---------------------------------------------------------------------------
# Graph helpers
# ---------------------------------------------------------------------------

async def _graph(method: str, path: str, body: dict | None = None, params: dict | None = None) -> dict:
    """Call Microsoft Graph and return the JSON response."""
    token = await _get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{GRAPH}{path}"

    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.request(method, url, headers=headers, json=body, params=params)

    if resp.status_code == 204:
        return {"status": "success"}
    if resp.status_code >= 400:
        return {"error": resp.status_code, "detail": resp.text}
    return resp.json()


def _pick(data: dict, keys: list[str]) -> dict:
    """Extract only the listed keys from a dict (keeps responses lean for LLMs)."""
    return {k: data[k] for k in keys if k in data}


# ---------------------------------------------------------------------------
# MCP Server
# ---------------------------------------------------------------------------

mcp = FastMCP("Microsoft Teams")

# ---- Profile ---------------------------------------------------------------

@mcp.tool()
async def get_my_profile() -> dict:
    """Get the signed-in user's profile (name, email, job title)."""
    data = await _graph("GET", "/me")
    if "error" in data:
        return data
    return _pick(data, ["displayName", "mail", "userPrincipalName", "jobTitle", "id"])


# ---- Teams & Channels ------------------------------------------------------

@mcp.tool()
async def list_teams() -> list[dict]:
    """List all Teams the user has joined."""
    data = await _graph("GET", "/me/joinedTeams")
    if "error" in data:
        return data
    return [_pick(t, ["id", "displayName", "description"]) for t in data.get("value", [])]


@mcp.tool()
async def list_channels(team_id: str) -> list[dict]:
    """List channels in a Team.

    Args:
        team_id: The Team's ID (from list_teams).
    """
    data = await _graph("GET", f"/teams/{team_id}/channels")
    if "error" in data:
        return data
    return [_pick(c, ["id", "displayName", "description"]) for c in data.get("value", [])]


@mcp.tool()
async def send_channel_message(team_id: str, channel_id: str, message: str) -> dict:
    """Send a message to a Teams channel.

    Args:
        team_id: The Team's ID.
        channel_id: The channel's ID (from list_channels).
        message: Plain-text message content.
    """
    body = {"body": {"contentType": "text", "content": message}}
    return await _graph("POST", f"/teams/{team_id}/channels/{channel_id}/messages", body)


@mcp.tool()
async def list_channel_messages(team_id: str, channel_id: str, top: int = 10) -> list[dict]:
    """Get recent messages from a channel.

    Args:
        team_id: The Team's ID.
        channel_id: The channel's ID.
        top: Number of messages to fetch (default 10, max 50).
    """
    data = await _graph("GET", f"/teams/{team_id}/channels/{channel_id}/messages", params={"$top": min(top, 50)})
    if "error" in data:
        return data
    results = []
    for m in data.get("value", []):
        results.append({
            "id": m.get("id"),
            "from": m.get("from", {}).get("user", {}).get("displayName"),
            "body": m.get("body", {}).get("content", "")[:500],
            "createdDateTime": m.get("createdDateTime"),
        })
    return results


# ---- Chats -----------------------------------------------------------------

@mcp.tool()
async def list_chats(top: int = 20) -> list[dict]:
    """List the user's recent chats.

    Args:
        top: Number of chats to return (default 20).
    """
    data = await _graph("GET", "/me/chats", params={"$top": min(top, 50), "$expand": "members"})
    if "error" in data:
        return data
    results = []
    for c in data.get("value", []):
        members = [m.get("displayName") for m in c.get("members", [])]
        results.append({
            "id": c.get("id"),
            "topic": c.get("topic"),
            "chatType": c.get("chatType"),
            "members": members,
        })
    return results


@mcp.tool()
async def send_chat_message(chat_id: str, message: str) -> dict:
    """Send a message in an existing chat.

    Args:
        chat_id: The chat's ID (from list_chats).
        message: Plain-text message content.
    """
    body = {"body": {"contentType": "text", "content": message}}
    return await _graph("POST", f"/chats/{chat_id}/messages", body)


@mcp.tool()
async def list_chat_messages(chat_id: str, top: int = 15) -> list[dict]:
    """Get recent messages from a chat.

    Args:
        chat_id: The chat's ID.
        top: Number of messages (default 15).
    """
    data = await _graph("GET", f"/chats/{chat_id}/messages", params={"$top": min(top, 50)})
    if "error" in data:
        return data
    results = []
    for m in data.get("value", []):
        results.append({
            "id": m.get("id"),
            "from": m.get("from", {}).get("user", {}).get("displayName"),
            "body": m.get("body", {}).get("content", "")[:500],
            "createdDateTime": m.get("createdDateTime"),
        })
    return results


@mcp.tool()
async def create_chat(member_emails: list[str], message: str | None = None) -> dict:
    """Start a new 1:1 or group chat.

    Args:
        member_emails: List of user email addresses to include (plus yourself).
        message: Optional first message to send in the chat.
    """
    # Build members list — the signed-in user is auto-included by Graph
    me = await _graph("GET", "/me")
    my_id = me.get("id")

    members = [
        {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": ["owner"],
            "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{my_id}')",
        }
    ]
    for email in member_emails:
        members.append({
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": ["owner"],
            "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{email}')",
        })

    chat_type = "oneOnOne" if len(member_emails) == 1 else "group"
    body: dict[str, Any] = {"chatType": chat_type, "members": members}

    result = await _graph("POST", "/chats", body)
    if "error" in result:
        return result

    # Optionally send the first message
    if message and result.get("id"):
        await send_chat_message(result["id"], message)

    return _pick(result, ["id", "chatType", "createdDateTime"])


# ---- Calendar --------------------------------------------------------------

@mcp.tool()
async def list_events(days: int = 7, top: int = 20) -> list[dict]:
    """List upcoming calendar events.

    Args:
        days: Look-ahead window in days (default 7).
        top: Max events to return (default 20).
    """
    now = datetime.now(timezone.utc)
    start = now.isoformat()
    end = now.replace(day=now.day + days).isoformat() if days <= 28 else now.isoformat()

    data = await _graph(
        "GET",
        "/me/calendarView",
        params={"startDateTime": start, "endDateTime": end, "$top": min(top, 50), "$orderby": "start/dateTime"},
    )
    if "error" in data:
        return data
    results = []
    for e in data.get("value", []):
        results.append({
            "id": e.get("id"),
            "subject": e.get("subject"),
            "start": e.get("start"),
            "end": e.get("end"),
            "location": e.get("location", {}).get("displayName"),
            "isOnlineMeeting": e.get("isOnlineMeeting"),
            "organizer": e.get("organizer", {}).get("emailAddress", {}).get("name"),
            "joinUrl": e.get("onlineMeeting", {}).get("joinUrl"),
        })
    return results


@mcp.tool()
async def create_event(
    subject: str,
    start: str,
    end: str,
    attendees: list[str] | None = None,
    body_text: str = "",
    is_online: bool = False,
    location: str = "",
) -> dict:
    """Create a calendar event.

    Args:
        subject: Event title.
        start: ISO 8601 datetime string (e.g. '2025-06-15T10:00:00').
        end: ISO 8601 datetime string.
        attendees: List of attendee email addresses.
        body_text: Optional event description.
        is_online: If True, creates a Teams meeting link.
        location: Optional physical location name.
    """
    event: dict[str, Any] = {
        "subject": subject,
        "start": {"dateTime": start, "timeZone": "UTC"},
        "end": {"dateTime": end, "timeZone": "UTC"},
        "isOnlineMeeting": is_online,
    }
    if is_online:
        event["onlineMeetingProvider"] = "teamsForBusiness"
    if body_text:
        event["body"] = {"contentType": "text", "content": body_text}
    if location:
        event["location"] = {"displayName": location}
    if attendees:
        event["attendees"] = [
            {"emailAddress": {"address": e}, "type": "required"} for e in attendees
        ]

    result = await _graph("POST", "/me/events", event)
    if "error" in result:
        return result
    return _pick(result, ["id", "subject", "start", "end", "webLink", "onlineMeeting"])


@mcp.tool()
async def update_event(
    event_id: str,
    subject: str | None = None,
    start: str | None = None,
    end: str | None = None,
    body_text: str | None = None,
    location: str | None = None,
) -> dict:
    """Update an existing calendar event.

    Args:
        event_id: The event ID (from list_events).
        subject: New title (optional).
        start: New start time as ISO 8601 (optional).
        end: New end time as ISO 8601 (optional).
        body_text: New description (optional).
        location: New location (optional).
    """
    patch: dict[str, Any] = {}
    if subject:
        patch["subject"] = subject
    if start:
        patch["start"] = {"dateTime": start, "timeZone": "UTC"}
    if end:
        patch["end"] = {"dateTime": end, "timeZone": "UTC"}
    if body_text is not None:
        patch["body"] = {"contentType": "text", "content": body_text}
    if location is not None:
        patch["location"] = {"displayName": location}

    return await _graph("PATCH", f"/me/events/{event_id}", patch)


@mcp.tool()
async def delete_event(event_id: str) -> dict:
    """Delete a calendar event.

    Args:
        event_id: The event ID to delete.
    """
    return await _graph("DELETE", f"/me/events/{event_id}")


@mcp.tool()
async def respond_to_event(event_id: str, response: str, message: str = "") -> dict:
    """Accept, tentatively accept, or decline a calendar event.

    Args:
        event_id: The event ID.
        response: One of 'accept', 'tentativelyAccept', or 'decline'.
        message: Optional response message.
    """
    if response not in ("accept", "tentativelyAccept", "decline"):
        return {"error": "response must be accept, tentativelyAccept, or decline"}

    body: dict[str, Any] = {"sendResponse": True}
    if message:
        body["comment"] = message

    return await _graph("POST", f"/me/events/{event_id}/{response}", body)


# ---- Online Meetings -------------------------------------------------------

@mcp.tool()
async def create_meeting(
    subject: str,
    start: str,
    end: str,
) -> dict:
    """Create a Teams online meeting and get the join link.

    Args:
        subject: Meeting title.
        start: ISO 8601 start datetime.
        end: ISO 8601 end datetime.
    """
    body = {
        "subject": subject,
        "startDateTime": start,
        "endDateTime": end,
    }
    result = await _graph("POST", "/me/onlineMeetings", body)
    if "error" in result:
        return result
    return _pick(result, ["id", "subject", "joinWebUrl", "startDateTime", "endDateTime"])


# ---- Status / Presence -----------------------------------------------------

@mcp.tool()
async def get_my_presence() -> dict:
    """Get the signed-in user's current presence/availability status."""
    return await _graph("GET", "/me/presence")


@mcp.tool()
async def set_status_message(message: str) -> dict:
    """Set a status message on your Teams profile.

    Args:
        message: The status message text.
    """
    body = {
        "statusMessage": {
            "message": {
                "content": message,
                "contentType": "text",
            }
        }
    }
    return await _graph("PATCH", f"/me/presence/setStatusMessage", body)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import uvicorn
    from starlette.middleware import Middleware
    from starlette.middleware.trustedhost import TrustedHostMiddleware

    transport = os.environ.get("MCP_TRANSPORT", "sse")
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "8000"))

    if transport == "stdio":
        mcp.run(transport="stdio")
    else:
        app = mcp.sse_app()
        # Allow all hosts — required when running behind Render's reverse proxy
        app.add_middleware(TrustedHostMiddleware, allowed_hosts=["*"])
        uvicorn.run(app, host=host, port=port, proxy_headers=True, forwarded_allow_ips="*")
