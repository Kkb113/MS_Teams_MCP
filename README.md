# Microsoft Teams MCP Server

A lightweight MCP (Model Context Protocol) server that lets any LLM manage Microsoft Teams — chats, channels, calendar events, meetings, and presence — via simple tool calls.

## What It Does

| Category | Tools |
|----------|-------|
| **Profile** | `get_my_profile` |
| **Teams & Channels** | `list_teams`, `list_channels`, `send_channel_message`, `list_channel_messages` |
| **Chats** | `list_chats`, `send_chat_message`, `list_chat_messages`, `create_chat` |
| **Calendar** | `list_events`, `create_event`, `update_event`, `delete_event`, `respond_to_event` |
| **Meetings** | `create_meeting` |
| **Presence** | `get_my_presence`, `set_status_message` |

---

## 1. Azure App Registration (One-Time Setup)

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Click **New registration**
   - **Name**: `Teams MCP`
   - **Supported account types**: Accounts in any org directory + personal Microsoft accounts
   - **Redirect URI**: Leave blank (device code flow doesn't need one)
3. Click **Register**
4. Copy the **Application (client) ID** — this is your `MS_CLIENT_ID`
5. Copy the **Directory (tenant) ID** — this is your `MS_TENANT_ID` (or use `common` for multi-tenant)
6. Go to **Authentication** → toggle **Allow public client flows** to **Yes** → Save
7. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** and add:
   - `User.Read`
   - `Chat.ReadWrite`
   - `Channel.ReadBasic.All`
   - `ChannelMessage.Send`
   - `Team.ReadBasic.All`
   - `Calendars.ReadWrite`
   - `OnlineMeetings.ReadWrite`
   - `Presence.ReadWrite`
8. Click **Grant admin consent** (or ask your admin)

---

## 2. Run Locally

```bash
# Clone and install
cd teams-mcp
pip install -r requirements.txt

# Set environment variables
export MS_CLIENT_ID="your-client-id-here"
export MS_TENANT_ID="your-tenant-id-or-common"

# Run with SSE (default, for remote connections)
python server.py

# Or run with stdio (for local MCP clients like Claude Desktop)
MCP_TRANSPORT=stdio python server.py
```

On first run, it prints a device code and opens your browser. Sign in with your Microsoft account and authorize. The token is cached in `token_cache.json`.

---

## 3. Deploy to Render

### Option A: Using render.yaml (recommended)

1. Push this repo to GitHub
2. Go to [Render Dashboard](https://dashboard.render.com/) → **New** → **Blueprint**
3. Connect your repo — Render reads `render.yaml` automatically
4. Set the environment variables `MS_CLIENT_ID` and `MS_TENANT_ID`
5. Deploy

### Option B: Manual

1. **New** → **Web Service** → connect your repo
2. **Runtime**: Docker
3. **Environment variables**:
   - `MS_CLIENT_ID` = your client ID
   - `MS_TENANT_ID` = your tenant ID
   - `MCP_TRANSPORT` = `sse`
   - `PORT` = `8000`
4. Deploy

Your MCP endpoint will be:
```
https://your-service-name.onrender.com/sse
```

---

## 4. Connect to an LLM Client

### Claude Desktop (`claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "teams": {
      "url": "https://your-service-name.onrender.com/sse"
    }
  }
}
```

### Claude.ai (MCP connector)

Add as a remote MCP server with URL:
```
https://your-service-name.onrender.com/sse
```

### Local stdio mode

```json
{
  "mcpServers": {
    "teams": {
      "command": "python",
      "args": ["/path/to/teams-mcp/server.py"],
      "env": {
        "MS_CLIENT_ID": "your-client-id",
        "MS_TENANT_ID": "your-tenant-id",
        "MCP_TRANSPORT": "stdio"
      }
    }
  }
}
```

---

## 5. Example LLM Prompts

Once connected, an LLM can run operations like:

- *"Show me my upcoming meetings this week"* → calls `list_events`
- *"Send 'hello team!' to the General channel in Engineering"* → calls `list_teams` → `list_channels` → `send_channel_message`
- *"Schedule a 30-min Teams meeting with alice@company.com tomorrow at 2pm"* → calls `create_event` with `is_online=True`
- *"What's my current status?"* → calls `get_my_presence`
- *"Accept the meeting invite about Q3 planning"* → calls `list_events` → `respond_to_event`

---

## Architecture

```
LLM Client ←→ MCP (SSE/stdio) ←→ server.py ←→ Microsoft Graph API
                                       ↑
                                  MSAL device-code
                                  auth (browser sign-in)
```

- **Single file** (`server.py`) — no framework bloat
- **MSAL** handles OAuth with automatic token refresh
- **httpx** for async Graph API calls
- **FastMCP** from the `mcp` SDK for protocol handling
- Responses are trimmed to essential fields so LLMs don't get overwhelmed

---

## License

MIT
