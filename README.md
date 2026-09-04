# teams-mcp

A Model Context Protocol (MCP) server for reading and posting to Microsoft Teams channels via the Microsoft Graph API.

## Features

- **List teams** — Discover the Teams the authenticated user has joined
- **List channels** — Discover the Channels within a Team
- **List channel messages** — Read recent top-level messages in a Channel
- **Send channel message** — Post a new top-level message to a Channel

Chats (1:1/group) and meetings are out of scope for v1 — see `docs/CONTEXT.md`.

## Authentication

Uses OAuth 2.0 **Device Code Flow** — no browser redirect required. Tokens are stored locally at `.access-token.txt` and refreshed automatically.

```
auth_start  →  visit the URL, enter the code  →  auth_poll
```

> **Note:** `auth_start` should only be called when another tool returns an authentication error. Do not call it proactively — all tools load the stored token automatically.

## Setup

### 1. Install dependencies and build

```bash
pnpm install
pnpm run build
```

### 2. Register in Claude Code

Add to your project's `.mcp.json`:

```json
{
  "mcpServers": {
    "teams": {
      "type": "stdio",
      "command": "node",
      "args": ["C:\\Users\\staff\\Documents\\mcp\\teams-mcp\\build\\index.js"]
    }
  }
}
```

### 3. Authenticate

Call `auth_start` once, follow the instructions, then call `auth_poll` to complete login.

## Tools

### Authentication
| Tool | Description |
|------|-------------|
| `auth_start` | Start the Microsoft authentication flow |
| `auth_poll` | Complete authentication after signing in |

### Teams & Channels
| Tool | Description |
|------|-------------|
| `list_teams` | List Teams the authenticated user has joined |
| `list_channels` | List the Channels in a Team |
| `list_channel_messages` | List recent top-level Channel Messages |
| `send_channel_message` | Post a new top-level Channel Message |

## Scopes

This MCP requests the following Microsoft Graph permissions:

- `Team.ReadBasic.All` — List joined Teams
- `Channel.ReadBasic.All` — List Channels
- `ChannelMessage.Read.All` — Read Channel Messages
- `ChannelMessage.Send` — Post Channel Messages
- `offline_access` — Keep the session alive with a refresh token

## Token Storage

Tokens are stored at `teams-mcp/.access-token.txt` (excluded from git via `.gitignore`).
