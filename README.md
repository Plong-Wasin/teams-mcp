# teams-mcp

A Model Context Protocol (MCP) server for reading and posting to Microsoft Teams channels via the Microsoft Graph API.

## Features

- **List teams** — Discover the Teams the authenticated user has joined
- **List channels** — Discover the Channels within a Team
- **List channel messages** — Read recent Root Messages in a Channel, with a Reply preview and count for each
- **Send channel message** — Post a new Root Message, opening a new Thread
- **Reply to a channel message** — Post a Reply into an existing Thread
- **Get a channel thread** — Read a Root Message together with a page of its Replies
- **Edit a channel message / reply** — Correct a Root Message or a Reply in place
- **Search messages** — Find a message and act on it, with Team/Channel identifiers included

Chats (1:1/group) and meetings are out of scope for v1 — see `CONTEXT.md`.

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
| `list_channels` | List the Channels in a Team (`layoutType` is always `null` here — a Graph API limitation) |
| `get_channel` | Get a single Channel, including its real `layoutType` (`post` or `chat`) |
| `list_channel_messages` | List recent Root Messages. By default each is expanded with a preview of up to 5 of its most recent Replies (oldest-first), the true Reply count, and a `repliesTruncated` flag; pass `includeReplies: false` to skip it |
| `send_channel_message` | Post a new Root Message, opening a new Thread, optionally with a `subject` (post-layout channels only) and `mentions` (@mention users by UPN/email) |
| `edit_channel_message` | Edit the body of an existing Root Message |
| `reply_to_channel_message` | Post a Reply into an existing Thread. A Reply ID passed as the target is automatically redirected to its parent Root Message |
| `get_channel_thread` | Read a Root Message together with a page of its Replies (oldest-first within the page); pages move backwards through the conversation, newest page first — use the returned `nextCursor` to page further into the past |
| `edit_channel_reply` | Replace the body of an existing Reply |
| `search_messages` | Keyword-search the signed-in user's Teams chat/channel messages (Microsoft Search API, supports KQL qualifiers like `from:`, `sent>`). Each Channel Message result carries its `teamId`/`channelId` at the top level, ready for `get_channel_thread` or `reply_to_channel_message` |

## Scopes

This MCP requests the following Microsoft Graph permissions:

- `User.Read` — Required as a baseline scope for the device code flow to resolve the other Graph permissions (see note below); also reads the signed-in user's own profile
- `User.ReadBasic.All` — Resolve users by UPN/email for `mentions` in `send_channel_message`
- `Team.ReadBasic.All` — List joined Teams
- `Channel.ReadBasic.All` — List Channels
- `ChannelMessage.Read.All` — Read Channel Messages and Threads, and search them via `search_messages`
- `ChannelMessage.Send` — Post Root Messages and Replies
- `ChannelMessage.ReadWrite` — Edit a Root Message (`edit_channel_message`) or a Reply (`edit_channel_reply`) after it has been posted
- `Chat.Read` — Required by the Microsoft Search API (`search_messages`) alongside `ChannelMessage.Read.All`, even though this server only searches channel messages
- `offline_access` — Keep the session alive with a refresh token

> **Note:** the Microsoft Graph CLI public client (the default `clientId`) rejects a `scope` request that contains only Teams-specific permissions (`AADSTS70011`) unless a baseline scope like `User.Read` is included alongside them. `User.Read` isn't optional here — remove it and `auth_start` fails.

## Token Storage

Tokens are stored at `teams-mcp/.access-token.txt` (excluded from git via `.gitignore`).
