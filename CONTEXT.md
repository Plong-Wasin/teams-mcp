# Teams MCP

An MCP server that lets Claude read and post to Microsoft Teams channels via the Microsoft Graph API.

## Language

**Team**:
A Microsoft Teams team the authenticated user has joined, backed by a Microsoft 365 Group. A Team contains Channels.
_Avoid_: group, workspace (when meaning a Team specifically)

**Channel**:
A named conversation space within a Team where Channel Messages are posted and read.
_Avoid_: room, thread (when meaning a Channel)

**Channel Message**:
A message posted to a Channel, visible to every Team member with access to that Channel. Distinct from a Chat Message (1:1 or group chat), which is out of scope for v1.
_Avoid_: post, chat message (when meaning a Channel Message specifically)

**Chat** _(out of scope for v1)_:
A 1:1 or group conversation outside any Team/Channel. Reserved for a later iteration — do not build against it yet.
