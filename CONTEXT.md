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
A message posted to a Channel, visible to every Team member with access to that Channel. Every Channel Message is either a Root Message or a Reply.
_Avoid_: post, chat message (when meaning a Channel Message specifically)

**Root Message**:
A Channel Message with no parent — the one that opens a Thread. Only a Root Message can carry a subject line.
_Avoid_: top-level message, parent message, original post

**Reply**:
A Channel Message posted underneath a Root Message, belonging to that Root Message's Thread. A Reply cannot itself be replied to — a Thread is exactly one level deep.
_Avoid_: comment, response, threaded message

**Thread**:
A Root Message together with all of its Replies — the unit a reader needs in order to answer in context.
_Avoid_: conversation, discussion, post

**Chat** _(out of scope)_:
A 1:1 or group conversation outside any Team/Channel. No tool targets a Chat directly. The one accepted exception is `search_messages`, which queries Microsoft Search across both Channel Messages and Chat messages because the Search API does not separate them — reading a Chat message that way is tolerated, writing to a Chat is not.
