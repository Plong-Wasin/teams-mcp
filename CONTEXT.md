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

**Attachment**:
Anything attached to a Channel Message, such as a Card or a reference to a file. A Channel Message may carry zero or more Attachments.
_Avoid_: embed, file (when meaning an Attachment in general)

**Card**:
An Attachment holding structured content (for example an Adaptive Card), typically posted by a bot or webhook. A Channel Message that carries a Card often has an almost empty body, so the Card is the message's real content.
_Avoid_: embed, rich message

**Chat** _(out of scope)_:
A 1:1 or group conversation outside any Team/Channel. No tool targets a Chat directly. The one accepted exception is `search_messages`, which queries Microsoft Search across both Channel Messages and Chat messages because the Search API does not separate them — reading a Chat message that way is tolerated, writing to a Chat is not.
