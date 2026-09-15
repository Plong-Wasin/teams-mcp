# One tool per operation, never an optional parameter that switches modes

When adding Reply support we repeatedly faced the same choice: add a new MCP tool, or add an
optional parameter to an existing one (`replyToMessageId` on `send_channel_message`, `replyId` on
`edit_channel_message`, `includeRoot` on the thread reader). Each time the parameter was the smaller
diff, and each time we chose the separate tool — because an MCP tool name is the only thing a model
reads before it decides what to call, and a parameter that silently invalidates a sibling parameter
is a footgun we have already paid for once in this repo (`subject` is silently dropped in
`layoutType: 'chat'` channels, which now costs a paragraph of warning text in two tool descriptions).

## Consequences

- The tool count grows faster than the feature count. That is accepted.
- `list_channel_messages` keeps an `includeReplies` flag, which looks like a violation. It is not:
  the flag enriches the same operation (listing a Channel) rather than switching it to a different
  one. The test is whether the tool's *verb* changes, not whether a parameter exists.
- Tool names end up in users' saved prompts, so renaming later is a real cost — which is precisely
  why the shape is worth getting right up front rather than refactoring into it.
