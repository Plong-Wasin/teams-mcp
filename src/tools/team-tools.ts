import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { Client } from "@microsoft/microsoft-graph-client";
import { z } from "zod";
import { getGraphClient } from "../graph/client.js";
import { handleGraphError } from "../utils/error-handler.js";
import { projectMessage, sortByCreatedAscending, REPLY_PREVIEW_SIZE } from "../graph/messages.js";

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

const CONTENT_DESCRIPTION =
  "The message body, sent to Teams as HTML (contentType 'html') — this is NOT plain text. " +
  "Literal \\n newlines will NOT produce a line break; use HTML tags instead: <br> for a line break, " +
  "<b>/<i> for bold/italic, <ul><li>...</li></ul> or <ol><li>...</li></ol> for lists, <p>...</p> for paragraphs. " +
  "Plain unformatted text with no tags still works and renders as a single line.";

const MENTIONS_DESCRIPTION =
  "User principal names or email addresses to @mention at the start of the message (e.g. ['jane@contoso.com']). Each is resolved to a user and rendered as a clickable @mention that notifies them.";

// Confirmed against a live tenant: GET /teams/{t}/channels/{c}/messages/{id} 404s when id is a
// Reply's own id — Graph only supports fetching a Reply nested under its known root
// (.../messages/{rootId}/replies/{replyId}), so a Reply's parent can never be resolved from the
// Reply's id alone. There is no way to tell that apart from "id doesn't exist at all" from the error.
function notRootMessageError(messageId: string): string {
  return (
    `Could not find "${messageId}" as a Root Message in this Channel. If this is the id of a Reply ` +
    "(for example, one returned by search_messages, or one previewed by list_channel_messages / " +
    "get_channel_thread), pass that Reply's parentId — its Root Message id — instead. Microsoft Graph " +
    "does not support resolving a Reply's parent from the Reply's own id."
  );
}

interface HtmlBody {
  body: { contentType: "html"; content: string };
  mentions?: Array<{
    id: number;
    mentionText: string;
    mentioned: { user: { id: string; displayName: string; userIdentityType: "aadUser" } };
  }>;
}

async function buildHtmlBody(client: Client, content: string, mentions?: string[]): Promise<HtmlBody> {
  if (!mentions || mentions.length === 0) {
    return { body: { contentType: "html", content } };
  }

  const resolvedMentions: HtmlBody["mentions"] = [];
  const mentionTags: string[] = [];

  for (let i = 0; i < mentions.length; i++) {
    const user = await client
      .api(`/users/${encodeURIComponent(mentions[i])}`)
      .select("id,displayName")
      .get();

    resolvedMentions.push({
      id: i,
      mentionText: user.displayName,
      mentioned: { user: { id: user.id, displayName: user.displayName, userIdentityType: "aadUser" } },
    });
    mentionTags.push(`<at id="${i}">${escapeHtml(user.displayName)}</at>`);
  }

  return {
    body: { contentType: "html", content: `${mentionTags.join(" ")} ${content}` },
    mentions: resolvedMentions,
  };
}

export function registerTeamTools(server: McpServer): void {
  // ── list_teams ───────────────────────────────────────────────────────────
  server.registerTool(
    "list_teams",
    {
      description:
        "List Microsoft Teams the authenticated user has joined. Returns each Team's id (needed for list_channels and other tools) and displayName.",
      inputSchema: {},
    },
    async () => {
      try {
        const client = await getGraphClient();

        const teams = await client
          .api("/me/joinedTeams?$select=id,displayName,description")
          .get();

        return {
          content: [{ type: "text", text: JSON.stringify(teams, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );

  // ── list_channels ────────────────────────────────────────────────────────
  server.registerTool(
    "list_channels",
    {
      description:
        "List the Channels in a Team. Returns each Channel's id (needed for list_channel_messages and send_channel_message), displayName, and layoutType ('post' = traditional post-reply channel, supports a subject line; 'chat' = continuous chat-like channel, subject is not meaningful there).",
      inputSchema: {
        teamId: z.string().describe("The Team ID (from list_teams)"),
      },
    },
    async ({ teamId }) => {
      try {
        const client = await getGraphClient();

        const channels = await client
          .api(`/teams/${teamId}/channels?$select=id,displayName,description,membershipType,layoutType`)
          .get();

        return {
          content: [{ type: "text", text: JSON.stringify(channels, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );

  // ── get_channel ──────────────────────────────────────────────────────────
  server.registerTool(
    "get_channel",
    {
      description:
        "Get details of a single Channel, including its true layoutType ('post' or 'chat'). Note: list_channels always returns layoutType as null due to a Microsoft Graph API limitation — use this tool to get the real value for a specific channel.",
      inputSchema: {
        teamId: z.string().describe("The Team ID (from list_teams)"),
        channelId: z.string().describe("The Channel ID (from list_channels)"),
      },
    },
    async ({ teamId, channelId }) => {
      try {
        const client = await getGraphClient();

        const channel = await client
          .api(`/teams/${teamId}/channels/${channelId}?$select=id,displayName,description,membershipType,layoutType,webUrl`)
          .get();

        return {
          content: [{ type: "text", text: JSON.stringify(channel, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );

  // ── list_channel_messages ────────────────────────────────────────────────
  server.registerTool(
    "list_channel_messages",
    {
      description:
        "List recent Root Messages in a Channel, most recent first — Replies are never returned as their own top-level entries here. " +
        "By default (includeReplies=true) each Root Message is expanded with a preview of at most 5 of its most recent Replies " +
        "(oldest-first), a replyCount, and a repliesTruncated flag telling you whether the preview left Replies out. " +
        "replyCount reflects what the Graph API's own $expand=replies returned, which Graph documents inconsistently and may " +
        "itself cap for a very high-traffic Root Message — treat it as a strong estimate, not a guarantee; get_channel_thread " +
        "has no such cap and is the source of truth for a specific Thread's full Reply count. " +
        "Five is fixed and not configurable — use get_channel_thread to read a Thread's Replies in full. " +
        "Set includeReplies=false to skip the expansion and save latency when you only need Root Messages.",
      inputSchema: {
        teamId: z.string().describe("The Team ID (from list_teams)"),
        channelId: z.string().describe("The Channel ID (from list_channels)"),
        top: z.number().optional().describe("Maximum number of Root Messages to return (default: 25)"),
        includeReplies: z
          .boolean()
          .optional()
          .describe("Expand each Root Message with a Reply preview and count (default: true)"),
      },
    },
    async ({ teamId, channelId, top, includeReplies }) => {
      try {
        const client = await getGraphClient();
        const expand = includeReplies === false ? "" : "&$expand=replies";

        const raw = await client
          .api(`/teams/${teamId}/channels/${channelId}/messages?$top=${top ?? 25}${expand}`)
          .get();

        const messages = (raw.value ?? []).map((rawMessage: any) => {
          const lean = projectMessage(rawMessage);
          if (includeReplies === false) return lean;

          const replies = (rawMessage.replies ?? []).map(projectMessage);
          const replyCount = replies.length;
          const preview = sortByCreatedAscending(replies).slice(-REPLY_PREVIEW_SIZE);

          return {
            ...lean,
            replyCount,
            replies: preview,
            repliesTruncated: replyCount > preview.length,
          };
        });

        return {
          content: [{ type: "text", text: JSON.stringify({ messages }, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );

  // ── send_channel_message ─────────────────────────────────────────────────
  server.registerTool(
    "send_channel_message",
    {
      description:
        "Post a new Root Message, opening a new Thread in the Channel. This is not a Reply — it does not attach to any existing " +
        "conversation. To answer inside an existing Thread instead, use reply_to_channel_message. The subject line is only " +
        "applied if the channel's layoutType is 'post'; it is dropped for 'chat'-layout channels, where a subject is not " +
        "meaningful (check layoutType via list_channels or get_channel).",
      inputSchema: {
        teamId: z.string().describe("The Team ID (from list_teams)"),
        channelId: z.string().describe("The Channel ID (from list_channels)"),
        subject: z.string().optional().describe("Optional bold subject/title line shown above the message body. Ignored for 'chat'-layout channels."),
        content: z.string().describe(CONTENT_DESCRIPTION),
        mentions: z.array(z.string()).optional().describe(MENTIONS_DESCRIPTION),
      },
    },
    async ({ teamId, channelId, subject, content, mentions }) => {
      try {
        const client = await getGraphClient();

        const channel = await client
          .api(`/teams/${teamId}/channels/${channelId}?$select=layoutType`)
          .get();

        const useSubject = subject && channel.layoutType !== "chat" ? subject : null;

        const { body, mentions: resolvedMentions } = await buildHtmlBody(client, content, mentions);
        const payload: Record<string, unknown> = { subject: useSubject, body };
        if (resolvedMentions) payload.mentions = resolvedMentions;

        const message = await client
          .api(`/teams/${teamId}/channels/${channelId}/messages`)
          .post(payload);

        return {
          content: [
            {
              type: "text",
              text: `✅ Root Message posted successfully.\n${JSON.stringify({ id: message.id, webUrl: message.webUrl, channelLayoutType: channel.layoutType, subjectApplied: !!useSubject, mentionedCount: mentions?.length ?? 0 }, null, 2)}`,
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );

  // ── edit_channel_message ─────────────────────────────────────────────────
  server.registerTool(
    "edit_channel_message",
    {
      description:
        "Edit the body of an existing Root Message that the authenticated user sent. To edit a Reply instead, use " +
        "edit_channel_reply. Replaces the entire message content — there is no partial/append edit. Teams shows an " +
        "'Edited' label on the message. The existing subject line is preserved automatically unless you pass a new one " +
        "(Microsoft Graph clears the subject on any PATCH that omits it, so this tool fetches and re-sends it for you).",
      inputSchema: {
        teamId: z.string().describe("The Team ID (from list_teams)"),
        channelId: z.string().describe("The Channel ID (from list_channels)"),
        messageId: z.string().describe("The Root Message ID to edit (the 'id' returned by send_channel_message or list_channel_messages)"),
        content: z.string().describe(CONTENT_DESCRIPTION),
        subject: z.string().optional().describe("Optional replacement subject line. Omit to keep the message's current subject unchanged."),
        mentions: z.array(z.string()).optional().describe(MENTIONS_DESCRIPTION),
      },
    },
    async ({ teamId, channelId, messageId, content, subject, mentions }) => {
      try {
        const client = await getGraphClient();

        let useSubject = subject ?? null;
        if (useSubject === null) {
          const existing = await client
            .api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}?$select=subject`)
            .get();
          useSubject = existing.subject ?? null;
        }

        const { body, mentions: resolvedMentions } = await buildHtmlBody(client, content, mentions);
        const payload: Record<string, unknown> = { body };
        if (useSubject !== null) payload.subject = useSubject;
        if (resolvedMentions) payload.mentions = resolvedMentions;

        await client
          .api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}`)
          .patch(payload);

        return {
          content: [
            {
              type: "text",
              text: `✅ Root Message updated successfully.\n${JSON.stringify({ id: messageId, subject: useSubject, mentionedCount: mentions?.length ?? 0 }, null, 2)}`,
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );

  // ── reply_to_channel_message ─────────────────────────────────────────────
  server.registerTool(
    "reply_to_channel_message",
    {
      description:
        "Post a Reply underneath a Root Message, so it appears attached to that Thread instead of floating as a new Root " +
        "Message. Operates on a Reply, not a Root Message — there is no subject parameter, since a subject is a Root Message " +
        "concept. messageId MUST be a Root Message id, not a Reply id — Microsoft Graph has no way to resolve a Reply's " +
        "parent from the Reply's own id, so this cannot be done automatically. If you have a Reply (for example, one " +
        "returned by search_messages), use that Reply's parentId as messageId instead of its own id. Passing a Reply's own " +
        "id fails with a clear error rather than posting somewhere unexpected.",
      inputSchema: {
        teamId: z.string().describe("The Team ID (from list_teams)"),
        channelId: z.string().describe("The Channel ID (from list_channels)"),
        messageId: z.string().describe("The Root Message to answer. Must be a Root Message id — if you only have a Reply, pass that Reply's parentId here instead."),
        content: z.string().describe(CONTENT_DESCRIPTION),
        mentions: z.array(z.string()).optional().describe(MENTIONS_DESCRIPTION),
      },
    },
    async ({ teamId, channelId, messageId, content, mentions }) => {
      try {
        const client = await getGraphClient();

        try {
          await client
            .api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}?$select=id`)
            .get();
        } catch (error: any) {
          if (error?.code === "NotFound") {
            return { content: [{ type: "text", text: notRootMessageError(messageId) }] };
          }
          throw error;
        }

        const { body, mentions: resolvedMentions } = await buildHtmlBody(client, content, mentions);
        const payload: Record<string, unknown> = { body };
        if (resolvedMentions) payload.mentions = resolvedMentions;

        const reply = await client
          .api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`)
          .post(payload);

        return {
          content: [
            {
              type: "text",
              text: `✅ Reply posted successfully.\n${JSON.stringify(
                {
                  id: reply.id,
                  webUrl: reply.webUrl,
                  rootMessageId: messageId,
                  mentionedCount: mentions?.length ?? 0,
                },
                null,
                2,
              )}`,
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );

  // ── edit_channel_reply ───────────────────────────────────────────────────
  server.registerTool(
    "edit_channel_reply",
    {
      description:
        "Edit the body of an existing Reply that the authenticated user sent. To edit a Root Message instead, use " +
        "edit_channel_message. Replaces the entire Reply content — there is no partial/append edit. There is no subject " +
        "parameter, since a subject is a Root Message concept.",
      inputSchema: {
        teamId: z.string().describe("The Team ID (from list_teams)"),
        channelId: z.string().describe("The Channel ID (from list_channels)"),
        rootMessageId: z.string().describe("The Root Message ID of the Thread the Reply belongs to"),
        replyId: z.string().describe("The Reply ID to edit"),
        content: z.string().describe(CONTENT_DESCRIPTION),
        mentions: z.array(z.string()).optional().describe(MENTIONS_DESCRIPTION),
      },
    },
    async ({ teamId, channelId, rootMessageId, replyId, content, mentions }) => {
      try {
        const client = await getGraphClient();

        const { body, mentions: resolvedMentions } = await buildHtmlBody(client, content, mentions);
        const payload: Record<string, unknown> = { body };
        if (resolvedMentions) payload.mentions = resolvedMentions;

        await client
          .api(`/teams/${teamId}/channels/${channelId}/messages/${rootMessageId}/replies/${replyId}`)
          .patch(payload);

        return {
          content: [
            {
              type: "text",
              text: `✅ Reply updated successfully.\n${JSON.stringify({ id: replyId, rootMessageId, mentionedCount: mentions?.length ?? 0 }, null, 2)}`,
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );

  // ── get_channel_thread ───────────────────────────────────────────────────
  server.registerTool(
    "get_channel_thread",
    {
      description:
        "Read a whole Thread: the Root Message together with a page of its Replies, in one response. messageId MUST be a " +
        "Root Message id, not a Reply id — Microsoft Graph has no way to resolve a Reply's parent from the Reply's own id, " +
        "so this cannot be done automatically. If you have a Reply (for example, one returned by search_messages, or one " +
        "previewed by list_channel_messages), use that Reply's parentId as messageId instead. A page holds at most 50 " +
        "Replies (the Graph API's own ceiling). IMPORTANT — paging direction: the Graph API returns Replies newest-first, " +
        "so the first page (no cursor) is the END of the conversation, not the beginning; passing the returned nextCursor " +
        "into another call moves further back into the Thread's past. Within each page, Replies are sorted oldest-first " +
        "for readability. This means a truncated page (nextCursor present) looks exactly like a Thread whose beginning is " +
        "missing — it is not.",
      inputSchema: {
        teamId: z.string().describe("The Team ID (from list_teams)"),
        channelId: z.string().describe("The Channel ID (from list_channels)"),
        messageId: z.string().describe("The Root Message identifying the Thread to read. Must be a Root Message id — if you only have a Reply, pass that Reply's parentId here instead."),
        cursor: z
          .string()
          .optional()
          .describe("Opaque nextCursor from a previous get_channel_thread call, passed back unchanged, to page further into the Thread's past"),
      },
    },
    async ({ teamId, channelId, messageId, cursor }) => {
      try {
        const client = await getGraphClient();

        let rootRaw: any;
        try {
          rootRaw = await client
            .api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}`)
            .get();
        } catch (error: any) {
          if (error?.code === "NotFound") {
            return { content: [{ type: "text", text: notRootMessageError(messageId) }] };
          }
          throw error;
        }

        // `cursor` is the opaque `@odata.nextLink` Graph handed back on a previous page — the Graph
        // client accepts a full URL directly, so it's passed straight through rather than parsed and
        // rebuilt, which would tie this to an unconfirmed assumption about the nextLink's query shape.
        const repliesRaw = cursor
          ? await client.api(cursor).get()
          : await client
              .api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies?$top=50`)
              .get();

        const replies = sortByCreatedAscending((repliesRaw.value ?? []).map(projectMessage));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  root: projectMessage(rootRaw),
                  replies,
                  nextCursor: repliesRaw["@odata.nextLink"] ?? null,
                },
                null,
                2,
              ),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );

  // ── search_messages ──────────────────────────────────────────────────────
  server.registerTool(
    "search_messages",
    {
      description:
        "Search across the authenticated user's Teams chat and channel messages using Microsoft Search (keyword search over message body, sender, and attachment metadata). Supports KQL-style qualifiers: from:<person>, to:<person>, sent><date>, hasAttachment:true/false, isMentioned:true/false, mentions:<userId>. Only returns messages the signed-in user has access to. " +
        "Each Channel Message result carries its Team and Channel identifiers at the top level, ready to hand to " +
        "get_channel_thread or reply_to_channel_message; a result from a Chat rather than a Channel has both set to null, " +
        "since Chats are out of scope for every other tool. get_channel_thread and reply_to_channel_message both require a " +
        "Root Message id: if a result's message.parentId is null the result IS a Root Message, so use message.id directly; " +
        "if message.parentId is set, the result is a Reply, so use message.parentId (its Root Message) instead — Microsoft " +
        "Graph cannot resolve a Reply's parent from the Reply's own id. " +
        "Microsoft Search does not report Attachments, so every hit has message.attachments = null (unknown, not empty). " +
        "A hit with an empty body may be a Card. For a Channel Message hit, read the Card with get_channel_thread or " +
        "list_channel_messages using the hit's Team and Channel identifiers; a Card in a Chat hit cannot be read.",
      inputSchema: {
        query: z
          .string()
          .describe("Search text, optionally combined with KQL qualifiers, e.g. \"deploy from:bob sent>2024-01-01\""),
        size: z.number().optional().describe("Maximum number of results to return (default: 25)"),
      },
    },
    async ({ query, size }) => {
      try {
        const client = await getGraphClient();

        const result = await client.api("/search/query").post({
          requests: [
            {
              entityTypes: ["chatMessage"],
              query: { queryString: query },
              from: 0,
              size: size ?? 25,
              enableTopResults: true,
            },
          ],
        });

        const hits = result?.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
        const results = hits.map((hit: any) => {
          const resource = hit.resource ?? {};
          const channelIdentity = resource.channelIdentity;
          return {
            teamId: channelIdentity?.teamId ?? null,
            channelId: channelIdentity?.channelId ?? null,
            message: { ...projectMessage(resource), attachments: null },
            summary: hit.summary ?? null,
          };
        });

        return {
          content: [{ type: "text", text: JSON.stringify({ results }, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );
}
