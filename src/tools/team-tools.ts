import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { Client } from "@microsoft/microsoft-graph-client";
import { z } from "zod";
import { getGraphClient } from "../graph/client.js";
import { handleGraphError } from "../utils/error-handler.js";

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
        "List recent top-level Channel Messages in a Channel, most recent first. Does not include threaded replies.",
      inputSchema: {
        teamId: z.string().describe("The Team ID (from list_teams)"),
        channelId: z.string().describe("The Channel ID (from list_channels)"),
        top: z.number().optional().describe("Maximum number of messages to return (default: 25)"),
      },
    },
    async ({ teamId, channelId, top }) => {
      try {
        const client = await getGraphClient();

        const messages = await client
          .api(`/teams/${teamId}/channels/${channelId}/messages?$top=${top ?? 25}`)
          .get();

        return {
          content: [{ type: "text", text: JSON.stringify(messages, null, 2) }],
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
        "Post a new top-level Channel Message. This is not a threaded reply to an existing message. The subject line is only applied if the channel's layoutType is 'post'; it is dropped for 'chat'-layout channels, where a subject is not meaningful (check layoutType via list_channels or get_channel).",
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
              text: `✅ Message posted successfully.\n${JSON.stringify({ id: message.id, webUrl: message.webUrl, channelLayoutType: channel.layoutType, subjectApplied: !!useSubject, mentionedCount: mentions?.length ?? 0 }, null, 2)}`,
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
        "Edit the body of an existing top-level Channel Message that the authenticated user sent. Replaces the entire message content — there is no partial/append edit. Teams shows an 'Edited' label on the message. The existing subject line is preserved automatically unless you pass a new one (Microsoft Graph clears the subject on any PATCH that omits it, so this tool fetches and re-sends it for you).",
      inputSchema: {
        teamId: z.string().describe("The Team ID (from list_teams)"),
        channelId: z.string().describe("The Channel ID (from list_channels)"),
        messageId: z.string().describe("The Message ID to edit (the 'id' returned by send_channel_message or list_channel_messages)"),
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
              text: `✅ Message updated successfully.\n${JSON.stringify({ id: messageId, subject: useSubject, mentionedCount: mentions?.length ?? 0 }, null, 2)}`,
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
        "Search across the authenticated user's Teams chat and channel messages using Microsoft Search (keyword search over message body, sender, and attachments). Supports KQL-style qualifiers: from:<person>, to:<person>, sent><date>, hasAttachment:true/false, isMentioned:true/false, mentions:<userId>. Only returns messages the signed-in user has access to.",
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

        return {
          content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );
}
