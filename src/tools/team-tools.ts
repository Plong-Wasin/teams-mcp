import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getGraphClient } from "../graph/client.js";
import { handleGraphError } from "../utils/error-handler.js";

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
        "List the Channels in a Team. Returns each Channel's id (needed for list_channel_messages and send_channel_message) and displayName.",
      inputSchema: {
        teamId: z.string().describe("The Team ID (from list_teams)"),
      },
    },
    async ({ teamId }) => {
      try {
        const client = await getGraphClient();

        const channels = await client
          .api(`/teams/${teamId}/channels?$select=id,displayName,description,membershipType`)
          .get();

        return {
          content: [{ type: "text", text: JSON.stringify(channels, null, 2) }],
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
        "Post a new top-level Channel Message. This is not a threaded reply to an existing message.",
      inputSchema: {
        teamId: z.string().describe("The Team ID (from list_teams)"),
        channelId: z.string().describe("The Channel ID (from list_channels)"),
        content: z.string().describe("The plain-text message body to post"),
      },
    },
    async ({ teamId, channelId, content }) => {
      try {
        const client = await getGraphClient();

        const message = await client
          .api(`/teams/${teamId}/channels/${channelId}/messages`)
          .post({ body: { content } });

        return {
          content: [
            {
              type: "text",
              text: `✅ Message posted successfully.\n${JSON.stringify({ id: message.id, webUrl: message.webUrl }, null, 2)}`,
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );
}
