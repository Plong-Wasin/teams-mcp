import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { startDeviceCodeFlow, pollForToken } from "../auth/device-code-flow.js";

export function registerAuthTools(server: McpServer): void {
  server.registerTool(
    "auth_start",
    {
      description:
        "Start the Microsoft Teams authentication flow. IMPORTANT: Only call this tool when you receive an authentication error (e.g. 'Not authenticated') from another tool, or when the user explicitly requests re-authentication. Do NOT call this proactively before using other Teams tools — they automatically use the stored token. If a valid token already exists, this returns immediately without starting a new auth flow.",
      inputSchema: {
        clientId: z
          .string()
          .optional()
          .describe("Azure AD application client ID (optional, uses Microsoft Graph Explorer by default)"),
        tenantId: z
          .string()
          .optional()
          .describe("Azure AD tenant ID (optional, uses 'common' by default)"),
        force: z
          .boolean()
          .optional()
          .describe("Force re-authentication even if a valid token exists (default: false)"),
      },
    },
    async ({ clientId, tenantId, force }) => {
      const result = await startDeviceCodeFlow(clientId, tenantId, force);
      return {
        content: [{ type: "text", text: result.message }],
      };
    },
  );

  server.registerTool(
    "auth_poll",
    {
      description:
        "Check if the authentication is complete. Only call this after auth_start returned a verification URL and the user has completed sign-in on the Microsoft website. Do NOT call this unless you just ran auth_start and it returned a code/URL for the user to visit.",
      inputSchema: {},
    },
    async () => {
      const result = await pollForToken();
      return {
        content: [{ type: "text", text: result.message }],
      };
    },
  );
}
