import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

import { registerAuthTools } from "./tools/auth-tools.js";
import { registerTeamTools } from "./tools/team-tools.js";

const server = new McpServer({
  name: "teams-mcp",
  version: "1.0.0",
});

registerAuthTools(server);
registerTeamTools(server);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Teams MCP server running on stdio");
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
});
