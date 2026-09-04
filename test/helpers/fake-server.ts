import { vi } from "vitest";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

export type ToolHandler = (input: any) => Promise<{ content: Array<{ type: string; text: string }> }>;

export interface FakeServer {
  server: McpServer;
  handlers: Map<string, ToolHandler>;
}

export function createFakeServer(): FakeServer {
  const handlers = new Map<string, ToolHandler>();
  const server = {
    registerTool: vi.fn((name: string, _config: unknown, handler: ToolHandler) => {
      handlers.set(name, handler);
    }),
  };
  return { server: server as unknown as McpServer, handlers };
}
