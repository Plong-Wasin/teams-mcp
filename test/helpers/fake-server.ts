import { vi } from "vitest";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

export type ToolHandler = (input: any) => Promise<{ content: Array<{ type: string; text: string }> }>;

export interface FakeServer {
  server: McpServer;
  handlers: Map<string, ToolHandler>;
  configs: Map<string, { description?: string; inputSchema?: Record<string, unknown> }>;
}

export function createFakeServer(): FakeServer {
  const handlers = new Map<string, ToolHandler>();
  const configs = new Map<string, { description?: string; inputSchema?: Record<string, unknown> }>();
  const server = {
    registerTool: vi.fn((name: string, config: { description?: string; inputSchema?: Record<string, unknown> }, handler: ToolHandler) => {
      handlers.set(name, handler);
      configs.set(name, config);
    }),
  };
  return { server: server as unknown as McpServer, handlers, configs };
}
