import { describe, it, expect, vi } from "vitest";
import { createFakeServer } from "./helpers/fake-server.js";
import { createFakeGraphClient } from "./helpers/fake-graph-client.js";

const getGraphClientMock = vi.hoisted(() => vi.fn());
vi.mock("../src/graph/client.js", () => ({
  getGraphClient: getGraphClientMock,
}));

const { registerTeamTools } = await import("../src/tools/team-tools.js");

const { server, handlers } = createFakeServer();
registerTeamTools(server);

function useGraphClient(config: { response?: unknown; error?: unknown } = {}) {
  const fake = createFakeGraphClient(config);
  getGraphClientMock.mockResolvedValue(fake.client);
  return fake;
}

describe("list_teams", () => {
  it("returns the joined teams", async () => {
    const fake = useGraphClient({ response: { value: [{ id: "t1", displayName: "Engineering" }] } });

    const result = await handlers.get("list_teams")!({});

    expect(fake.requests[0]).toMatchObject({ path: "/me/joinedTeams?$select=id,displayName,description", method: "get" });
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.value[0].displayName).toBe("Engineering");
  });

  it("surfaces a Graph error", async () => {
    useGraphClient({ error: { code: "Forbidden", message: "no access" } });

    const result = await handlers.get("list_teams")!({});

    expect(result.content[0].text).toContain("Forbidden");
  });
});

describe("list_channels", () => {
  it("lists channels for a team", async () => {
    const fake = useGraphClient({ response: { value: [{ id: "c1", displayName: "General" }] } });

    const result = await handlers.get("list_channels")!({ teamId: "t1" });

    expect(fake.requests[0]).toMatchObject({
      path: "/teams/t1/channels?$select=id,displayName,description,membershipType",
      method: "get",
    });
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.value[0].displayName).toBe("General");
  });
});

describe("list_channel_messages", () => {
  it("lists messages with the default page size", async () => {
    const fake = useGraphClient({ response: { value: [] } });

    await handlers.get("list_channel_messages")!({ teamId: "t1", channelId: "c1" });

    expect(fake.requests[0]).toMatchObject({
      path: "/teams/t1/channels/c1/messages?$top=25",
      method: "get",
    });
  });

  it("respects a custom top", async () => {
    const fake = useGraphClient({ response: { value: [] } });

    await handlers.get("list_channel_messages")!({ teamId: "t1", channelId: "c1", top: 5 });

    expect(fake.requests[0]).toMatchObject({
      path: "/teams/t1/channels/c1/messages?$top=5",
    });
  });
});

describe("send_channel_message", () => {
  it("posts a top-level message", async () => {
    const fake = useGraphClient({ response: { id: "m1", webUrl: "https://teams.microsoft.com/m1" } });

    const result = await handlers.get("send_channel_message")!({
      teamId: "t1",
      channelId: "c1",
      content: "hello team",
    });

    expect(fake.requests[0]).toMatchObject({
      path: "/teams/t1/channels/c1/messages",
      method: "post",
      body: { body: { content: "hello team" } },
    });
    expect(result.content[0].text).toContain("posted successfully");
  });

  it("surfaces a Graph error for a bad channelId", async () => {
    useGraphClient({ error: { code: "ItemNotFound", message: "channel not found" } });

    const result = await handlers.get("send_channel_message")!({
      teamId: "t1",
      channelId: "bad",
      content: "hi",
    });

    expect(result.content[0].text).toContain("ItemNotFound");
  });
});
