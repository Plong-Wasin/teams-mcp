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

    expect(fake.requests[fake.requests.length - 1]).toMatchObject({
      path: "/teams/t1/channels/c1/messages",
      method: "post",
      body: { body: { contentType: "html", content: "hello team" } },
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

describe("edit_channel_message", () => {
  it("patches the message body as html", async () => {
    const fake = useGraphClient({ response: {} });

    const result = await handlers.get("edit_channel_message")!({
      teamId: "t1",
      channelId: "c1",
      messageId: "m1",
      content: "line one<br>line two",
    });

    expect(fake.requests[fake.requests.length - 1]).toMatchObject({
      path: "/teams/t1/channels/c1/messages/m1",
      method: "patch",
      body: { body: { contentType: "html", content: "line one<br>line two" } },
    });
    expect(result.content[0].text).toContain("updated successfully");
  });

  it("preserves the existing subject when none is passed", async () => {
    const fake = useGraphClient({ response: { subject: "Original subject" } });

    await handlers.get("edit_channel_message")!({
      teamId: "t1",
      channelId: "c1",
      messageId: "m1",
      content: "updated body",
    });

    expect(fake.requests[0]).toMatchObject({
      path: "/teams/t1/channels/c1/messages/m1?$select=subject",
      method: "get",
    });
    expect(fake.requests[1]).toMatchObject({
      path: "/teams/t1/channels/c1/messages/m1",
      method: "patch",
      body: { subject: "Original subject" },
    });
  });

  it("overrides the subject when one is passed, without fetching the old one", async () => {
    const fake = useGraphClient({ response: {} });

    await handlers.get("edit_channel_message")!({
      teamId: "t1",
      channelId: "c1",
      messageId: "m1",
      content: "updated body",
      subject: "New subject",
    });

    expect(fake.requests).toHaveLength(1);
    expect(fake.requests[0]).toMatchObject({
      path: "/teams/t1/channels/c1/messages/m1",
      method: "patch",
      body: { subject: "New subject" },
    });
  });

  it("resolves and prepends mentions", async () => {
    const fake = useGraphClient({ response: { id: "u1", displayName: "Jane Doe" } });

    await handlers.get("edit_channel_message")!({
      teamId: "t1",
      channelId: "c1",
      messageId: "m1",
      content: "please review",
      mentions: ["jane@contoso.com"],
    });

    const patchRequest = fake.requests[fake.requests.length - 1];
    expect(patchRequest).toMatchObject({ path: "/teams/t1/channels/c1/messages/m1", method: "patch" });
    const body = (patchRequest.body as any).body;
    expect(body.contentType).toBe("html");
    expect(body.content).toBe('<at id="0">Jane Doe</at> please review');
  });

  it("surfaces a Graph error for a bad messageId", async () => {
    useGraphClient({ error: { code: "ItemNotFound", message: "message not found" } });

    const result = await handlers.get("edit_channel_message")!({
      teamId: "t1",
      channelId: "c1",
      messageId: "bad",
      content: "hi",
    });

    expect(result.content[0].text).toContain("ItemNotFound");
  });
});
