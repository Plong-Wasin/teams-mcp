import { describe, it, expect, vi } from "vitest";
import { createFakeServer } from "./helpers/fake-server.js";
import { createFakeGraphClient, type FakeGraphClientConfig } from "./helpers/fake-graph-client.js";

const getGraphClientMock = vi.hoisted(() => vi.fn());
vi.mock("../src/graph/client.js", () => ({
  getGraphClient: getGraphClientMock,
}));

const { registerTeamTools } = await import("../src/tools/team-tools.js");

const { server, handlers, configs } = createFakeServer();
registerTeamTools(server);

function useGraphClient(config: FakeGraphClientConfig = {}) {
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
      path: "/teams/t1/channels?$select=id,displayName,description,membershipType,layoutType",
      method: "get",
    });
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.value[0].displayName).toBe("General");
  });
});

function rawReply(id: string, createdDateTime: string) {
  return {
    id,
    replyToId: "root1",
    from: { user: { displayName: `Author ${id}` } },
    createdDateTime,
    subject: null,
    body: { content: `<p>${id}</p>` },
    webUrl: `https://teams.microsoft.com/${id}`,
  };
}

describe("list_channel_messages", () => {
  it("lists Root Messages with replies expanded by default", async () => {
    const fake = useGraphClient({ response: { value: [] } });

    await handlers.get("list_channel_messages")!({ teamId: "t1", channelId: "c1" });

    expect(fake.requests[0]).toMatchObject({
      path: "/teams/t1/channels/c1/messages?$top=25&$expand=replies",
      method: "get",
    });
  });

  it("respects a custom top", async () => {
    const fake = useGraphClient({ response: { value: [] } });

    await handlers.get("list_channel_messages")!({ teamId: "t1", channelId: "c1", top: 5 });

    expect(fake.requests[0]).toMatchObject({
      path: "/teams/t1/channels/c1/messages?$top=5&$expand=replies",
    });
  });

  it("skips the $expand when includeReplies is false", async () => {
    const fake = useGraphClient({ response: { value: [] } });

    await handlers.get("list_channel_messages")!({ teamId: "t1", channelId: "c1", includeReplies: false });

    expect(fake.requests[0]).toMatchObject({
      path: "/teams/t1/channels/c1/messages?$top=25",
    });
  });

  it("projects Root Messages to the lean shape", async () => {
    useGraphClient({
      response: {
        value: [
          {
            id: "root1",
            replyToId: null,
            from: { user: { displayName: "Alice" } },
            createdDateTime: "2024-01-01T00:00:00Z",
            subject: "Deploy question",
            body: { content: "<p>Can we deploy?</p>" },
            webUrl: "https://teams.microsoft.com/root1",
            etag: "should-be-dropped",
            policyViolation: null,
          },
        ],
      },
    });

    const result = await handlers.get("list_channel_messages")!({ teamId: "t1", channelId: "c1" });
    const parsed = JSON.parse(result.content[0].text);

    expect(parsed.messages[0]).toMatchObject({
      id: "root1",
      parentId: null,
      from: "Alice",
      createdDateTime: "2024-01-01T00:00:00Z",
      subject: "Deploy question",
      body: "<p>Can we deploy?</p>",
      webUrl: "https://teams.microsoft.com/root1",
    });
    expect(parsed.messages[0].etag).toBeUndefined();
  });

  it("previews at most 5 of the most recent Replies, oldest-first, with the true count", async () => {
    const replies = [
      rawReply("r1", "2024-01-01T01:00:00Z"),
      rawReply("r2", "2024-01-01T02:00:00Z"),
      rawReply("r3", "2024-01-01T03:00:00Z"),
      rawReply("r4", "2024-01-01T04:00:00Z"),
      rawReply("r5", "2024-01-01T05:00:00Z"),
      rawReply("r6", "2024-01-01T06:00:00Z"),
      rawReply("r7", "2024-01-01T07:00:00Z"),
    ];
    // Shuffle to prove ordering is computed, not inherited from Graph.
    const shuffled = [replies[3], replies[0], replies[6], replies[1], replies[5], replies[2], replies[4]];

    useGraphClient({
      response: {
        value: [
          {
            id: "root1",
            replyToId: null,
            from: { user: { displayName: "Alice" } },
            createdDateTime: "2024-01-01T00:00:00Z",
            subject: "Deploy question",
            body: { content: "<p>Can we deploy?</p>" },
            webUrl: "https://teams.microsoft.com/root1",
            replies: shuffled,
          },
        ],
      },
    });

    const result = await handlers.get("list_channel_messages")!({ teamId: "t1", channelId: "c1" });
    const parsed = JSON.parse(result.content[0].text);
    const root = parsed.messages[0];

    expect(root.replyCount).toBe(7);
    expect(root.repliesTruncated).toBe(true);
    expect(root.replies.map((r: any) => r.id)).toEqual(["r3", "r4", "r5", "r6", "r7"]);
  });

  it("does not truncate when there are 5 or fewer Replies", async () => {
    useGraphClient({
      response: {
        value: [
          {
            id: "root1",
            replyToId: null,
            from: { user: { displayName: "Alice" } },
            createdDateTime: "2024-01-01T00:00:00Z",
            subject: null,
            body: { content: "<p>hi</p>" },
            webUrl: "https://teams.microsoft.com/root1",
            replies: [rawReply("r1", "2024-01-01T01:00:00Z"), rawReply("r2", "2024-01-01T02:00:00Z")],
          },
        ],
      },
    });

    const result = await handlers.get("list_channel_messages")!({ teamId: "t1", channelId: "c1" });
    const root = JSON.parse(result.content[0].text).messages[0];

    expect(root.replyCount).toBe(2);
    expect(root.repliesTruncated).toBe(false);
    expect(root.replies.map((r: any) => r.id)).toEqual(["r1", "r2"]);
  });

  it("omits reply preview fields when includeReplies is false", async () => {
    useGraphClient({
      response: {
        value: [
          {
            id: "root1",
            replyToId: null,
            from: { user: { displayName: "Alice" } },
            createdDateTime: "2024-01-01T00:00:00Z",
            subject: null,
            body: { content: "<p>hi</p>" },
            webUrl: "https://teams.microsoft.com/root1",
            replies: [rawReply("r1", "2024-01-01T01:00:00Z")],
          },
        ],
      },
    });

    const result = await handlers.get("list_channel_messages")!({
      teamId: "t1",
      channelId: "c1",
      includeReplies: false,
    });
    const root = JSON.parse(result.content[0].text).messages[0];

    expect(root.replyCount).toBeUndefined();
    expect(root.replies).toBeUndefined();
    expect(root.repliesTruncated).toBeUndefined();
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
    const fake = useGraphClient({
      responses: {
        "/teams/t1/channels/c1/messages/m1?$select=subject": { response: { subject: "Original subject" } },
        "/teams/t1/channels/c1/messages/m1": { response: {} },
      },
    });

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

describe("reply_to_channel_message", () => {
  it("posts a Reply under a Root Message", async () => {
    const fake = useGraphClient({
      responses: {
        "/teams/t1/channels/c1/messages/root1?$select=id": { response: { id: "root1" } },
        "/teams/t1/channels/c1/messages/root1/replies": { response: { id: "reply1", webUrl: "https://teams.microsoft.com/reply1" } },
      },
    });

    const result = await handlers.get("reply_to_channel_message")!({
      teamId: "t1",
      channelId: "c1",
      messageId: "root1",
      content: "here's the answer",
    });

    expect(fake.requests[0]).toMatchObject({ path: "/teams/t1/channels/c1/messages/root1?$select=id", method: "get" });
    expect(fake.requests[1]).toMatchObject({
      path: "/teams/t1/channels/c1/messages/root1/replies",
      method: "post",
      body: { body: { contentType: "html", content: "here's the answer" } },
    });
    const parsed = JSON.parse(result.content[0].text.split("\n").slice(1).join("\n"));
    expect(parsed.rootMessageId).toBe("root1");
    expect(result.content[0].text).toContain("Reply posted successfully");
  });

  // Confirmed against a live tenant: Graph 404s a GET by a Reply's own id — there is no way to
  // resolve a Reply's parent from the Reply's id alone, so the tool must fail loudly instead of
  // guessing where to post (it has no delete capability to clean up a wrong guess).
  it("fails with an actionable message rather than posting, when messageId is not a Root Message", async () => {
    const fake = useGraphClient({
      responses: {
        "/teams/t1/channels/c1/messages/reply5?$select=id": { error: { code: "NotFound", message: "message not found" } },
      },
    });

    const result = await handlers.get("reply_to_channel_message")!({
      teamId: "t1",
      channelId: "c1",
      messageId: "reply5",
      content: "correction",
    });

    expect(result.content[0].text).toContain("reply5");
    expect(result.content[0].text).toContain("parentId");
    expect(fake.requests).toHaveLength(1);
  });

  it("resolves and prepends mentions", async () => {
    const fake = useGraphClient({
      responses: {
        "/teams/t1/channels/c1/messages/root1?$select=id": { response: { id: "root1" } },
        "/users/jane%40contoso.com": { response: { id: "u1", displayName: "Jane Doe" } },
        "/teams/t1/channels/c1/messages/root1/replies": { response: { id: "reply1" } },
      },
    });

    await handlers.get("reply_to_channel_message")!({
      teamId: "t1",
      channelId: "c1",
      messageId: "root1",
      content: "please review",
      mentions: ["jane@contoso.com"],
    });

    const postRequest = fake.requests[fake.requests.length - 1];
    const body = (postRequest.body as any).body;
    expect(body.content).toBe('<at id="0">Jane Doe</at> please review');
  });

  it("surfaces a non-NotFound Graph error rather than throwing", async () => {
    useGraphClient({ error: { code: "Forbidden", message: "no access" } });

    const result = await handlers.get("reply_to_channel_message")!({
      teamId: "t1",
      channelId: "c1",
      messageId: "root1",
      content: "hi",
    });

    expect(result.content[0].text).toContain("Forbidden");
  });
});

describe("edit_channel_reply", () => {
  it("replaces an existing Reply's body", async () => {
    const fake = useGraphClient({ response: {} });

    const result = await handlers.get("edit_channel_reply")!({
      teamId: "t1",
      channelId: "c1",
      rootMessageId: "root1",
      replyId: "reply1",
      content: "corrected text",
    });

    expect(fake.requests[fake.requests.length - 1]).toMatchObject({
      path: "/teams/t1/channels/c1/messages/root1/replies/reply1",
      method: "patch",
      body: { body: { contentType: "html", content: "corrected text" } },
    });
    expect(result.content[0].text).toContain("Reply updated successfully");
  });

  it("resolves and prepends mentions", async () => {
    const fake = useGraphClient({
      responses: {
        "/users/jane%40contoso.com": { response: { id: "u1", displayName: "Jane Doe" } },
        "/teams/t1/channels/c1/messages/root1/replies/reply1": { response: {} },
      },
    });

    await handlers.get("edit_channel_reply")!({
      teamId: "t1",
      channelId: "c1",
      rootMessageId: "root1",
      replyId: "reply1",
      content: "please review",
      mentions: ["jane@contoso.com"],
    });

    const patchRequest = fake.requests[fake.requests.length - 1];
    const body = (patchRequest.body as any).body;
    expect(body.content).toBe('<at id="0">Jane Doe</at> please review');
  });

  it("succeeds against a response with no body, without attempting to parse one", async () => {
    useGraphClient({ response: undefined });

    const result = await handlers.get("edit_channel_reply")!({
      teamId: "t1",
      channelId: "c1",
      rootMessageId: "root1",
      replyId: "reply1",
      content: "fixed",
    });

    expect(result.content[0].text).toContain("Reply updated successfully");
  });

  it("surfaces a Graph error rather than throwing", async () => {
    useGraphClient({ error: { code: "ItemNotFound", message: "reply not found" } });

    const result = await handlers.get("edit_channel_reply")!({
      teamId: "t1",
      channelId: "c1",
      rootMessageId: "root1",
      replyId: "bad",
      content: "fixed",
    });

    expect(result.content[0].text).toContain("ItemNotFound");
  });
});

describe("get_channel_thread", () => {
  const rootRaw = {
    id: "root1",
    replyToId: null,
    from: { user: { displayName: "Alice" } },
    createdDateTime: "2024-01-01T00:00:00Z",
    subject: "Deploy question",
    body: { content: "<p>Can we deploy?</p>" },
    webUrl: "https://teams.microsoft.com/root1",
  };

  it("returns the Root Message together with Replies sorted oldest-first", async () => {
    const shuffledReplies = [rawReply("r3", "2024-01-01T03:00:00Z"), rawReply("r1", "2024-01-01T01:00:00Z"), rawReply("r2", "2024-01-01T02:00:00Z")];

    const fake = useGraphClient({
      responses: {
        "/teams/t1/channels/c1/messages/root1": { response: rootRaw },
        "/teams/t1/channels/c1/messages/root1/replies?$top=50": { response: { value: shuffledReplies } },
      },
    });

    const result = await handlers.get("get_channel_thread")!({ teamId: "t1", channelId: "c1", messageId: "root1" });
    const parsed = JSON.parse(result.content[0].text);

    expect(fake.requests[0]).toMatchObject({ path: "/teams/t1/channels/c1/messages/root1", method: "get" });
    expect(parsed.root).toMatchObject({ id: "root1", parentId: null, subject: "Deploy question" });
    expect(parsed.replies.map((r: any) => r.id)).toEqual(["r1", "r2", "r3"]);
    expect(parsed.nextCursor).toBeNull();
  });

  // Confirmed against a live tenant: Graph 404s a GET by a Reply's own id, so there is no way to
  // resolve a Reply's parent from the Reply's id alone — the tool must fail loudly and say so,
  // rather than silently returning an empty Thread (which would read as "nobody answered").
  it("fails with an actionable message rather than an empty Thread, when messageId is a Reply id", async () => {
    const fake = useGraphClient({
      responses: {
        "/teams/t1/channels/c1/messages/reply3": { error: { code: "NotFound", message: "message not found" } },
      },
    });

    const result = await handlers.get("get_channel_thread")!({ teamId: "t1", channelId: "c1", messageId: "reply3" });

    expect(result.content[0].text).toContain("reply3");
    expect(result.content[0].text).toContain("parentId");
    expect(fake.requests).toHaveLength(1);
  });

  it("returns a cursor when more Replies exist, and consumes it verbatim on the next call", async () => {
    const nextLink = "https://graph.microsoft.com/v1.0/teams/t1/channels/c1/messages/root1/replies?$top=50&$skiptoken=abc123";

    useGraphClient({
      responses: {
        "/teams/t1/channels/c1/messages/root1": { response: rootRaw },
        "/teams/t1/channels/c1/messages/root1/replies?$top=50": {
          response: {
            value: [rawReply("r5", "2024-01-01T05:00:00Z")],
            "@odata.nextLink": nextLink,
          },
        },
      },
    });

    const first = await handlers.get("get_channel_thread")!({ teamId: "t1", channelId: "c1", messageId: "root1" });
    const firstParsed = JSON.parse(first.content[0].text);
    expect(firstParsed.nextCursor).toBe(nextLink);

    const fake2 = useGraphClient({
      responses: {
        "/teams/t1/channels/c1/messages/root1": { response: rootRaw },
        [nextLink]: { response: { value: [rawReply("r1", "2024-01-01T01:00:00Z")] } },
      },
    });

    const second = await handlers.get("get_channel_thread")!({ teamId: "t1", channelId: "c1", messageId: "root1", cursor: nextLink });
    const secondParsed = JSON.parse(second.content[0].text);

    expect(fake2.requests.some((r) => r.path === nextLink)).toBe(true);
    expect(secondParsed.replies.map((r: any) => r.id)).toEqual(["r1"]);
  });

  it("surfaces a Graph error rather than throwing", async () => {
    useGraphClient({ error: { code: "ItemNotFound", message: "message not found" } });

    const result = await handlers.get("get_channel_thread")!({ teamId: "t1", channelId: "c1", messageId: "bad" });

    expect(result.content[0].text).toContain("ItemNotFound");
  });
});

describe("search_messages", () => {
  it("lifts Team and Channel identifiers to the top level and preserves highlight markers", async () => {
    useGraphClient({
      response: {
        value: [
          {
            hitsContainers: [
              {
                hits: [
                  {
                    summary: "found the <c0>deploy</c0> question",
                    resource: {
                      id: "m1",
                      replyToId: null,
                      from: { user: { displayName: "Bob" } },
                      createdDateTime: "2024-01-01T00:00:00Z",
                      subject: "Deploy",
                      body: { content: "<p>deploy?</p>" },
                      webUrl: "https://teams.microsoft.com/m1",
                      channelIdentity: { teamId: "t1", channelId: "c1" },
                    },
                  },
                ],
              },
            ],
          },
        ],
      },
    });

    const result = await handlers.get("search_messages")!({ query: "deploy" });
    const parsed = JSON.parse(result.content[0].text);

    expect(parsed.results[0]).toMatchObject({
      teamId: "t1",
      channelId: "c1",
      summary: "found the <c0>deploy</c0> question",
    });
    expect(parsed.results[0].message).toMatchObject({ id: "m1", subject: "Deploy" });
  });

  it("sets Team and Channel to null for a Chat message hit without breaking the projection", async () => {
    useGraphClient({
      response: {
        value: [
          {
            hitsContainers: [
              {
                hits: [
                  {
                    summary: "a <c0>chat</c0> hit",
                    resource: {
                      id: "chatmsg1",
                      from: { user: { displayName: "Carol" } },
                      createdDateTime: "2024-01-01T00:00:00Z",
                      body: { content: "<p>hi</p>" },
                      webUrl: "https://teams.microsoft.com/chatmsg1",
                    },
                  },
                ],
              },
            ],
          },
        ],
      },
    });

    const result = await handlers.get("search_messages")!({ query: "chat" });
    const parsed = JSON.parse(result.content[0].text);

    expect(parsed.results[0].teamId).toBeNull();
    expect(parsed.results[0].channelId).toBeNull();
    expect(parsed.results[0].message.id).toBe("chatmsg1");
  });

  it("surfaces a Graph error rather than throwing", async () => {
    useGraphClient({ error: { code: "Forbidden", message: "no access" } });

    const result = await handlers.get("search_messages")!({ query: "deploy" });

    expect(result.content[0].text).toContain("Forbidden");
  });
});

describe("Reply tools have no subject parameter", () => {
  it("reply_to_channel_message exposes no subject parameter", () => {
    expect(configs.get("reply_to_channel_message")?.inputSchema).not.toHaveProperty("subject");
  });

  it("edit_channel_reply exposes no subject parameter", () => {
    expect(configs.get("edit_channel_reply")?.inputSchema).not.toHaveProperty("subject");
  });
});

describe("attachments in message projection", () => {
  it("returns a parsed Adaptive Card and a file reference from list_channel_messages", async () => {
    const card = { type: "AdaptiveCard", body: [{ type: "TextBlock", text: "Build failed" }] };
    useGraphClient({
      response: {
        value: [
          {
            id: "m1",
            createdDateTime: "2026-01-01T00:00:00Z",
            body: { contentType: "html", content: '<attachment id="a1"></attachment>' },
            attachments: [
              { id: "a1", contentType: "application/vnd.microsoft.card.adaptive", content: JSON.stringify(card) },
              { id: "a2", contentType: "reference", name: "report.pdf", contentUrl: "https://example/report.pdf" },
              { id: "a3", contentType: "application/vnd.microsoft.card.hero", content: "not json" },
            ],
          },
        ],
      },
    });

    const result = await handlers.get("list_channel_messages")!({ teamId: "t1", channelId: "c1", includeReplies: false });

    const [a1, a2, a3] = JSON.parse(result.content[0].text).messages[0].attachments;
    expect(a1.content).toEqual(card);
    expect(a2).toMatchObject({ name: "report.pdf", contentUrl: "https://example/report.pdf" });
    expect(a2).not.toHaveProperty("content");
    expect(a3.content).toBe("not json");
  });

  it("returns an empty attachments array when the message has none", async () => {
    useGraphClient({ response: { value: [{ id: "m1", createdDateTime: "2026-01-01T00:00:00Z", body: { content: "hi" } }] } });

    const result = await handlers.get("list_channel_messages")!({ teamId: "t1", channelId: "c1", includeReplies: false });

    expect(JSON.parse(result.content[0].text).messages[0].attachments).toEqual([]);
  });
});

describe("attachments in other read paths", () => {
  const card = { type: "AdaptiveCard", body: [{ type: "TextBlock", text: "Deploy done" }] };
  const cardAttachment = {
    id: "a1",
    contentType: "application/vnd.microsoft.card.adaptive",
    content: JSON.stringify(card),
  };

  it("get_channel_thread returns a Card carried by a Reply", async () => {
    useGraphClient({
      responses: {
        "/teams/t1/channels/c1/messages/root1": {
          response: { id: "root1", replyToId: null, createdDateTime: "2024-01-01T00:00:00Z", body: { content: "q" } },
        },
        "/teams/t1/channels/c1/messages/root1/replies?$top=50": {
          response: {
            value: [
              {
                id: "r1",
                replyToId: "root1",
                createdDateTime: "2024-01-01T01:00:00Z",
                body: { content: '<attachment id="a1"></attachment>' },
                attachments: [cardAttachment],
              },
            ],
          },
        },
      },
    });

    const result = await handlers.get("get_channel_thread")!({ teamId: "t1", channelId: "c1", messageId: "root1" });
    const parsed = JSON.parse(result.content[0].text);

    expect(parsed.root.attachments).toEqual([]);
    expect(parsed.replies[0].attachments[0].content).toEqual(card);
  });

  it("list_channel_messages returns a Card carried by a Reply preview", async () => {
    useGraphClient({
      response: {
        value: [
          {
            id: "root1",
            createdDateTime: "2024-01-01T00:00:00Z",
            body: { content: "q" },
            replies: [
              {
                id: "r1",
                replyToId: "root1",
                createdDateTime: "2024-01-01T01:00:00Z",
                body: { content: "" },
                attachments: [cardAttachment],
              },
            ],
          },
        ],
      },
    });

    const result = await handlers.get("list_channel_messages")!({ teamId: "t1", channelId: "c1" });

    expect(JSON.parse(result.content[0].text).messages[0].replies[0].attachments[0].content).toEqual(card);
  });

  it("search_messages reports attachments as null (unknown), never as an empty array", async () => {
    useGraphClient({
      response: {
        value: [
          {
            hitsContainers: [
              {
                hits: [
                  {
                    summary: "Card - access it on ...",
                    resource: {
                      id: "m1",
                      createdDateTime: "2024-01-01T00:00:00Z",
                      body: { content: "" },
                      channelIdentity: { teamId: "t1", channelId: "c1" },
                    },
                  },
                ],
              },
            ],
          },
        ],
      },
    });

    const result = await handlers.get("search_messages")!({ query: "x" });

    expect(JSON.parse(result.content[0].text).results[0].message.attachments).toBeNull();
  });
});
