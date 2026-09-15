import { describe, it, expect } from "vitest";
import { createFakeGraphClient } from "./fake-graph-client.js";

describe("createFakeGraphClient", () => {
  it("answers different paths with different configured responses", async () => {
    const fake = createFakeGraphClient({
      responses: {
        "/a": { response: { from: "a" } },
        "/b": { response: { from: "b" } },
      },
    });

    const a = await (fake.client as any).api("/a").get();
    const b = await (fake.client as any).api("/b").get();

    expect(a).toEqual({ from: "a" });
    expect(b).toEqual({ from: "b" });
  });

  it("still supports a single response used for every call", async () => {
    const fake = createFakeGraphClient({ response: { ok: true } });

    const a = await (fake.client as any).api("/a").get();
    const b = await (fake.client as any).api("/b").post({});

    expect(a).toEqual({ ok: true });
    expect(b).toEqual({ ok: true });
    expect(fake.requests).toHaveLength(2);
  });

  it("throws per-path configured errors", async () => {
    const fake = createFakeGraphClient({
      responses: {
        "/ok": { response: { fine: true } },
        "/bad": { error: { code: "ItemNotFound", message: "nope" } },
      },
    });

    await expect((fake.client as any).api("/bad").get()).rejects.toMatchObject({ code: "ItemNotFound" });
  });

  it("fails loudly, naming the path, when a path has no configured response", async () => {
    const fake = createFakeGraphClient({
      responses: {
        "/known": { response: {} },
      },
    });

    await expect((fake.client as any).api("/unknown").get()).rejects.toThrow(/\/unknown/);
  });
});
