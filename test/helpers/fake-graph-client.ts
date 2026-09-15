import type { Client } from "@microsoft/microsoft-graph-client";

export interface RecordedRequest {
  path: string;
  method: "get" | "post" | "patch" | "delete";
  body?: unknown;
  headers: Record<string, string>;
  search?: string;
}

export interface FakeGraphClient {
  client: Client;
  requests: RecordedRequest[];
}

export interface FakeGraphClientConfig {
  /** Fallback used for any path not found in `responses`, or for every call when `responses` is omitted. */
  response?: unknown;
  error?: unknown;
  /** Per-path responses, keyed by the exact string passed to `client.api(...)`. Takes precedence over the fallback. */
  responses?: Record<string, { response?: unknown; error?: unknown }>;
}

export function createFakeGraphClient(config: FakeGraphClientConfig = {}): FakeGraphClient {
  const requests: RecordedRequest[] = [];

  function api(path: string) {
    const headers: Record<string, string> = {};
    let search: string | undefined;

    function record(method: RecordedRequest["method"], body: unknown) {
      requests.push({ path, method, body, headers, search });

      const configured = config.responses?.[path];
      if (configured) {
        if (configured.error) throw configured.error;
        return configured.response ?? {};
      }

      if (config.responses && config.error === undefined && config.response === undefined) {
        throw new Error(`No fake Graph response configured for path: ${path}`);
      }

      if (config.error) throw config.error;
      return config.response ?? {};
    }

    const request: any = {
      header(name: string, value: string) {
        headers[name] = value;
        return request;
      },
      search(query: string) {
        search = query;
        return request;
      },
      select(_fields: string) {
        return request;
      },
      get: async () => record("get", undefined),
      post: async (body?: unknown) => record("post", body),
      patch: async (body?: unknown) => record("patch", body),
      delete: async () => record("delete", undefined),
    };

    return request;
  }

  return {
    client: { api } as unknown as Client,
    requests,
  };
}
