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

export function createFakeGraphClient(config: { response?: unknown; error?: unknown } = {}): FakeGraphClient {
  const requests: RecordedRequest[] = [];

  function api(path: string) {
    const headers: Record<string, string> = {};
    let search: string | undefined;

    function record(method: RecordedRequest["method"], body: unknown) {
      requests.push({ path, method, body, headers, search });
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
