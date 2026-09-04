import { Client } from "@microsoft/microsoft-graph-client";
import { getAccessToken, isTokenExpired, refreshAccessToken, loadTokenFromFile } from "../auth/token-manager.js";
import { log } from "../utils/logger.js";

export async function getGraphClient(): Promise<Client> {
  if (!getAccessToken()) {
    await loadTokenFromFile();
  }

  if (!getAccessToken()) {
    throw new Error("Not authenticated. Please use auth_start if you received an authentication error, or set GRAPH_ACCESS_TOKEN environment variable.");
  }

  if (isTokenExpired()) {
    log("INFO", "Access token expired, attempting to refresh");
    const refreshed = await refreshAccessToken();
    if (!refreshed || !getAccessToken()) {
      throw new Error("Access token expired and refresh failed. Please re-authenticate using auth_start.");
    }
  }

  const client = Client.init({
    authProvider: (done) => {
      done(null, getAccessToken()!);
    },
  });

  return client;
}
