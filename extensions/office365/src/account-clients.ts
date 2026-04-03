import type { PluginLogger } from "openclaw/plugin-sdk";
import type { Office365Config } from "./types.js";
import { isMultiAccountMode, listOffice365AccountIds, resolveOffice365Account } from "./accounts.js";
import { createGraphClient, type GraphClient } from "./graph-client.js";

/**
 * Creates a Map of accountId → GraphClient.
 *
 * - Legacy mode (no `accounts` in config): returns `Map { "default" => singleClient }`
 * - Multi-account mode: returns one client per configured account, each with its own
 *   credential file and token cache.
 */
export function createAccountClients(params: {
  config: Office365Config;
  stateDir: string;
  logger: PluginLogger;
}): Map<string, GraphClient> {
  const { config, stateDir, logger } = params;
  const clients = new Map<string, GraphClient>();

  if (!isMultiAccountMode(config)) {
    // Legacy: single client with default credential path
    clients.set(
      "default",
      createGraphClient({ config, stateDir, logger }),
    );
    return clients;
  }

  // Multi-account: one client per account with merged config
  for (const accountId of listOffice365AccountIds(config)) {
    const resolved = resolveOffice365Account({ config, accountId });
    clients.set(
      accountId,
      createGraphClient({
        config: resolved.config,
        stateDir,
        logger,
        accountId,
      }),
    );
  }

  return clients;
}
