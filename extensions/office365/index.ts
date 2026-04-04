import { existsSync } from "node:fs";
import { join } from "node:path";
import {
  buildOauthProviderAuthResult,
  type OpenClawPluginApi,
  type ProviderAuthContext,
} from "openclaw/plugin-sdk";
import { loginMicrosoftOAuth, refreshMicrosoftTokens } from "./src/oauth.js";
import type { GraphClient } from "./src/graph-client.js";
import { createAccountClients } from "./src/account-clients.js";
import { createEmailListTool } from "./src/tools/email-list.js";
import { createEmailReadTool } from "./src/tools/email-read.js";
import { createEmailSendTool } from "./src/tools/email-send.js";
import { createEmailReplyTool } from "./src/tools/email-reply.js";
import { createEmailSearchTool } from "./src/tools/email-search.js";
import { GraphApiError } from "./src/types.js";
import type { Office365AccountConfig, Office365Config } from "./src/types.js";
import {
  isMultiAccountMode,
  listOffice365AccountIds,
  resolveAccountForTool,
  resolveOffice365Account,
  isToolPermittedForAccount,
  validateAccountsConfig,
} from "./src/accounts.js";

// ── Constants ───────────────────────────────────────────────────────────────

const PROVIDER_ID = "microsoft-graph";
const DEFAULT_SCOPES = [
  "Mail.Read",
  "Mail.Send",
  "User.Read",
  "offline_access",
];
const DEFAULT_REDIRECT_URI = "http://localhost:8080/callback";

// ── Config parsing ──────────────────────────────────────────────────────────

function parseConfig(raw: Record<string, unknown> | undefined): Office365Config {
  const cfg = raw && typeof raw === "object" && !Array.isArray(raw) ? raw : {};

  const str = (val: unknown, envFallback: string | undefined): string => {
    if (typeof val === "string") return val;
    if (typeof envFallback === "string") return envFallback;
    return "";
  };

  const config: Office365Config = {
    clientId: str(cfg.clientId, process.env.OFFICE365_CLIENT_ID),
    tenantId: str(cfg.tenantId, process.env.OFFICE365_TENANT_ID),
    clientSecret: str(cfg.clientSecret, process.env.OFFICE365_CLIENT_SECRET),
    redirectUri: str(cfg.redirectUri, process.env.OFFICE365_REDIRECT_URI) || DEFAULT_REDIRECT_URI,
    scopes: Array.isArray(cfg.scopes)
      ? cfg.scopes.filter((s): s is string => typeof s === "string")
      : DEFAULT_SCOPES,
  };

  // Pass through multi-account fields
  if (typeof cfg.defaultAccount === "string") {
    config.defaultAccount = cfg.defaultAccount;
  }
  if (cfg.accounts && typeof cfg.accounts === "object" && !Array.isArray(cfg.accounts)) {
    config.accounts = cfg.accounts as Record<string, Office365AccountConfig>;
  }

  return config;
}

const office365ConfigSchema = {
  parse(value: unknown) {
    return parseConfig(value as Record<string, unknown> | undefined);
  },
  uiHints: {
    clientId: { label: "Azure Client ID" },
    tenantId: { label: "Azure Tenant ID" },
    clientSecret: { label: "Azure Client Secret", sensitive: true },
    redirectUri: { label: "OAuth Redirect URI" },
    scopes: { label: "Graph API Scopes", advanced: true },
  },
};

// ── Plugin ──────────────────────────────────────────────────────────────────

const office365Plugin = {
  id: "office365",
  name: "Microsoft 365",
  description: "Email and calendar tools via Microsoft Graph API",
  configSchema: office365ConfigSchema,
  register(api: OpenClawPluginApi) {
    const config = parseConfig(api.pluginConfig as Record<string, unknown>);
    const stateDir = api.runtime.state.resolveStateDir();

    // ── Startup validation ────────────────────────────────────────────────

    if (isMultiAccountMode(config)) {
      const errors = validateAccountsConfig(config);
      if (errors.length > 0) {
        throw new Error(
          `Office365 multi-account config invalid:\n  - ${errors.join("\n  - ")}`,
        );
      }
    }

    // ── Client creation ───────────────────────────────────────────────────

    const clients = createAccountClients({ config, stateDir, logger: api.logger });

    // ── resolveClient closure (used by tools in multi-account mode) ───────

    function resolveClient(toolName: string, accountId?: string): GraphClient {
      // Determine which account to use
      const targetId = accountId ?? resolveAccountForTool(toolName, config);

      // Validate account exists
      if (!clients.has(targetId)) {
        const available = [...clients.keys()].join(", ");
        throw new GraphApiError(
          `Unknown account '${targetId}'. Available accounts: ${available}.`,
          "user_input",
          400,
        );
      }

      // Policy check
      if (isMultiAccountMode(config) && !isToolPermittedForAccount(toolName, targetId, config)) {
        const permitted = listOffice365AccountIds(config).filter(
          (id) => isToolPermittedForAccount(toolName, id, config),
        );
        throw new GraphApiError(
          `Tool ${toolName} is not permitted for account '${targetId}'. Allowed accounts: [${permitted.map((a) => `'${a}'`).join(", ")}].`,
          "user_input",
          403,
        );
      }

      return clients.get(targetId)!;
    }

    // ── Provider + tool registration ─────────────────────────────────────

    if (isMultiAccountMode(config)) {
      // Multi-account: register per-account OAuth providers
      for (const accountId of listOffice365AccountIds(config)) {
        const resolved = resolveOffice365Account({ config, accountId });
        const providerId = `${PROVIDER_ID}-${accountId}`;
        const accountClient = clients.get(accountId)!;

        api.registerProvider({
          id: providerId,
          label: `Microsoft Graph (${resolved.name})`,
          aliases: [`office365-${accountId}`, `o365-${accountId}`],
          auth: [
            {
              id: "oauth",
              label: `OAuth 2.0 (PKCE) — ${resolved.name}`,
              hint: `Authorization code flow for ${resolved.email ?? accountId}`,
              kind: "oauth",
              run: async (ctx: ProviderAuthContext) => {
                if (!resolved.config.clientId || !resolved.config.tenantId) {
                  throw new Error(
                    `Missing clientId or tenantId for account '${accountId}'.`,
                  );
                }

                const spin = ctx.prompter.progress(
                  `Starting Microsoft 365 OAuth for ${resolved.name}…`,
                );
                try {
                  const result = await loginMicrosoftOAuth(
                    {
                      isRemote: ctx.isRemote,
                      openUrl: ctx.openUrl,
                      log: (msg) => ctx.runtime.log(msg),
                      note: ctx.prompter.note,
                      prompt: async (message) =>
                        String(await ctx.prompter.text({ message })),
                      progress: spin,
                    },
                    resolved.config,
                  );

                  accountClient.setCredential(result);

                  spin.stop(`Microsoft 365 OAuth complete for ${resolved.name}`);
                  return buildOauthProviderAuthResult({
                    providerId,
                    defaultModel: "",
                    access: result.access,
                    refresh: result.refresh,
                    expires: result.expires,
                    email: result.email,
                    credentialExtra: {
                      clientId: resolved.config.clientId,
                      tenantId: resolved.config.tenantId,
                      accountId,
                    },
                    configPatch: {},
                    notes: [
                      `Microsoft 365 account '${resolved.name}' authenticated. Tools: ${resolved.tools.join(", ") || "all"}.`,
                    ],
                  });
                } catch (err) {
                  spin.stop(`Microsoft 365 OAuth failed for ${resolved.name}`);
                  await ctx.prompter.note(
                    `If you're having trouble with ${resolved.name}, ensure your Azure app registration has the correct redirect URI and API permissions for scopes: ${resolved.config.scopes?.join(", ") ?? "(default)"}.`,
                    "OAuth help",
                  );
                  throw err;
                }
              },
            },
          ],
          refreshOAuth: async (cred) => {
            if (!cred.refresh) {
              throw new Error(
                `No refresh token for account '${accountId}'. Run \`openclaw auth\` and select ${providerId}.`,
              );
            }
            const refreshed = await refreshMicrosoftTokens(cred.refresh, resolved.config);
            return {
              ...cred,
              access: refreshed.access,
              refresh: refreshed.refresh,
              expires: refreshed.expires,
            };
          },
        });
      }

      // Register tools with resolveClient
      const toolDeps = { graphClient: clients.values().next().value!, resolveClient };
      api.registerTool(createEmailListTool(toolDeps));
      api.registerTool(createEmailReadTool(toolDeps));
      api.registerTool(createEmailSendTool(toolDeps));
      api.registerTool(createEmailReplyTool(toolDeps));
      api.registerTool(createEmailSearchTool(toolDeps));

      // Multi-account prompt context
      api.on("before_prompt_build", async () => {
        const lines: string[] = [];
        for (const accountId of listOffice365AccountIds(config)) {
          const resolved = resolveOffice365Account({ config, accountId });
          const credFile = accountId === "default"
            ? join(stateDir, "office365-credentials.json")
            : join(stateDir, `office365-credentials-${accountId}.json`);
          const authed = existsSync(credFile);
          const toolsList = resolved.tools.length > 0
            ? resolved.tools.join(", ")
            : "all";
          if (authed) {
            lines.push(`${resolved.name} (${accountId}): authenticated (${toolsList})`);
          } else {
            lines.push(`${resolved.name} (${accountId}): not authenticated — run \`openclaw auth\` and select microsoft-graph-${accountId}`);
          }
        }
        return {
          prependContext: `Microsoft 365 multi-account mode:\n${lines.join("\n")}`,
        };
      });
    } else {
      // ── Legacy single-account mode ─────────────────────────────────────

      const graphClient = clients.get("default")!;

      api.registerProvider({
        id: PROVIDER_ID,
        label: "Microsoft Graph",
        aliases: ["office365", "o365"],
        auth: [
          {
            id: "oauth",
            label: "OAuth 2.0 (PKCE)",
            hint: "Authorization code flow with PKCE + localhost callback",
            kind: "oauth",
            run: async (ctx: ProviderAuthContext) => {
              if (!config.clientId || !config.tenantId) {
                throw new Error(
                  "Missing clientId or tenantId. Set OFFICE365_CLIENT_ID and OFFICE365_TENANT_ID environment variables, or configure in plugins.entries.office365.config.",
                );
              }

              const spin = ctx.prompter.progress(
                "Starting Microsoft 365 OAuth…",
              );
              try {
                const result = await loginMicrosoftOAuth(
                  {
                    isRemote: ctx.isRemote,
                    openUrl: ctx.openUrl,
                    log: (msg) => ctx.runtime.log(msg),
                    note: ctx.prompter.note,
                    prompt: async (message) =>
                      String(await ctx.prompter.text({ message })),
                    progress: spin,
                  },
                  config,
                );

                graphClient.setCredential(result);

                spin.stop("Microsoft 365 OAuth complete");
                return buildOauthProviderAuthResult({
                  providerId: PROVIDER_ID,
                  defaultModel: "",
                  access: result.access,
                  refresh: result.refresh,
                  expires: result.expires,
                  email: result.email,
                  credentialExtra: {
                    clientId: config.clientId,
                    tenantId: config.tenantId,
                  },
                  configPatch: {},
                  notes: [
                    "Microsoft 365 email tools are now available: email_list, email_read, email_send, email_reply, email_search.",
                  ],
                });
              } catch (err) {
                spin.stop("Microsoft 365 OAuth failed");
                await ctx.prompter.note(
                  "If you're having trouble, ensure your Azure app registration has the correct redirect URI (http://localhost:8080/callback) and API permissions.",
                  "OAuth help",
                );
                throw err;
              }
            },
          },
        ],
        refreshOAuth: async (cred) => {
          if (!cred.refresh) {
            throw new Error(
              "No refresh token. Run `openclaw auth` to re-authenticate with Microsoft 365.",
            );
          }
          const refreshed = await refreshMicrosoftTokens(cred.refresh, config);
          return {
            ...cred,
            access: refreshed.access,
            refresh: refreshed.refresh,
            expires: refreshed.expires,
          };
        },
      });

      api.registerTool(createEmailListTool({ graphClient }));
      api.registerTool(createEmailReadTool({ graphClient }));
      api.registerTool(createEmailSendTool({ graphClient }));
      api.registerTool(createEmailReplyTool({ graphClient }));
      api.registerTool(createEmailSearchTool({ graphClient }));

      const credentialFile = join(stateDir, "office365-credentials.json");

      api.on("before_prompt_build", async () => {
        if (!existsSync(credentialFile)) {
          return {};
        }
        return {
          prependContext:
            "Microsoft 365 email tools are available: email_list (list/search messages), email_read (read full message by ID), email_send (compose and send new email), email_reply (reply to existing email), email_search (structured search across all folders with date ranges, sender/recipient/subject filters).",
        };
      });
    }
  },
};

export { parseConfig, DEFAULT_SCOPES };
export default office365Plugin;
