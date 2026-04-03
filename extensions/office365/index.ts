import { existsSync } from "node:fs";
import { join } from "node:path";
import {
  buildOauthProviderAuthResult,
  type OpenClawPluginApi,
  type ProviderAuthContext,
} from "openclaw/plugin-sdk";
import { loginMicrosoftOAuth, refreshMicrosoftTokens } from "./src/oauth.js";
import { createGraphClient } from "./src/graph-client.js";
import { createEmailListTool } from "./src/tools/email-list.js";
import { createEmailReadTool } from "./src/tools/email-read.js";
import { createEmailSendTool } from "./src/tools/email-send.js";
import { createEmailReplyTool } from "./src/tools/email-reply.js";
import type { Office365Config } from "./src/types.js";

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

  return {
    clientId: str(cfg.clientId, process.env.OFFICE365_CLIENT_ID),
    tenantId: str(cfg.tenantId, process.env.OFFICE365_TENANT_ID),
    clientSecret: str(cfg.clientSecret, process.env.OFFICE365_CLIENT_SECRET),
    redirectUri: str(cfg.redirectUri, process.env.OFFICE365_REDIRECT_URI) || DEFAULT_REDIRECT_URI,
    scopes: Array.isArray(cfg.scopes)
      ? cfg.scopes.filter((s): s is string => typeof s === "string")
      : DEFAULT_SCOPES,
  };
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
    const graphClient = createGraphClient({
      config,
      stateDir,
      logger: api.logger,
    });

    // ── Provider registration (OAuth flow) ────────────────────────────────

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

              // Store credential in plugin's own state dir
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
                  "Microsoft 365 email tools are now available: email_list, email_read, email_send, email_reply.",
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

    // ── Tool registration ─────────────────────────────────────────────────

    api.registerTool(createEmailListTool({ graphClient }));
    api.registerTool(createEmailReadTool({ graphClient }));
    api.registerTool(createEmailSendTool({ graphClient }));
    api.registerTool(createEmailReplyTool({ graphClient }));

    // ── Prompt context ────────────────────────────────────────────────────

    const credentialFile = join(stateDir, "office365-credentials.json");

    api.on("before_prompt_build", async () => {
      if (!existsSync(credentialFile)) {
        return {};
      }
      return {
        prependContext:
          "Microsoft 365 email tools are available: email_list (list/search messages), email_read (read full message by ID), email_send (compose and send new email), email_reply (reply to existing email).",
      };
    });
  },
};

export { parseConfig, DEFAULT_SCOPES };
export default office365Plugin;
