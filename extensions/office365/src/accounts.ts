import type { Office365Config, ResolvedOffice365Account } from "./types.js";

// ── Constants ──────────────────────────────────────────────────────────────

export const ACCOUNT_ID_RE = /^[a-z0-9][a-z0-9_-]*$/;

const KNOWN_TOOLS = new Set([
  "email_list",
  "email_read",
  "email_send",
  "email_reply",
  "email_search",
  "calendar_list",
  "calendar_create",
  "calendar_update",
  "calendar_delete",
]);

// ── Mode detection ─────────────────────────────────────────────────────────

export function isMultiAccountMode(config: Office365Config): boolean {
  return !!config.accounts && Object.keys(config.accounts).length > 0;
}

// ── Account listing ────────────────────────────────────────────────────────

export function listOffice365AccountIds(config: Office365Config): string[] {
  if (!isMultiAccountMode(config)) {
    return ["default"];
  }
  return Object.keys(config.accounts!).sort();
}

// ── Default account resolution ─────────────────────────────────────────────

export function resolveDefaultOffice365AccountId(
  config: Office365Config,
): string {
  if (!isMultiAccountMode(config)) {
    return "default";
  }
  const ids = listOffice365AccountIds(config);
  if (config.defaultAccount && ids.includes(config.defaultAccount)) {
    return config.defaultAccount;
  }
  return ids[0] ?? "default";
}

// ── Account resolution (merge account config with top-level) ───────────────

export function resolveOffice365Account(params: {
  config: Office365Config;
  accountId: string;
}): ResolvedOffice365Account {
  const { config, accountId } = params;
  const accountCfg = config.accounts?.[accountId];

  // Merge: account-level overrides top-level
  const merged: Office365Config = {
    clientId: accountCfg?.clientId || config.clientId,
    tenantId: accountCfg?.tenantId || config.tenantId,
    clientSecret: accountCfg?.clientSecret || config.clientSecret,
    redirectUri: accountCfg?.redirectUri || config.redirectUri,
    scopes: accountCfg?.scopes ?? config.scopes,
  };

  return {
    accountId,
    name: accountCfg?.name ?? accountId,
    email: accountCfg?.email,
    config: merged,
    tools: accountCfg?.tools ?? [],
  };
}

// ── Tool → account mapping ─────────────────────────────────────────────────

export function resolveAccountForTool(
  toolName: string,
  config: Office365Config,
): string {
  if (!isMultiAccountMode(config)) {
    return "default";
  }
  for (const [accountId, accountCfg] of Object.entries(config.accounts!)) {
    if (accountCfg.tools?.includes(toolName)) {
      return accountId;
    }
  }
  return resolveDefaultOffice365AccountId(config);
}

// ── Tool permission check ──────────────────────────────────────────────────

export function isToolPermittedForAccount(
  toolName: string,
  accountId: string,
  config: Office365Config,
): boolean {
  if (!isMultiAccountMode(config)) {
    return true; // Legacy mode: all tools permitted
  }
  const accountCfg = config.accounts?.[accountId];
  if (!accountCfg) {
    return false;
  }
  // If account has no tools restriction, allow all
  if (!accountCfg.tools || accountCfg.tools.length === 0) {
    return true;
  }
  return accountCfg.tools.includes(toolName);
}

// ── Config validation ──────────────────────────────────────────────────────

export function validateAccountsConfig(config: Office365Config): string[] {
  const errors: string[] = [];

  if (!config.accounts) {
    return errors; // Legacy mode — no validation needed
  }

  const accountIds = Object.keys(config.accounts);

  // Reject empty accounts object
  if (accountIds.length === 0) {
    errors.push("accounts object is empty. Remove it for single-account mode or add at least one account.");
    return errors;
  }

  // Validate account IDs
  for (const id of accountIds) {
    if (!ACCOUNT_ID_RE.test(id)) {
      errors.push(
        `Invalid account ID '${id}'. Must match ${ACCOUNT_ID_RE} (lowercase alphanumeric, hyphens, underscores).`,
      );
    }
  }

  // defaultAccount must exist in accounts
  if (config.defaultAccount && !accountIds.includes(config.defaultAccount)) {
    errors.push(
      `defaultAccount '${config.defaultAccount}' does not exist in accounts. Available: ${accountIds.join(", ")}.`,
    );
  }

  // Validate each account
  const toolOwnership = new Map<string, string>();
  for (const [accountId, accountCfg] of Object.entries(config.accounts)) {
    // Require email in multi-account mode
    if (!accountCfg.email) {
      errors.push(
        `Account '${accountId}' is missing required 'email' field (needed for login_hint in OAuth).`,
      );
    }

    // Require scopes in multi-account mode
    if (!accountCfg.scopes || accountCfg.scopes.length === 0) {
      errors.push(
        `Account '${accountId}' is missing required 'scopes' field.`,
      );
    }

    // Validate tool names and check for duplicates
    for (const tool of accountCfg.tools ?? []) {
      if (!KNOWN_TOOLS.has(tool)) {
        errors.push(
          `Account '${accountId}' references unknown tool '${tool}'. Known tools: ${[...KNOWN_TOOLS].join(", ")}.`,
        );
      }
      const existingOwner = toolOwnership.get(tool);
      if (existingOwner) {
        errors.push(
          `Tool '${tool}' is owned by both '${existingOwner}' and '${accountId}'. Each tool must belong to exactly one account.`,
        );
      } else {
        toolOwnership.set(tool, accountId);
      }
    }
  }

  return errors;
}
