import { describe, expect, it } from "vitest";
import {
  isMultiAccountMode,
  listOffice365AccountIds,
  resolveDefaultOffice365AccountId,
  resolveOffice365Account,
  resolveAccountForTool,
  isToolPermittedForAccount,
  validateAccountsConfig,
  ACCOUNT_ID_RE,
} from "./src/accounts.js";
import type { Office365Config } from "./src/types.js";

// ── Test helpers ───────────────────────────────────────────────────────────

const LEGACY_CONFIG: Office365Config = {
  clientId: "test-client",
  tenantId: "test-tenant",
  clientSecret: "test-secret",
  redirectUri: "http://localhost:8080/callback",
  scopes: ["Mail.Read", "Mail.Send", "User.Read", "offline_access"],
};

const MULTI_CONFIG: Office365Config = {
  clientId: "test-client",
  tenantId: "test-tenant",
  clientSecret: "test-secret",
  redirectUri: "http://localhost:8080/callback",
  scopes: ["User.Read", "offline_access"],
  defaultAccount: "rod",
  accounts: {
    rod: {
      name: "Rod (Read Only)",
      email: "Rod@CerebusTechnologies.onmicrosoft.com",
      scopes: ["Mail.Read", "User.Read", "offline_access"],
      tools: ["email_list", "email_read"],
    },
    openclaw: {
      name: "OpenClaw (Sender)",
      email: "openclaw@cerebustechnologies.com",
      scopes: ["Mail.Send", "Mail.Read", "User.Read", "offline_access"],
      tools: ["email_send", "email_reply"],
    },
  },
};

// ── isMultiAccountMode ─────────────────────────────────────────────────────

describe("isMultiAccountMode", () => {
  it("returns false for legacy config without accounts", () => {
    expect(isMultiAccountMode(LEGACY_CONFIG)).toBe(false);
  });

  it("returns true when accounts is populated", () => {
    expect(isMultiAccountMode(MULTI_CONFIG)).toBe(true);
  });

  it("returns false for empty accounts object", () => {
    expect(isMultiAccountMode({ ...LEGACY_CONFIG, accounts: {} })).toBe(false);
  });
});

// ── listOffice365AccountIds ────────────────────────────────────────────────

describe("listOffice365AccountIds", () => {
  it("returns ['default'] for legacy config", () => {
    expect(listOffice365AccountIds(LEGACY_CONFIG)).toEqual(["default"]);
  });

  it("returns sorted account IDs for multi-account config", () => {
    expect(listOffice365AccountIds(MULTI_CONFIG)).toEqual(["openclaw", "rod"]);
  });
});

// ── resolveDefaultOffice365AccountId ───────────────────────────────────────

describe("resolveDefaultOffice365AccountId", () => {
  it("returns 'default' for legacy config", () => {
    expect(resolveDefaultOffice365AccountId(LEGACY_CONFIG)).toBe("default");
  });

  it("prefers defaultAccount when set and exists in accounts", () => {
    expect(resolveDefaultOffice365AccountId(MULTI_CONFIG)).toBe("rod");
  });

  it("falls back to first alphabetical account if defaultAccount is missing from accounts", () => {
    const config: Office365Config = {
      ...MULTI_CONFIG,
      defaultAccount: "nonexistent",
    };
    expect(resolveDefaultOffice365AccountId(config)).toBe("openclaw");
  });

  it("falls back to first alphabetical account if defaultAccount is not set", () => {
    const config: Office365Config = {
      ...MULTI_CONFIG,
      defaultAccount: undefined,
    };
    expect(resolveDefaultOffice365AccountId(config)).toBe("openclaw");
  });
});

// ── resolveOffice365Account ────────────────────────────────────────────────

describe("resolveOffice365Account", () => {
  it("merges account config with top-level defaults", () => {
    const resolved = resolveOffice365Account({ config: MULTI_CONFIG, accountId: "rod" });

    expect(resolved.accountId).toBe("rod");
    expect(resolved.name).toBe("Rod (Read Only)");
    expect(resolved.email).toBe("Rod@CerebusTechnologies.onmicrosoft.com");
    expect(resolved.tools).toEqual(["email_list", "email_read"]);
    // Merged config inherits top-level clientId/tenantId/clientSecret
    expect(resolved.config.clientId).toBe("test-client");
    expect(resolved.config.tenantId).toBe("test-tenant");
    expect(resolved.config.clientSecret).toBe("test-secret");
    // But uses account-specific scopes
    expect(resolved.config.scopes).toEqual(["Mail.Read", "User.Read", "offline_access"]);
  });

  it("uses top-level scopes as fallback when account omits scopes", () => {
    const config: Office365Config = {
      ...MULTI_CONFIG,
      accounts: {
        ...MULTI_CONFIG.accounts,
        bare: { name: "Bare Account" },
      },
    };
    const resolved = resolveOffice365Account({ config, accountId: "bare" });
    expect(resolved.config.scopes).toEqual(config.scopes);
  });

  it("allows account-specific config to override top-level credentials", () => {
    const config: Office365Config = {
      ...MULTI_CONFIG,
      accounts: {
        ...MULTI_CONFIG.accounts,
        custom: {
          name: "Custom",
          clientId: "override-client",
          tenantId: "override-tenant",
          scopes: ["Mail.Read"],
          tools: [],
        },
      },
    };
    const resolved = resolveOffice365Account({ config, accountId: "custom" });
    expect(resolved.config.clientId).toBe("override-client");
    expect(resolved.config.tenantId).toBe("override-tenant");
    // clientSecret not overridden, falls back to top-level
    expect(resolved.config.clientSecret).toBe("test-secret");
  });

  it("returns default account for legacy config", () => {
    const resolved = resolveOffice365Account({ config: LEGACY_CONFIG, accountId: "default" });
    expect(resolved.accountId).toBe("default");
    expect(resolved.config.clientId).toBe("test-client");
    expect(resolved.config.scopes).toEqual(LEGACY_CONFIG.scopes);
  });
});

// ── resolveAccountForTool ──────────────────────────────────────────────────

describe("resolveAccountForTool", () => {
  it("maps email_list to rod account", () => {
    expect(resolveAccountForTool("email_list", MULTI_CONFIG)).toBe("rod");
  });

  it("maps email_send to openclaw account", () => {
    expect(resolveAccountForTool("email_send", MULTI_CONFIG)).toBe("openclaw");
  });

  it("returns defaultAccount for unclaimed tools", () => {
    expect(resolveAccountForTool("unknown_tool", MULTI_CONFIG)).toBe("rod");
  });

  it("returns 'default' for legacy config", () => {
    expect(resolveAccountForTool("email_list", LEGACY_CONFIG)).toBe("default");
  });
});

// ── isToolPermittedForAccount ──────────────────────────────────────────────

describe("isToolPermittedForAccount", () => {
  it("allows email_list for rod", () => {
    expect(isToolPermittedForAccount("email_list", "rod", MULTI_CONFIG)).toBe(true);
  });

  it("denies email_send for rod", () => {
    expect(isToolPermittedForAccount("email_send", "rod", MULTI_CONFIG)).toBe(false);
  });

  it("allows email_send for openclaw", () => {
    expect(isToolPermittedForAccount("email_send", "openclaw", MULTI_CONFIG)).toBe(true);
  });

  it("denies email_list for openclaw", () => {
    expect(isToolPermittedForAccount("email_list", "openclaw", MULTI_CONFIG)).toBe(false);
  });

  it("allows any tool in legacy mode", () => {
    expect(isToolPermittedForAccount("email_send", "default", LEGACY_CONFIG)).toBe(true);
  });
});

// ── validateAccountsConfig ─────────────────────────────────────────────────

describe("validateAccountsConfig", () => {
  it("returns no errors for legacy config", () => {
    expect(validateAccountsConfig(LEGACY_CONFIG)).toEqual([]);
  });

  it("returns no errors for valid multi-account config", () => {
    expect(validateAccountsConfig(MULTI_CONFIG)).toEqual([]);
  });

  it("rejects when defaultAccount does not exist in accounts", () => {
    const config: Office365Config = {
      ...MULTI_CONFIG,
      defaultAccount: "nonexistent",
    };
    const errors = validateAccountsConfig(config);
    expect(errors.length).toBeGreaterThan(0);
    expect(errors[0]).toContain("defaultAccount");
  });

  it("rejects duplicate tool ownership across accounts", () => {
    const config: Office365Config = {
      ...MULTI_CONFIG,
      defaultAccount: "a",
      accounts: {
        a: { name: "A", email: "a@test.com", scopes: ["Mail.Read"], tools: ["email_list"] },
        b: { name: "B", email: "b@test.com", scopes: ["Mail.Read"], tools: ["email_list"] },
      },
    };
    const errors = validateAccountsConfig(config);
    expect(errors.some((e) => e.includes("email_list"))).toBe(true);
  });

  it("rejects unknown tool names", () => {
    const config: Office365Config = {
      ...MULTI_CONFIG,
      accounts: {
        rod: { name: "Rod", email: "r@test.com", scopes: ["Mail.Read"], tools: ["bogus_tool"] },
      },
    };
    const errors = validateAccountsConfig(config);
    expect(errors.length).toBeGreaterThan(0);
    expect(errors[0]).toContain("bogus_tool");
  });

  it("rejects empty accounts object", () => {
    const config: Office365Config = { ...LEGACY_CONFIG, accounts: {} };
    const errors = validateAccountsConfig(config);
    expect(errors.length).toBeGreaterThan(0);
  });

  it("rejects unsafe account IDs", () => {
    const config: Office365Config = {
      ...MULTI_CONFIG,
      defaultAccount: undefined,
      accounts: {
        "../etc/passwd": { name: "Evil", email: "e@test.com", scopes: ["Mail.Read"], tools: [] },
      },
    };
    const errors = validateAccountsConfig(config);
    expect(errors.length).toBeGreaterThan(0);
    expect(errors[0]).toContain("account ID");
  });

  it("rejects account IDs with spaces", () => {
    const config: Office365Config = {
      ...MULTI_CONFIG,
      defaultAccount: undefined,
      accounts: {
        "my account": { name: "Spaced", email: "s@test.com", scopes: ["Mail.Read"], tools: [] },
      },
    };
    const errors = validateAccountsConfig(config);
    expect(errors.length).toBeGreaterThan(0);
  });

  it("rejects account IDs with uppercase", () => {
    const config: Office365Config = {
      ...MULTI_CONFIG,
      defaultAccount: undefined,
      accounts: {
        "Rod": { name: "Upper", email: "r@test.com", scopes: ["Mail.Read"], tools: [] },
      },
    };
    const errors = validateAccountsConfig(config);
    expect(errors.length).toBeGreaterThan(0);
  });

  it("requires email for each account in multi-account mode", () => {
    const config: Office365Config = {
      ...MULTI_CONFIG,
      accounts: {
        rod: { name: "Rod", scopes: ["Mail.Read"], tools: ["email_list"] },
      },
    };
    const errors = validateAccountsConfig(config);
    expect(errors.length).toBeGreaterThan(0);
    expect(errors[0]).toContain("email");
  });

  it("requires scopes for each account in multi-account mode", () => {
    const config: Office365Config = {
      ...MULTI_CONFIG,
      accounts: {
        rod: { name: "Rod", email: "r@test.com", tools: ["email_list"] },
      },
    };
    const errors = validateAccountsConfig(config);
    expect(errors.length).toBeGreaterThan(0);
    expect(errors[0]).toContain("scopes");
  });
});

// ── ACCOUNT_ID_RE ──────────────────────────────────────────────────────────

describe("ACCOUNT_ID_RE", () => {
  it("accepts lowercase alphanumeric with hyphens and underscores", () => {
    expect(ACCOUNT_ID_RE.test("rod")).toBe(true);
    expect(ACCOUNT_ID_RE.test("open-claw")).toBe(true);
    expect(ACCOUNT_ID_RE.test("account_1")).toBe(true);
    expect(ACCOUNT_ID_RE.test("a-b_c-123")).toBe(true);
  });

  it("rejects unsafe patterns", () => {
    expect(ACCOUNT_ID_RE.test("../etc")).toBe(false);
    expect(ACCOUNT_ID_RE.test("my account")).toBe(false);
    expect(ACCOUNT_ID_RE.test("Rod")).toBe(false);
    expect(ACCOUNT_ID_RE.test("")).toBe(false);
    expect(ACCOUNT_ID_RE.test("a/b")).toBe(false);
  });
});

// ── Calendar tool routing ──────────────────────────────────────────────────

describe("calendar tool routing", () => {
  const CALENDAR_CONFIG: Office365Config = {
    clientId: "test-client",
    tenantId: "test-tenant",
    clientSecret: "test-secret",
    redirectUri: "http://localhost:8080/callback",
    scopes: ["User.Read", "offline_access"],
    defaultAccount: "rod",
    accounts: {
      rod: {
        name: "Rod",
        email: "Rod@CerebusTechnologies.onmicrosoft.com",
        scopes: ["Mail.Read", "Calendars.ReadWrite", "User.Read", "offline_access"],
        tools: ["email_list", "email_read", "email_search", "calendar_list", "calendar_create", "calendar_update", "calendar_delete"],
      },
      openclaw: {
        name: "OpenClaw (Sender)",
        email: "openclaw@cerebustechnologies.com",
        scopes: ["Mail.Send", "Mail.Read", "User.Read", "offline_access"],
        tools: ["email_send", "email_reply"],
      },
    },
  };

  it("resolveAccountForTool maps calendar tools to rod", () => {
    expect(resolveAccountForTool("calendar_list", CALENDAR_CONFIG)).toBe("rod");
    expect(resolveAccountForTool("calendar_create", CALENDAR_CONFIG)).toBe("rod");
    expect(resolveAccountForTool("calendar_update", CALENDAR_CONFIG)).toBe("rod");
    expect(resolveAccountForTool("calendar_delete", CALENDAR_CONFIG)).toBe("rod");
  });

  it("validateAccountsConfig accepts calendar tool names", () => {
    const errors = validateAccountsConfig(CALENDAR_CONFIG);
    expect(errors).toEqual([]);
  });

  it("isToolPermittedForAccount allows calendar tools for rod", () => {
    expect(isToolPermittedForAccount("calendar_list", "rod", CALENDAR_CONFIG)).toBe(true);
    expect(isToolPermittedForAccount("calendar_create", "rod", CALENDAR_CONFIG)).toBe(true);
  });

  it("isToolPermittedForAccount denies calendar tools for openclaw", () => {
    expect(isToolPermittedForAccount("calendar_list", "openclaw", CALENDAR_CONFIG)).toBe(false);
    expect(isToolPermittedForAccount("calendar_create", "openclaw", CALENDAR_CONFIG)).toBe(false);
  });

  it("validateAccountsConfig accepts email_attachment_read", () => {
    const config: Office365Config = {
      clientId: "c",
      tenantId: "t",
      clientSecret: "s",
      redirectUri: "http://localhost:8080/callback",
      scopes: ["Mail.Read"],
      accounts: {
        rod: {
          email: "rod@test.com",
          scopes: ["Mail.Read"],
          tools: ["email_list", "email_read", "email_attachment_read"],
        },
      },
    };
    const errors = validateAccountsConfig(config);
    expect(errors).toEqual([]);
  });

  it("isToolPermittedForAccount allows email_attachment_read for assigned account", () => {
    const config: Office365Config = {
      clientId: "c",
      tenantId: "t",
      clientSecret: "s",
      redirectUri: "http://localhost:8080/callback",
      scopes: ["Mail.Read"],
      accounts: {
        rod: {
          email: "rod@test.com",
          scopes: ["Mail.Read"],
          tools: ["email_attachment_read"],
        },
        openclaw: {
          email: "openclaw@test.com",
          scopes: ["Mail.Send"],
          tools: ["email_send"],
        },
      },
    };
    expect(isToolPermittedForAccount("email_attachment_read", "rod", config)).toBe(true);
    expect(isToolPermittedForAccount("email_attachment_read", "openclaw", config)).toBe(false);
  });
});
