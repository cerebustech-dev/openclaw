import { describe, expect, it, vi, beforeEach } from "vitest";
import { parseConfig, DEFAULT_SCOPES } from "./index.js";

vi.mock("node:fs", () => ({
  existsSync: vi.fn().mockReturnValue(false),
  writeFileSync: vi.fn(),
  readFileSync: vi.fn(),
  renameSync: vi.fn(),
  mkdirSync: vi.fn(),
  chmodSync: vi.fn(),
  unlinkSync: vi.fn(),
}));

vi.mock("node:crypto", async () => {
  const actual = await vi.importActual<typeof import("node:crypto")>("node:crypto");
  return {
    ...actual,
    randomBytes: (n: number) => Buffer.alloc(n, 0xab),
  };
});

vi.mock("openclaw/plugin-sdk", () => ({
  fetchWithSsrFGuard: vi.fn(),
  buildOauthProviderAuthResult: vi.fn().mockReturnValue({ providerId: "test" }),
}));

vi.mock("./src/oauth.js", () => ({
  refreshMicrosoftTokens: vi.fn(),
  loginMicrosoftOAuth: vi.fn(),
}));

// ── Issue 1: DEFAULT_SCOPES least-privilege ────────────────────────────────

describe("DEFAULT_SCOPES", () => {
  it("includes Phase 2 scopes (read + send email, user, offline)", () => {
    expect(DEFAULT_SCOPES).toEqual(["Mail.Read", "Mail.Send", "User.Read", "offline_access"]);
  });

  it("includes Mail.Send for Phase 2", () => {
    expect(DEFAULT_SCOPES).toContain("Mail.Send");
  });

  it("does not include Mail.ReadWrite", () => {
    expect(DEFAULT_SCOPES).not.toContain("Mail.ReadWrite");
  });

  it("does not include Calendars.ReadWrite", () => {
    expect(DEFAULT_SCOPES).not.toContain("Calendars.ReadWrite");
  });
});

// ── Issue 10: parseConfig type validation ──────────────────────────────────

describe("parseConfig", () => {
  it("returns valid config from string values", () => {
    const result = parseConfig({
      clientId: "abc",
      tenantId: "def",
      clientSecret: "ghi",
    });
    expect(result.clientId).toBe("abc");
    expect(result.tenantId).toBe("def");
    expect(result.clientSecret).toBe("ghi");
  });

  it("does not coerce undefined to the string 'undefined'", () => {
    const result = parseConfig({ clientId: undefined });
    expect(result.clientId).not.toBe("undefined");
    // Should fall back to env var or empty string
    expect(result.clientId).toBe(process.env.OFFICE365_CLIENT_ID ?? "");
  });

  it("does not coerce null to the string 'null'", () => {
    const result = parseConfig({ clientId: null });
    expect(result.clientId).not.toBe("null");
    expect(result.clientId).toBe(process.env.OFFICE365_CLIENT_ID ?? "");
  });

  it("does not coerce objects to '[object Object]'", () => {
    const result = parseConfig({ tenantId: {} });
    expect(result.tenantId).not.toBe("[object Object]");
    expect(result.tenantId).toBe(process.env.OFFICE365_TENANT_ID ?? "");
  });

  it("does not coerce numbers to strings for clientId", () => {
    const result = parseConfig({ clientId: 12345 });
    expect(result.clientId).not.toBe("12345");
    expect(result.clientId).toBe(process.env.OFFICE365_CLIENT_ID ?? "");
  });

  it("handles undefined raw config gracefully", () => {
    const result = parseConfig(undefined);
    expect(result.clientId).toBe(process.env.OFFICE365_CLIENT_ID ?? "");
    expect(result.tenantId).toBe(process.env.OFFICE365_TENANT_ID ?? "");
  });

  it("filters non-string elements from scopes array", () => {
    const result = parseConfig({
      scopes: ["Mail.Read", 123, null, undefined, {}, "User.Read"],
    });
    // Only actual strings should be kept
    expect(result.scopes).toEqual(["Mail.Read", "User.Read"]);
  });

  it("preserves accounts and defaultAccount from config", () => {
    const result = parseConfig({
      clientId: "abc",
      tenantId: "def",
      clientSecret: "ghi",
      defaultAccount: "rod",
      accounts: {
        rod: {
          name: "Rod (Read Only)",
          email: "Rod@test.com",
          scopes: ["Mail.Read", "User.Read", "offline_access"],
          tools: ["email_list", "email_read"],
        },
        openclaw: {
          name: "OpenClaw (Sender)",
          email: "openclaw@test.com",
          scopes: ["Mail.Send", "Mail.Read", "User.Read", "offline_access"],
          tools: ["email_send", "email_reply"],
        },
      },
    });
    expect(result.defaultAccount).toBe("rod");
    expect(result.accounts).toBeDefined();
    expect(Object.keys(result.accounts!)).toEqual(["rod", "openclaw"]);
    expect(result.accounts!.rod.email).toBe("Rod@test.com");
    expect(result.accounts!.openclaw.tools).toEqual(["email_send", "email_reply"]);
  });

  it("does not add accounts when not in raw config", () => {
    const result = parseConfig({ clientId: "abc", tenantId: "def" });
    expect(result.accounts).toBeUndefined();
    expect(result.defaultAccount).toBeUndefined();
  });

  it("account IDs with hyphens and underscores pass through", () => {
    const result = parseConfig({
      clientId: "abc",
      tenantId: "def",
      clientSecret: "ghi",
      accounts: {
        "my-account_1": {
          name: "Test",
          email: "test@test.com",
          scopes: ["Mail.Read"],
          tools: ["email_list"],
        },
      },
    });
    expect(result.accounts!["my-account_1"]).toBeDefined();
  });
});

// ── Phase 4: Multi-account registration ───────────────────────────────────

describe("office365Plugin.register", () => {
  function makeMockApi(rawConfig: Record<string, unknown>) {
    const registeredProviders: Array<{ id: string; label: string }> = [];
    const registeredTools: Array<{ name: string }> = [];
    const eventHandlers = new Map<string, Function>();

    return {
      api: {
        pluginConfig: rawConfig,
        runtime: {
          state: { resolveStateDir: () => "/tmp/test-state" },
          log: vi.fn(),
        },
        logger: {
          info: vi.fn(),
          warn: vi.fn(),
          debug: vi.fn(),
          error: vi.fn(),
        },
        registerProvider: vi.fn((p: { id: string; label: string }) => {
          registeredProviders.push(p);
        }),
        registerTool: vi.fn((t: { name: string }) => {
          registeredTools.push(t);
        }),
        on: vi.fn((event: string, handler: Function) => {
          eventHandlers.set(event, handler);
        }),
      },
      registeredProviders,
      registeredTools,
      eventHandlers,
    };
  }

  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("registers single provider in legacy mode", async () => {
    const { default: plugin } = await import("./index.js");
    const { api, registeredProviders, registeredTools } = makeMockApi({
      clientId: "test-client",
      tenantId: "test-tenant",
      clientSecret: "test-secret",
    });

    plugin.register(api as any);

    expect(registeredProviders).toHaveLength(1);
    expect(registeredProviders[0].id).toBe("microsoft-graph");
    expect(registeredTools).toHaveLength(5);
  });

  it("registers per-account providers in multi-account mode", async () => {
    const { default: plugin } = await import("./index.js");
    const { api, registeredProviders, registeredTools } = makeMockApi({
      clientId: "test-client",
      tenantId: "test-tenant",
      clientSecret: "test-secret",
      defaultAccount: "rod",
      accounts: {
        rod: {
          name: "Rod",
          email: "rod@test.com",
          scopes: ["Mail.Read", "User.Read", "offline_access"],
          tools: ["email_list", "email_read"],
        },
        openclaw: {
          name: "OpenClaw",
          email: "openclaw@test.com",
          scopes: ["Mail.Send", "Mail.Read", "User.Read", "offline_access"],
          tools: ["email_send", "email_reply"],
        },
      },
    });

    plugin.register(api as any);

    // Should register per-account providers with sanitized IDs
    const providerIds = registeredProviders.map((p) => p.id);
    expect(providerIds).toContain("microsoft-graph-rod");
    expect(providerIds).toContain("microsoft-graph-openclaw");
    // Should still register all 4 tools
    expect(registeredTools).toHaveLength(5);
  });

  it("throws at startup for invalid multi-account config", async () => {
    const { default: plugin } = await import("./index.js");
    const { api } = makeMockApi({
      clientId: "test-client",
      tenantId: "test-tenant",
      clientSecret: "test-secret",
      defaultAccount: "nonexistent",
      accounts: {
        rod: {
          name: "Rod",
          email: "rod@test.com",
          scopes: ["Mail.Read", "User.Read", "offline_access"],
          tools: ["email_list"],
        },
      },
    });

    expect(() => plugin.register(api as any)).toThrow("defaultAccount");
  });

  it("multi-account prompt context shows per-account status", async () => {
    const { default: plugin } = await import("./index.js");
    const { existsSync } = await import("node:fs");
    (existsSync as ReturnType<typeof vi.fn>).mockReturnValue(false);

    const { api, eventHandlers } = makeMockApi({
      clientId: "test-client",
      tenantId: "test-tenant",
      clientSecret: "test-secret",
      defaultAccount: "rod",
      accounts: {
        rod: {
          name: "Rod (Read Only)",
          email: "rod@test.com",
          scopes: ["Mail.Read", "User.Read", "offline_access"],
          tools: ["email_list", "email_read"],
        },
        openclaw: {
          name: "OpenClaw (Sender)",
          email: "openclaw@test.com",
          scopes: ["Mail.Send", "Mail.Read", "User.Read", "offline_access"],
          tools: ["email_send", "email_reply"],
        },
      },
    });

    plugin.register(api as any);

    const handler = eventHandlers.get("before_prompt_build");
    expect(handler).toBeDefined();

    const result = await handler!();
    // Should mention account names and auth status
    expect(result.prependContext).toBeDefined();
    expect(result.prependContext).toContain("rod");
    expect(result.prependContext).toContain("openclaw");
  });

  it("multi-account prompt context indicates authenticated accounts", async () => {
    const { default: plugin } = await import("./index.js");
    const { existsSync } = await import("node:fs");
    // rod is authenticated, openclaw is not
    (existsSync as ReturnType<typeof vi.fn>).mockImplementation((path: string) => {
      return (path as string).includes("credentials-rod");
    });

    const { api, eventHandlers } = makeMockApi({
      clientId: "test-client",
      tenantId: "test-tenant",
      clientSecret: "test-secret",
      defaultAccount: "rod",
      accounts: {
        rod: {
          name: "Rod (Read Only)",
          email: "rod@test.com",
          scopes: ["Mail.Read", "User.Read", "offline_access"],
          tools: ["email_list", "email_read"],
        },
        openclaw: {
          name: "OpenClaw (Sender)",
          email: "openclaw@test.com",
          scopes: ["Mail.Send", "Mail.Read", "User.Read", "offline_access"],
          tools: ["email_send", "email_reply"],
        },
      },
    });

    plugin.register(api as any);

    const handler = eventHandlers.get("before_prompt_build");
    const result = await handler!();
    // Should differentiate authenticated vs not
    expect(result.prependContext).toContain("authenticated");
    expect(result.prependContext).toMatch(/not authenticated|needs auth/i);
  });
});
