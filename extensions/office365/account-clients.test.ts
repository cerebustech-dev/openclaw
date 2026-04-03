import { describe, expect, it, vi, beforeEach } from "vitest";
import type { PluginLogger } from "openclaw/plugin-sdk";
import type { Office365Config, Office365Credential } from "./src/types.js";

// ── Mocks ──────────────────────────────────────────────────────────────────

vi.mock("node:fs", () => ({
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
}));

vi.mock("./src/oauth.js", () => ({
  refreshMicrosoftTokens: vi.fn(),
}));

import { createAccountClients } from "./src/account-clients.js";
import * as fs from "node:fs";

const fsMock = vi.mocked(fs);

// ── Test fixtures ──────────────────────────────────────────────────────────

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

const TEST_CRED: Office365Credential = {
  access: "access-token-123",
  refresh: "refresh-token-456",
  expires: Date.now() + 3600_000,
  email: "user@example.com",
};

function makeLogger(): PluginLogger {
  return {
    info: vi.fn(),
    warn: vi.fn(),
    debug: vi.fn(),
    error: vi.fn(),
  } as unknown as PluginLogger;
}

// ── createAccountClients ──────────────────────────────────────────────────

describe("createAccountClients", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("returns a single 'default' client for legacy config", () => {
    const clients = createAccountClients({
      config: LEGACY_CONFIG,
      stateDir: "/tmp/test-state",
      logger: makeLogger(),
    });

    expect(clients).toBeInstanceOf(Map);
    expect(clients.size).toBe(1);
    expect(clients.has("default")).toBe(true);
  });

  it("returns one client per account for multi-account config", () => {
    const clients = createAccountClients({
      config: MULTI_CONFIG,
      stateDir: "/tmp/test-state",
      logger: makeLogger(),
    });

    expect(clients.size).toBe(2);
    expect(clients.has("rod")).toBe(true);
    expect(clients.has("openclaw")).toBe(true);
  });

  it("each client writes to its own credential file", () => {
    const clients = createAccountClients({
      config: MULTI_CONFIG,
      stateDir: "/tmp/test-state",
      logger: makeLogger(),
    });

    // Write credential via rod client
    clients.get("rod")!.setCredential(TEST_CRED);
    const rodWritePath = fsMock.writeFileSync.mock.calls[0][0] as string;
    expect(rodWritePath).toMatch(/office365-credentials-rod\.json/);

    vi.clearAllMocks();

    // Write credential via openclaw client
    clients.get("openclaw")!.setCredential(TEST_CRED);
    const openclawWritePath = fsMock.writeFileSync.mock.calls[0][0] as string;
    expect(openclawWritePath).toMatch(/office365-credentials-openclaw\.json/);
  });

  it("legacy client writes to original credential filename", () => {
    const clients = createAccountClients({
      config: LEGACY_CONFIG,
      stateDir: "/tmp/test-state",
      logger: makeLogger(),
    });

    clients.get("default")!.setCredential(TEST_CRED);
    const writePath = fsMock.writeFileSync.mock.calls[0][0] as string;
    expect(writePath).toMatch(/office365-credentials\.json/);
    expect(writePath).not.toContain("office365-credentials-");
  });
});
