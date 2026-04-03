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

import { createGraphClient } from "./src/graph-client.js";
import * as fs from "node:fs";

const fsMock = vi.mocked(fs);

const TEST_CONFIG: Office365Config = {
  clientId: "test-client-id",
  tenantId: "550e8400-e29b-41d4-a716-446655440000",
  clientSecret: "test-secret",
  redirectUri: "http://localhost:8080/callback",
  scopes: ["Mail.ReadWrite", "User.Read", "offline_access"],
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

// ── Issue 5: writeFileSync TOCTOU race ─────────────────────────────────────

describe("writeCredentialFile sets mode on creation (Issue 5)", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("passes mode 0o600 to writeFileSync so file is never world-readable", () => {
    const client = createGraphClient({
      config: TEST_CONFIG,
      stateDir: "/tmp/test-state",
      logger: makeLogger(),
    });

    client.setCredential(TEST_CRED);

    // writeFileSync should be called with mode option
    expect(fsMock.writeFileSync).toHaveBeenCalledTimes(1);
    const writeCall = fsMock.writeFileSync.mock.calls[0];
    // Third arg should include mode: 0o600
    expect(writeCall[2]).toEqual(expect.objectContaining({ mode: 0o600 }));
  });
});

// ── Issue 6: chmod catch silently ignores on Linux ─────────────────────────

describe("chmodSync platform-aware error handling (Issue 6)", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("re-throws chmod errors on non-Windows platforms", () => {
    const originalPlatform = process.platform;
    Object.defineProperty(process, "platform", { value: "linux", configurable: true });

    try {
      fsMock.chmodSync.mockImplementation(() => {
        throw new Error("chmod failed: operation not permitted");
      });

      const client = createGraphClient({
        config: TEST_CONFIG,
        stateDir: "/tmp/test-state",
        logger: makeLogger(),
      });

      expect(() => client.setCredential(TEST_CRED)).toThrow("chmod failed");
    } finally {
      Object.defineProperty(process, "platform", { value: originalPlatform, configurable: true });
    }
  });

  it("catches chmod errors on Windows without throwing", () => {
    const originalPlatform = process.platform;
    Object.defineProperty(process, "platform", { value: "win32", configurable: true });

    try {
      fsMock.chmodSync.mockImplementation(() => {
        throw new Error("chmod not supported on Windows");
      });

      const logger = makeLogger();
      const client = createGraphClient({
        config: TEST_CONFIG,
        stateDir: "/tmp/test-state",
        logger,
      });

      // Should not throw on Windows
      expect(() => client.setCredential(TEST_CRED)).not.toThrow();
      // Should log a warning
      expect(logger.warn).toHaveBeenCalled();
    } finally {
      Object.defineProperty(process, "platform", { value: originalPlatform, configurable: true });
    }
  });
});

// ── Issue 13: mkdirSync without mode ───────────────────────────────────────

describe("mkdirSync uses restrictive mode (Issue 13)", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("passes mode 0o700 to mkdirSync", () => {
    const client = createGraphClient({
      config: TEST_CONFIG,
      stateDir: "/tmp/test-state",
      logger: makeLogger(),
    });

    client.setCredential(TEST_CRED);

    expect(fsMock.mkdirSync).toHaveBeenCalledTimes(1);
    const mkdirCall = fsMock.mkdirSync.mock.calls[0];
    expect(mkdirCall[1]).toEqual(expect.objectContaining({ recursive: true, mode: 0o700 }));
  });
});

// ── Issue 14: readCredentialFile doesn't validate types ─────────────────────

describe("readCredentialFile validates types (Issue 14)", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("returns null when access is not a string", async () => {
    fsMock.readFileSync.mockReturnValue(
      JSON.stringify({ access: 12345, refresh: "refresh", expires: Date.now() + 3600_000 }),
    );

    const { fetchWithSsrFGuard } = await import("openclaw/plugin-sdk");
    (fetchWithSsrFGuard as ReturnType<typeof vi.fn>).mockRejectedValue(new Error("should not reach fetch"));

    const client = createGraphClient({
      config: TEST_CONFIG,
      stateDir: "/tmp/test-state",
      logger: makeLogger(),
    });

    // Attempting to fetch should fail because credential file is invalid
    await expect(client.fetchJson("/me")).rejects.toThrow("not authenticated");
  });

  it("returns null when refresh is not a string", async () => {
    fsMock.readFileSync.mockReturnValue(
      JSON.stringify({ access: "access", refresh: null, expires: Date.now() + 3600_000 }),
    );

    const { fetchWithSsrFGuard } = await import("openclaw/plugin-sdk");
    (fetchWithSsrFGuard as ReturnType<typeof vi.fn>).mockRejectedValue(new Error("should not reach fetch"));

    const client = createGraphClient({
      config: TEST_CONFIG,
      stateDir: "/tmp/test-state",
      logger: makeLogger(),
    });

    await expect(client.fetchJson("/me")).rejects.toThrow("not authenticated");
  });

  it("returns null when expires is a string instead of number", async () => {
    fsMock.readFileSync.mockReturnValue(
      JSON.stringify({ access: "access", refresh: "refresh", expires: "not-a-number" }),
    );

    const { fetchWithSsrFGuard } = await import("openclaw/plugin-sdk");
    (fetchWithSsrFGuard as ReturnType<typeof vi.fn>).mockRejectedValue(new Error("should not reach fetch"));

    const client = createGraphClient({
      config: TEST_CONFIG,
      stateDir: "/tmp/test-state",
      logger: makeLogger(),
    });

    await expect(client.fetchJson("/me")).rejects.toThrow("not authenticated");
  });

  it("returns null when expires is negative", async () => {
    fsMock.readFileSync.mockReturnValue(
      JSON.stringify({ access: "access", refresh: "refresh", expires: -1 }),
    );

    const { fetchWithSsrFGuard } = await import("openclaw/plugin-sdk");
    (fetchWithSsrFGuard as ReturnType<typeof vi.fn>).mockRejectedValue(new Error("should not reach fetch"));

    const client = createGraphClient({
      config: TEST_CONFIG,
      stateDir: "/tmp/test-state",
      logger: makeLogger(),
    });

    await expect(client.fetchJson("/me")).rejects.toThrow("not authenticated");
  });

  it("accepts valid credential with correct types", () => {
    fsMock.readFileSync.mockReturnValue(
      JSON.stringify({
        access: "valid-access",
        refresh: "valid-refresh",
        expires: Date.now() + 3600_000,
      }),
    );

    // If readCredentialFile returns a valid object, createGraphClient should work
    // We can verify by creating the client without errors
    const client = createGraphClient({
      config: TEST_CONFIG,
      stateDir: "/tmp/test-state",
      logger: makeLogger(),
    });

    expect(client).toBeDefined();
  });
});

// ── Phase 2: credentialPath with accountId ────────────────────────────────

describe("credentialPath with accountId", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("uses original filename for default accountId", () => {
    const client = createGraphClient({
      config: TEST_CONFIG,
      stateDir: "/tmp/test-state",
      logger: makeLogger(),
    });

    client.setCredential(TEST_CRED);

    // writeFileSync should write to the original path (no account suffix)
    const writeCall = fsMock.writeFileSync.mock.calls[0];
    const writtenPath = writeCall[0] as string;
    expect(writtenPath).toMatch(/office365-credentials\.json\./);
    expect(writtenPath).not.toContain("office365-credentials-");
  });

  it("uses original filename when accountId is omitted", () => {
    const client = createGraphClient({
      config: TEST_CONFIG,
      stateDir: "/tmp/test-state",
      logger: makeLogger(),
    });

    client.setCredential(TEST_CRED);

    const writeCall = fsMock.writeFileSync.mock.calls[0];
    const writtenPath = writeCall[0] as string;
    expect(writtenPath).toMatch(/office365-credentials\.json\./);
    expect(writtenPath).not.toContain("office365-credentials-");
  });

  it("uses account-specific filename for non-default accountId", () => {
    const client = createGraphClient({
      config: TEST_CONFIG,
      stateDir: "/tmp/test-state",
      logger: makeLogger(),
      accountId: "rod",
    });

    client.setCredential(TEST_CRED);

    const writeCall = fsMock.writeFileSync.mock.calls[0];
    const writtenPath = writeCall[0] as string;
    expect(writtenPath).toMatch(/office365-credentials-rod\.json\./);
  });

  it("uses account-specific filename for openclaw accountId", () => {
    const client = createGraphClient({
      config: TEST_CONFIG,
      stateDir: "/tmp/test-state",
      logger: makeLogger(),
      accountId: "openclaw",
    });

    client.setCredential(TEST_CRED);

    const writeCall = fsMock.writeFileSync.mock.calls[0];
    const writtenPath = writeCall[0] as string;
    expect(writtenPath).toMatch(/office365-credentials-openclaw\.json\./);
  });

  it("reads from account-specific credential file", async () => {
    fsMock.readFileSync.mockReturnValue(JSON.stringify(TEST_CRED));

    const { fetchWithSsrFGuard } = await import("openclaw/plugin-sdk");
    (fetchWithSsrFGuard as ReturnType<typeof vi.fn>).mockResolvedValue({
      response: new Response(JSON.stringify({ id: "123" }), { status: 200 }),
      release: vi.fn(),
    });

    const client = createGraphClient({
      config: TEST_CONFIG,
      stateDir: "/tmp/test-state",
      logger: makeLogger(),
      accountId: "rod",
    });

    await client.fetchJson("/me");

    // readFileSync should be called with the account-specific path
    const readCall = fsMock.readFileSync.mock.calls[0];
    expect(readCall[0]).toContain("office365-credentials-rod.json");
  });

  it("rejects accountId that fails ACCOUNT_ID_RE validation", () => {
    expect(() =>
      createGraphClient({
        config: TEST_CONFIG,
        stateDir: "/tmp/test-state",
        logger: makeLogger(),
        accountId: "../evil",
      }),
    ).toThrow(/invalid account ID/i);
  });
});
