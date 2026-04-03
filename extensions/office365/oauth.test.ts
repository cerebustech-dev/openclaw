import { describe, expect, it, vi, beforeEach } from "vitest";

vi.mock("openclaw/plugin-sdk", () => ({
  fetchWithSsrFGuard: async (params: {
    url: string;
    init?: RequestInit;
  }) => {
    const response = await globalThis.fetch(params.url, params.init);
    return { response, finalUrl: params.url, release: async () => {} };
  },
}));

import {
  generatePkce,
  buildAuthorizeUrl,
  parseCallbackInput,
  exchangeCodeForTokens,
  refreshMicrosoftTokens,
  getUserInfo,
  validateTenantId,
} from "./src/oauth.js";
import type { Office365Config } from "./src/types.js";

const TEST_CONFIG: Office365Config = {
  clientId: "test-client-id",
  tenantId: "550e8400-e29b-41d4-a716-446655440000",
  clientSecret: "test-secret",
  redirectUri: "http://localhost:8080/callback",
  scopes: ["Mail.ReadWrite", "User.Read", "offline_access"],
};

describe("generatePkce", () => {
  it("produces a valid verifier and challenge pair", () => {
    const { verifier, challenge } = generatePkce();
    expect(verifier).toHaveLength(64); // 32 bytes as hex
    expect(challenge).toBeTruthy();
    expect(challenge).not.toBe(verifier);
    // Challenge should be base64url encoded (no +, /, or =)
    expect(challenge).toMatch(/^[A-Za-z0-9_-]+$/);
  });

  it("produces unique pairs on each call", () => {
    const a = generatePkce();
    const b = generatePkce();
    expect(a.verifier).not.toBe(b.verifier);
    expect(a.challenge).not.toBe(b.challenge);
  });
});

describe("buildAuthorizeUrl", () => {
  it("includes all required params with tenant-specific URL", () => {
    const url = buildAuthorizeUrl(TEST_CONFIG, "test-challenge", "test-state");
    const parsed = new URL(url);
    expect(parsed.origin).toBe("https://login.microsoftonline.com");
    expect(parsed.pathname).toBe("/550e8400-e29b-41d4-a716-446655440000/oauth2/v2.0/authorize");
    expect(parsed.searchParams.get("client_id")).toBe("test-client-id");
    expect(parsed.searchParams.get("response_type")).toBe("code");
    expect(parsed.searchParams.get("redirect_uri")).toBe(
      "http://localhost:8080/callback",
    );
    expect(parsed.searchParams.get("scope")).toBe(
      "Mail.ReadWrite User.Read offline_access",
    );
    expect(parsed.searchParams.get("code_challenge")).toBe("test-challenge");
    expect(parsed.searchParams.get("code_challenge_method")).toBe("S256");
    expect(parsed.searchParams.get("state")).toBe("test-state");
    expect(parsed.searchParams.get("prompt")).toBe("consent");
  });

  it("includes login_hint when provided", () => {
    const url = buildAuthorizeUrl(TEST_CONFIG, "challenge", "state", "rod@example.com");
    const parsed = new URL(url);
    expect(parsed.searchParams.get("login_hint")).toBe("rod@example.com");
  });

  it("omits login_hint when not provided", () => {
    const url = buildAuthorizeUrl(TEST_CONFIG, "challenge", "state");
    const parsed = new URL(url);
    expect(parsed.searchParams.get("login_hint")).toBeNull();
  });
});

describe("parseCallbackInput", () => {
  const STATE = "expected-state-123";

  it("extracts code and state from a full URL", () => {
    const result = parseCallbackInput(
      `http://localhost:8080/callback?code=auth-code-123&state=${STATE}`,
      STATE,
    );
    expect(result).toEqual({ code: "auth-code-123", state: STATE });
  });

  it("handles query-string-only paste", () => {
    const result = parseCallbackInput(
      `?code=auth-code-456&state=${STATE}`,
      STATE,
    );
    expect(result).toEqual({ code: "auth-code-456", state: STATE });
  });

  it("handles query-string without leading ?", () => {
    const result = parseCallbackInput(
      `code=auth-code-789&state=${STATE}`,
      STATE,
    );
    expect(result).toEqual({ code: "auth-code-789", state: STATE });
  });

  it("returns error for truncated URL missing state", () => {
    const result = parseCallbackInput(
      "http://localhost:8080/callback?code=auth-code-123",
      STATE,
    );
    expect("error" in result).toBe(true);
    if ("error" in result) {
      expect(result.error).toContain("state");
      expect(result.error).toContain("truncated");
    }
  });

  it("returns error for state mismatch with diagnostic", () => {
    const result = parseCallbackInput(
      `http://localhost:8080/callback?code=abc&state=wrong-state`,
      STATE,
    );
    expect("error" in result).toBe(true);
    if ("error" in result) {
      expect(result.error).toContain("mismatch");
      expect(result.error).toContain("tenant");
    }
  });

  it("returns error for empty input", () => {
    const result = parseCallbackInput("", STATE);
    expect("error" in result).toBe(true);
  });

  it("returns error for garbage input", () => {
    const result = parseCallbackInput("not a url at all", STATE);
    expect("error" in result).toBe(true);
  });
});

describe("exchangeCodeForTokens", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it("sends correct POST body and parses response", async () => {
    const mockFetch = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          access_token: "access-123",
          refresh_token: "refresh-456",
          expires_in: 3600,
        }),
        { status: 200 },
      ),
    );
    globalThis.fetch = mockFetch;

    // Also mock the user info call
    mockFetch.mockResolvedValueOnce(
      new Response(
        JSON.stringify({
          access_token: "access-123",
          refresh_token: "refresh-456",
          expires_in: 3600,
        }),
        { status: 200 },
      ),
    ).mockResolvedValueOnce(
      new Response(
        JSON.stringify({ mail: "user@example.com" }),
        { status: 200 },
      ),
    );

    const result = await exchangeCodeForTokens("auth-code", "verifier", TEST_CONFIG);

    expect(result.access).toBe("access-123");
    expect(result.refresh).toBe("refresh-456");
    expect(result.email).toBe("user@example.com");
    expect(result.expires).toBeGreaterThan(Date.now());

    // Verify the token exchange call
    const firstCall = mockFetch.mock.calls[0];
    expect(firstCall[0]).toContain("550e8400-e29b-41d4-a716-446655440000/oauth2/v2.0/token");
  });

  it("throws on non-200 response", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue(
      new Response("Bad request", { status: 400 }),
    );

    await expect(
      exchangeCodeForTokens("bad-code", "verifier", TEST_CONFIG),
    ).rejects.toThrow("Token exchange failed");
  });

  it("throws when no refresh token received", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({ access_token: "access-123", expires_in: 3600 }),
        { status: 200 },
      ),
    );

    await expect(
      exchangeCodeForTokens("code", "verifier", TEST_CONFIG),
    ).rejects.toThrow("No refresh token");
  });
});

describe("refreshMicrosoftTokens", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it("sends correct refresh POST body", async () => {
    const mockFetch = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          access_token: "new-access",
          refresh_token: "new-refresh",
          expires_in: 3600,
        }),
        { status: 200 },
      ),
    );
    globalThis.fetch = mockFetch;

    const result = await refreshMicrosoftTokens("old-refresh", TEST_CONFIG);

    expect(result.access).toBe("new-access");
    expect(result.refresh).toBe("new-refresh");
    expect(result.expires).toBeGreaterThan(Date.now());

    const call = mockFetch.mock.calls[0];
    expect(call[0]).toContain("550e8400-e29b-41d4-a716-446655440000/oauth2/v2.0/token");
  });

  it("keeps old refresh token when not provided in response (RFC 6749 §6)", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          access_token: "new-access",
          expires_in: 3600,
        }),
        { status: 200 },
      ),
    );

    const result = await refreshMicrosoftTokens("old-refresh", TEST_CONFIG);
    expect(result.refresh).toBe("old-refresh");
  });

  it("throws clear re-auth message for revoked refresh token", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({ error: "invalid_grant", error_description: "AADSTS700082" }),
        { status: 400 },
      ),
    );

    await expect(
      refreshMicrosoftTokens("revoked-token", TEST_CONFIG),
    ).rejects.toThrow("re-authenticate");
  });
});

describe("getUserInfo", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it("extracts email from Graph /me response", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          mail: "user@contoso.com",
          displayName: "Test User",
          userPrincipalName: "user@contoso.onmicrosoft.com",
        }),
        { status: 200 },
      ),
    );

    const email = await getUserInfo("token");
    expect(email).toBe("user@contoso.com");
  });

  it("falls back to userPrincipalName when mail is null", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          mail: null,
          userPrincipalName: "user@contoso.onmicrosoft.com",
        }),
        { status: 200 },
      ),
    );

    const email = await getUserInfo("token");
    expect(email).toBe("user@contoso.onmicrosoft.com");
  });

  it("handles non-ASCII display names without error", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          mail: "tëst@example.com",
          displayName: "Ünïcödë Üser 日本語",
        }),
        { status: 200 },
      ),
    );

    const email = await getUserInfo("token");
    expect(email).toBe("tëst@example.com");
  });

  it("returns undefined on API error", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue(
      new Response("Unauthorized", { status: 401 }),
    );

    const email = await getUserInfo("bad-token");
    expect(email).toBeUndefined();
  });
});

// ── Issue 2: tenantId validation ───────────────────────────────────────────

describe("validateTenantId", () => {
  it("accepts a valid GUID", () => {
    expect(() => validateTenantId("550e8400-e29b-41d4-a716-446655440000")).not.toThrow();
  });

  it("accepts well-known value 'common'", () => {
    expect(() => validateTenantId("common")).not.toThrow();
  });

  it("accepts well-known value 'organizations'", () => {
    expect(() => validateTenantId("organizations")).not.toThrow();
  });

  it("accepts well-known value 'consumers'", () => {
    expect(() => validateTenantId("consumers")).not.toThrow();
  });

  it("rejects tenantId containing path traversal slash", () => {
    expect(() => validateTenantId("../../evil")).toThrow();
  });

  it("rejects tenantId containing query separator ?", () => {
    expect(() => validateTenantId("tenant?client_id=evil")).toThrow();
  });

  it("rejects tenantId containing fragment #", () => {
    expect(() => validateTenantId("tenant#fragment")).toThrow();
  });

  it("rejects empty tenantId", () => {
    expect(() => validateTenantId("")).toThrow();
  });

  it("rejects tenantId that looks like a GUID but is wrong length", () => {
    expect(() => validateTenantId("550e8400-e29b-41d4-a716")).toThrow();
  });
});

describe("buildAuthorizeUrl tenantId validation", () => {
  it("throws for malicious tenantId in buildAuthorizeUrl", () => {
    const badConfig = { ...TEST_CONFIG, tenantId: "../../attack?evil=1" };
    expect(() => buildAuthorizeUrl(badConfig, "challenge", "state")).toThrow();
  });
});

describe("exchangeCodeForTokens tenantId validation", () => {
  it("throws for malicious tenantId in token exchange", async () => {
    const badConfig = { ...TEST_CONFIG, tenantId: "tenant/../../evil" };
    await expect(
      exchangeCodeForTokens("code", "verifier", badConfig),
    ).rejects.toThrow();
  });
});

describe("refreshMicrosoftTokens tenantId validation", () => {
  it("throws for malicious tenantId in token refresh", async () => {
    const badConfig = { ...TEST_CONFIG, tenantId: "tenant?inject=true" };
    await expect(
      refreshMicrosoftTokens("refresh-token", badConfig),
    ).rejects.toThrow();
  });
});

// ── Issue 3: Loopback enforcement ──────────────────────────────────────────

describe("loginMicrosoftOAuth loopback enforcement", () => {
  it("rejects non-loopback redirectUri", async () => {
    const remoteConfig = { ...TEST_CONFIG, redirectUri: "http://evil.com:8080/callback" };
    const ctx = {
      isRemote: false,
      openUrl: vi.fn().mockResolvedValue(undefined),
      log: vi.fn(),
      note: vi.fn().mockResolvedValue(undefined),
      prompt: vi.fn().mockResolvedValue(""),
      progress: { update: vi.fn(), stop: vi.fn() },
    };

    const { loginMicrosoftOAuth } = await import("./src/oauth.js");
    await expect(loginMicrosoftOAuth(ctx, remoteConfig)).rejects.toThrow("loopback");
  });
});

// ── Issue 4: CSRF state check on error handler ────────────────────────────

describe("callback server CSRF protection on error responses", () => {
  it("ignores error callback without valid state parameter", async () => {
    // This test verifies that the callback server does not process error
    // responses that lack a valid state parameter (CSRF protection).
    // We test this through loginMicrosoftOAuth's local flow.
    const { loginMicrosoftOAuth } = await import("./src/oauth.js");

    const localConfig = { ...TEST_CONFIG, redirectUri: "http://localhost:9876/callback" };
    const ctx = {
      isRemote: false,
      openUrl: vi.fn().mockResolvedValue(undefined),
      log: vi.fn(),
      note: vi.fn().mockResolvedValue(undefined),
      prompt: vi.fn().mockResolvedValue(""),
      progress: { update: vi.fn(), stop: vi.fn() },
    };

    // Start the OAuth flow (will start a server on port 9876)
    const oauthPromise = loginMicrosoftOAuth(ctx, localConfig);

    // Wait for server to start listening
    await new Promise((r) => setTimeout(r, 200));

    // Send a CSRF attack: error without state
    try {
      await fetch("http://localhost:9876/callback?error=access_denied");
    } catch {
      // Connection errors are fine if server rejects
    }

    // The OAuth flow should NOT have terminated from the stateless error.
    // Send another CSRF attack: error with wrong state
    try {
      await fetch("http://localhost:9876/callback?error=access_denied&state=wrong-state");
    } catch {
      // Connection errors are fine
    }

    // Give it a moment to process
    await new Promise((r) => setTimeout(r, 100));

    // The flow should still be running (not rejected by CSRF).
    // Now kill it with a timeout-style abort by racing with a short timer.
    const result = await Promise.race([
      oauthPromise.catch((e: Error) => ({ error: e.message })),
      new Promise<{ stillRunning: true }>((r) => setTimeout(() => r({ stillRunning: true }), 500)),
    ]);

    // If the flow is still running (didn't die from CSRF), that's the correct behavior.
    // Clean up by aborting — the timeout in loginMicrosoftOAuth will eventually clean up,
    // but we can't easily cancel it here, so we just verify the CSRF didn't kill it.
    if ("stillRunning" in result) {
      // Good: the CSRF error without state did NOT terminate the flow
      expect(result.stillRunning).toBe(true);
    } else {
      // If it errored, it should NOT be because of the CSRF attack
      expect(result.error).not.toContain("access_denied");
    }
  });
});
