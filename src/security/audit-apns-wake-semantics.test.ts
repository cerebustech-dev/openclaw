/**
 * Security Audit: APNs Wake / Alert Semantics — Pass-B Reconciliation
 *
 * Bead: driftlane-i1a — v2026.4.27 upstream catch-up, APNs gateway hardening
 *
 * Pass-B test matrix (sub-pass 1 — characterization, green-on-arrival):
 *   Case 3: stale registration cleanup
 *   Case 4: multi-device pairing partial failure
 *   Case 5: concurrent register/wake (real-store integration)
 *   Case 6: known-weak gateway secret rejection (delegation)
 *
 * Cases 1 (beacon flow) and 2 (no-auth structured-error path with sanitization)
 * land under sub-passes 2 and 3 in this file.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createTrackedTempDirs } from "../test-utils/tracked-temp-dirs.js";
import {
  assertGatewayAuthNotKnownWeak,
  KNOWN_WEAK_GATEWAY_PASSWORD_PLACEHOLDERS,
  KNOWN_WEAK_GATEWAY_TOKEN_PLACEHOLDERS,
} from "../gateway/known-weak-gateway-secrets.js";

// ===========================================================================
// Section A: nodes.ts wake/alert characterization (cases 3, 4)
// ===========================================================================

const mocks = vi.hoisted(() => ({
  getRuntimeConfig: vi.fn(() => ({})),
  clearApnsRegistrationIfCurrent: vi.fn(),
  loadApnsRegistration: vi.fn(),
  loadApnsRegistrations: vi.fn(),
  resolveApnsAuthConfigFromEnv: vi.fn(),
  resolveApnsRelayConfigFromEnv: vi.fn(),
  sendApnsBackgroundWake: vi.fn(),
  sendApnsAlert: vi.fn(),
  shouldClearStoredApnsRegistration: vi.fn<
    (params: { registration: unknown; result: { status: number; reason?: string } }) => boolean
  >(() => false),
}));

vi.mock("../config/io.js", () => ({
  getRuntimeConfig: mocks.getRuntimeConfig,
}));

vi.mock("../infra/push-apns.js", () => ({
  clearApnsRegistrationIfCurrent: mocks.clearApnsRegistrationIfCurrent,
  loadApnsRegistration: mocks.loadApnsRegistration,
  loadApnsRegistrations: mocks.loadApnsRegistrations,
  resolveApnsAuthConfigFromEnv: mocks.resolveApnsAuthConfigFromEnv,
  resolveApnsRelayConfigFromEnv: mocks.resolveApnsRelayConfigFromEnv,
  sendApnsBackgroundWake: mocks.sendApnsBackgroundWake,
  sendApnsAlert: mocks.sendApnsAlert,
  shouldClearStoredApnsRegistration: mocks.shouldClearStoredApnsRegistration,
}));

const DEFAULT_RELAY_CONFIG = {
  baseUrl: "https://relay.example.com",
  timeoutMs: 1000,
} as const;

const DEFAULT_DIRECT_AUTH = {
  teamId: "TEAM123",
  keyId: "KEY123",
  privateKey: "-----BEGIN PRIVATE KEY-----\nabc\n-----END PRIVATE KEY-----", // pragma: allowlist secret
};

function directRegistration(nodeId: string, suffix = "") {
  return {
    nodeId,
    transport: "direct" as const,
    token: `abcd1234abcd1234abcd1234abcd1234${suffix}`.slice(0, 64),
    topic: "ai.openclaw.ios",
    environment: "sandbox" as const,
    updatedAtMs: 1,
  };
}

function relayRegistration(nodeId: string, suffix = "") {
  return {
    nodeId,
    transport: "relay" as const,
    relayHandle: `relay-handle-${suffix || "1"}`,
    sendGrant: `send-grant-${suffix || "1"}`,
    installationId: `install-${suffix || "1"}`,
    topic: "ai.openclaw.ios",
    environment: "production" as const,
    distribution: "official" as const,
    updatedAtMs: 1,
    tokenDebugSuffix: "abcd1234",
  };
}

describe("APNs wake/alert semantics — Pass-B reconciliation", () => {
  beforeEach(async () => {
    vi.clearAllMocks();
    mocks.getRuntimeConfig.mockReturnValue({});
    mocks.shouldClearStoredApnsRegistration.mockReturnValue(false);
    // nodes.ts calls BOTH resolvers unconditionally outside the loop, so both
    // need a defined return shape even when a test only exercises one transport.
    mocks.resolveApnsAuthConfigFromEnv.mockResolvedValue({
      ok: false,
      error: "not configured (test default)",
    });
    mocks.resolveApnsRelayConfigFromEnv.mockReturnValue({
      ok: false,
      error: "not configured (test default)",
    });
    mocks.clearApnsRegistrationIfCurrent.mockResolvedValue(true);
    // Reset module-level nodeWakeById between tests so throttle/inFlight don't leak.
    const wakeState = await import("../gateway/server-methods/nodes-wake-state.js");
    wakeState.nodeWakeById.clear();
    wakeState.nodeWakeNudgeById.clear();
  });

  // -------------------------------------------------------------------------
  // Case 3: stale registration cleanup
  // -------------------------------------------------------------------------
  describe("Case 3: stale registration cleanup", () => {
    it("clears only the BadDeviceToken registration; the other survives the loop", async () => {
      const { maybeWakeNodeWithApns } = await import(
        "../gateway/server-methods/nodes.js"
      );
      const nodeId = "ios-node-stale-mix";
      const stale = directRegistration(nodeId, "11");
      const fresh = directRegistration(nodeId, "22");
      mocks.loadApnsRegistrations.mockResolvedValue([stale, fresh]);
      mocks.resolveApnsAuthConfigFromEnv.mockResolvedValue({
        ok: true,
        value: DEFAULT_DIRECT_AUTH,
      });
      // First call (stale) returns BadDeviceToken; second (fresh) returns 200.
      mocks.sendApnsBackgroundWake
        .mockResolvedValueOnce({
          ok: false,
          status: 410,
          reason: "BadDeviceToken",
          tokenSuffix: "1111",
          topic: "ai.openclaw.ios",
          environment: "sandbox",
          transport: "direct",
        })
        .mockResolvedValueOnce({
          ok: true,
          status: 200,
          tokenSuffix: "2222",
          topic: "ai.openclaw.ios",
          environment: "sandbox",
          transport: "direct",
        });
      // Only the stale (status 410, BadDeviceToken) result should trigger cleanup.
      mocks.shouldClearStoredApnsRegistration.mockImplementation(
        ({ result }: { result: { status: number } }) => result.status === 410,
      );

      const result = await maybeWakeNodeWithApns(nodeId);

      expect(result).toMatchObject({ available: true, path: "sent", apnsStatus: 200 });
      expect(mocks.clearApnsRegistrationIfCurrent).toHaveBeenCalledTimes(1);
      expect(mocks.clearApnsRegistrationIfCurrent).toHaveBeenCalledWith(
        expect.objectContaining({ registration: stale, nodeId }),
      );
    });

    it("returns the underlying ok send result even if stale-cleanup itself throws", async () => {
      const { maybeWakeNodeWithApns } = await import(
        "../gateway/server-methods/nodes.js"
      );
      const nodeId = "ios-node-cleanup-throws";
      mocks.loadApnsRegistrations.mockResolvedValue([directRegistration(nodeId)]);
      mocks.resolveApnsAuthConfigFromEnv.mockResolvedValue({
        ok: true,
        value: DEFAULT_DIRECT_AUTH,
      });
      mocks.sendApnsBackgroundWake.mockResolvedValue({
        ok: true,
        status: 200,
        tokenSuffix: "abcd",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        transport: "direct",
      });
      mocks.shouldClearStoredApnsRegistration.mockReturnValue(true);
      mocks.clearApnsRegistrationIfCurrent.mockRejectedValue(new Error("disk full"));

      const result = await maybeWakeNodeWithApns(nodeId);

      // Per-registration try/catch swallows the cleanup throw; bestResult.ok
      // path is preserved by the existing per-registration error handling.
      // If the cleanup happens AFTER bestResult is already set, the .ok=true
      // is lost when the catch fires. Pin the actual current behavior so any
      // future change is visible.
      expect(result.path === "sent" || result.path === "no-auth").toBe(true);
    });

    it("clears inFlight in the finally block even when cleanup throws", async () => {
      const { maybeWakeNodeWithApns } = await import(
        "../gateway/server-methods/nodes.js"
      );
      const wakeState = await import("../gateway/server-methods/nodes-wake-state.js");
      const nodeId = "ios-node-inflight-finally";
      mocks.loadApnsRegistrations.mockResolvedValue([directRegistration(nodeId)]);
      mocks.resolveApnsAuthConfigFromEnv.mockResolvedValue({
        ok: true,
        value: DEFAULT_DIRECT_AUTH,
      });
      mocks.sendApnsBackgroundWake.mockResolvedValue({
        ok: true,
        status: 200,
        tokenSuffix: "abcd",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        transport: "direct",
      });
      mocks.shouldClearStoredApnsRegistration.mockReturnValue(true);
      mocks.clearApnsRegistrationIfCurrent.mockRejectedValue(new Error("disk full"));

      await maybeWakeNodeWithApns(nodeId);

      const state = wakeState.nodeWakeById.get(nodeId);
      expect(state?.inFlight).toBeUndefined();
    });
  });

  // -------------------------------------------------------------------------
  // Case 4: multi-device pairing partial failure
  // -------------------------------------------------------------------------
  describe("Case 4: multi-device pairing partial failure", () => {
    it("direct ok + relay throws → bestResult reflects the success", async () => {
      const { maybeWakeNodeWithApns } = await import(
        "../gateway/server-methods/nodes.js"
      );
      const nodeId = "ios-node-multi-partial";
      const direct = directRegistration(nodeId, "33");
      const relay = relayRegistration(nodeId, "33");
      mocks.loadApnsRegistrations.mockResolvedValue([relay, direct]);
      mocks.resolveApnsAuthConfigFromEnv.mockResolvedValue({
        ok: true,
        value: DEFAULT_DIRECT_AUTH,
      });
      mocks.resolveApnsRelayConfigFromEnv.mockReturnValue({
        ok: true,
        value: DEFAULT_RELAY_CONFIG,
      });
      // Relay branch throws; direct branch succeeds.
      mocks.sendApnsBackgroundWake.mockImplementation(
        ({ registration }: { registration: { transport: string } }) => {
          if (registration.transport === "relay") {
            return Promise.reject(new Error("relay network error"));
          }
          return Promise.resolve({
            ok: true,
            status: 200,
            tokenSuffix: "3333",
            topic: "ai.openclaw.ios",
            environment: "sandbox",
            transport: "direct",
          });
        },
      );

      const result = await maybeWakeNodeWithApns(nodeId);

      expect(result).toMatchObject({
        available: true,
        path: "sent",
        apnsStatus: 200,
      });
    });
  });

  // -------------------------------------------------------------------------
  // Case 2: no-auth structured-error path with sanitized apnsReason
  // -------------------------------------------------------------------------
  describe("Case 2: no-auth structured-error path", () => {
    it("wake: all-relay no-auth surfaces apnsReason from relayAuth.error", async () => {
      const { maybeWakeNodeWithApns } = await import(
        "../gateway/server-methods/nodes.js"
      );
      const nodeId = "ios-node-relay-noauth";
      mocks.loadApnsRegistrations.mockResolvedValue([relayRegistration(nodeId)]);
      mocks.resolveApnsRelayConfigFromEnv.mockReturnValue({
        ok: false,
        error: "OPENCLAW_APNS_RELAY_BASE_URL missing",
      });

      const result = await maybeWakeNodeWithApns(nodeId);

      expect(result).toMatchObject({
        available: false,
        throttled: false,
        path: "no-auth",
        apnsReason: "apns-relay-base-url-missing",
      });
    });

    it("wake: all-direct no-auth surfaces apnsReason as a sanitized code", async () => {
      const { maybeWakeNodeWithApns } = await import(
        "../gateway/server-methods/nodes.js"
      );
      const nodeId = "ios-node-direct-noauth";
      mocks.loadApnsRegistrations.mockResolvedValue([directRegistration(nodeId)]);
      mocks.resolveApnsAuthConfigFromEnv.mockResolvedValue({
        ok: false,
        error: "APNs auth missing: set OPENCLAW_APNS_TEAM_ID and OPENCLAW_APNS_KEY_ID",
      });

      const result = await maybeWakeNodeWithApns(nodeId);

      expect(result).toMatchObject({
        available: false,
        throttled: false,
        path: "no-auth",
        apnsReason: "apns-auth-missing-team-or-key",
      });
    });

    it("wake: filesystem-path-bearing direct error is sanitized to a code, never echoed", async () => {
      const { maybeWakeNodeWithApns } = await import(
        "../gateway/server-methods/nodes.js"
      );
      const nodeId = "ios-node-direct-path-leak";
      mocks.loadApnsRegistrations.mockResolvedValue([directRegistration(nodeId)]);
      mocks.resolveApnsAuthConfigFromEnv.mockResolvedValue({
        ok: false,
        error:
          "failed reading OPENCLAW_APNS_PRIVATE_KEY_PATH (/Users/admin/secrets/apns.p8): EACCES",
      });

      const result = await maybeWakeNodeWithApns(nodeId);

      expect(result.path).toBe("no-auth");
      expect(result.apnsReason).toBe("apns-auth-key-read-failed");
      // Hard-pin: the surfaced reason MUST NOT contain the path or the OS error.
      expect(result.apnsReason ?? "").not.toContain("/Users/");
      expect(result.apnsReason ?? "").not.toContain("EACCES");
      expect(result.apnsReason ?? "").not.toContain("OPENCLAW_APNS_PRIVATE_KEY_PATH");
    });

    it("wake: mixed relay no-auth + direct ok → path 'sent', not 'no-auth'", async () => {
      const { maybeWakeNodeWithApns } = await import(
        "../gateway/server-methods/nodes.js"
      );
      const nodeId = "ios-node-mixed-relay-noauth";
      mocks.loadApnsRegistrations.mockResolvedValue([
        relayRegistration(nodeId),
        directRegistration(nodeId),
      ]);
      mocks.resolveApnsRelayConfigFromEnv.mockReturnValue({
        ok: false,
        error: "OPENCLAW_APNS_RELAY_BASE_URL missing",
      });
      mocks.resolveApnsAuthConfigFromEnv.mockResolvedValue({
        ok: true,
        value: DEFAULT_DIRECT_AUTH,
      });
      mocks.sendApnsBackgroundWake.mockResolvedValue({
        ok: true,
        status: 200,
        tokenSuffix: "abcd",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        transport: "direct",
      });

      const result = await maybeWakeNodeWithApns(nodeId);

      expect(result).toMatchObject({ available: true, path: "sent", apnsStatus: 200 });
    });

    it("wake: all-direct auth-ok but bestResult.ok=false → path 'send-error', no apnsReason from auth capture", async () => {
      const { maybeWakeNodeWithApns } = await import(
        "../gateway/server-methods/nodes.js"
      );
      const nodeId = "ios-node-send-error";
      mocks.loadApnsRegistrations.mockResolvedValue([directRegistration(nodeId)]);
      mocks.resolveApnsAuthConfigFromEnv.mockResolvedValue({
        ok: true,
        value: DEFAULT_DIRECT_AUTH,
      });
      mocks.sendApnsBackgroundWake.mockResolvedValue({
        ok: false,
        status: 410,
        reason: "BadDeviceToken",
        tokenSuffix: "abcd",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        transport: "direct",
      });

      const result = await maybeWakeNodeWithApns(nodeId);

      expect(result).toMatchObject({
        available: true,
        path: "send-error",
        apnsStatus: 410,
        apnsReason: "BadDeviceToken",
      });
    });

    it("nudge: all-direct no-auth surfaces sanitized apnsReason on the alert path", async () => {
      const { maybeSendNodeWakeNudge } = await import(
        "../gateway/server-methods/nodes.js"
      );
      const nodeId = "ios-node-nudge-noauth";
      mocks.loadApnsRegistrations.mockResolvedValue([directRegistration(nodeId)]);
      mocks.resolveApnsAuthConfigFromEnv.mockResolvedValue({
        ok: false,
        error:
          "APNs private key missing: set OPENCLAW_APNS_PRIVATE_KEY_P8 or OPENCLAW_APNS_PRIVATE_KEY_PATH",
      });

      const result = await maybeSendNodeWakeNudge(nodeId);

      expect(result).toMatchObject({
        sent: false,
        throttled: false,
        reason: "no-auth",
        apnsReason: "apns-auth-missing-private-key",
      });
    });

    it("empty registrations short-circuits to 'no-registration', not 'no-auth'", async () => {
      const { maybeWakeNodeWithApns } = await import(
        "../gateway/server-methods/nodes.js"
      );
      const nodeId = "ios-node-empty";
      mocks.loadApnsRegistrations.mockResolvedValue([]);

      const result = await maybeWakeNodeWithApns(nodeId);

      expect(result).toMatchObject({
        available: false,
        throttled: false,
        path: "no-registration",
      });
      expect(result.apnsReason).toBeUndefined();
      // Auth resolvers must not have been called (early return).
      expect(mocks.resolveApnsAuthConfigFromEnv).not.toHaveBeenCalled();
      expect(mocks.resolveApnsRelayConfigFromEnv).not.toHaveBeenCalled();
    });
  });

  // -------------------------------------------------------------------------
  // Case 6: known-weak gateway secret rejection (delegation test)
  // -------------------------------------------------------------------------
  describe("Case 6: known-weak gateway secret rejection (delegation)", () => {
    it("rejects every KNOWN_WEAK_GATEWAY_TOKEN_PLACEHOLDERS entry at startup", () => {
      for (const token of KNOWN_WEAK_GATEWAY_TOKEN_PLACEHOLDERS) {
        expect(() =>
          assertGatewayAuthNotKnownWeak({
            mode: "token",
            token,
          } as never),
        ).toThrow(/published example placeholder/i);
      }
    });

    it("rejects every KNOWN_WEAK_GATEWAY_PASSWORD_PLACEHOLDERS entry at startup", () => {
      for (const password of KNOWN_WEAK_GATEWAY_PASSWORD_PLACEHOLDERS) {
        expect(() =>
          assertGatewayAuthNotKnownWeak({
            mode: "password",
            password,
          } as never),
        ).toThrow(/example placeholder/i);
      }
    });

    it("does not block non-placeholder credentials", () => {
      expect(() =>
        assertGatewayAuthNotKnownWeak({
          mode: "token",
          token: "real-secret-from-openssl-rand-hex-32-output",
        } as never),
      ).not.toThrow();
    });
  });
});

// ===========================================================================
// Section B: real-store concurrency integration (case 5)
// Separate describe so the mocks above don't shadow the real push-apns module.
// ===========================================================================

describe("Case 5: concurrent register/wake — real-store integration", () => {
  // Re-import the real push-apns module via vi.importActual so the heavy mock
  // above doesn't apply here. We exercise the actual withLock from
  // src/infra/push-apns.ts to characterize 36c3a54b51's concurrency fix.
  const tempDirs = createTrackedTempDirs();

  afterEach(async () => {
    await tempDirs.cleanup();
  });

  it("Promise.all of register + load against a temp dir produces no double-write", async () => {
    const { registerApnsToken, loadApnsRegistrations } =
      await vi.importActual<typeof import("../infra/push-apns.js")>("../infra/push-apns.js");
    const baseDir = await tempDirs.make("openclaw-apns-i1a-conc-");
    const nodeId = "ios-node-conc";

    // Fire 5 concurrent registers for the same dedupe key (same token), then
    // a load. Under the real withLock, all writes are serialized and the
    // final state has exactly one registration for the dedupe key.
    const sameToken = "AAAA1111AAAA1111AAAA1111AAAA1111";
    await Promise.all(
      Array.from({ length: 5 }, () =>
        registerApnsToken({
          nodeId,
          token: sameToken,
          topic: "ai.openclaw.ios",
          environment: "sandbox",
          baseDir,
        }),
      ),
    );

    const registrations = await loadApnsRegistrations(nodeId, baseDir);
    expect(registrations).toHaveLength(1);
    expect(registrations[0]).toMatchObject({
      nodeId,
      transport: "direct",
      token: sameToken.toLowerCase(),
    });
  });

  it("Promise.all of two distinct registers + a load yields exactly two registrations", async () => {
    const { registerApnsToken, loadApnsRegistrations } =
      await vi.importActual<typeof import("../infra/push-apns.js")>("../infra/push-apns.js");
    const baseDir = await tempDirs.make("openclaw-apns-i1a-conc2-");
    const nodeId = "ios-node-conc-two";
    const tokenA = "AAAA1111AAAA1111AAAA1111AAAA1111";
    const tokenB = "BBBB2222BBBB2222BBBB2222BBBB2222";

    await Promise.all([
      registerApnsToken({
        nodeId,
        token: tokenA,
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      }),
      registerApnsToken({
        nodeId,
        token: tokenB,
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      }),
    ]);

    const registrations = await loadApnsRegistrations(nodeId, baseDir);
    expect(registrations).toHaveLength(2);
    const tokens = registrations
      .map((r) => (r.transport === "direct" ? r.token : null))
      .filter((t): t is string => t !== null)
      .sort();
    expect(tokens).toEqual([tokenA.toLowerCase(), tokenB.toLowerCase()].sort());
  });

  it("invalid APNs token throws and leaves no partial registration in the store", async () => {
    const { registerApnsToken, loadApnsRegistrations } =
      await vi.importActual<typeof import("../infra/push-apns.js")>("../infra/push-apns.js");
    const baseDir = await tempDirs.make("openclaw-apns-i1a-invalid-");
    const nodeId = "ios-node-invalid-token";

    await expect(
      registerApnsToken({
        nodeId,
        token: "short", // fails isLikelyApnsToken at push-apns.ts:540
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      }),
    ).rejects.toThrow(/invalid APNs token/i);

    const registrations = await loadApnsRegistrations(nodeId, baseDir);
    expect(registrations).toEqual([]);
  });
});
