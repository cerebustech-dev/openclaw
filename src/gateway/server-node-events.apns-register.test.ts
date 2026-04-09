import { createHash } from "node:crypto";
import { afterEach, beforeEach, describe, expect, it, vi, type Mock } from "vitest";
import type { NodeEventContext, NodeEvent } from "./server-node-events-types.js";
import { handleNodeEvent, _resetApnsRateLimitState } from "./server-node-events.js";

// ---------------------------------------------------------------------------
// Mocks
// ---------------------------------------------------------------------------

vi.mock("../infra/push-apns.js", () => ({
  registerApnsRegistration: vi.fn().mockResolvedValue({
    nodeId: "test-node",
    transport: "direct",
    token: "ABCD1234",
    topic: "ai.openclaw.ios",
    environment: "sandbox",
    updatedAtMs: 1,
  }),
}));

vi.mock("../infra/device-identity.js", () => ({
  loadOrCreateDeviceIdentity: vi.fn().mockReturnValue({
    deviceId: "gateway-device-1",
    privateKeyPem: "fake-pem",
  }),
}));

vi.mock("../config/config.js", () => ({
  loadConfig: vi.fn().mockReturnValue({}),
}));

// Import the mocked function so we can inspect / reset it
import { registerApnsRegistration } from "../infra/push-apns.js";

const mockRegister = registerApnsRegistration as Mock;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeCtx(overrides?: Partial<NodeEventContext>): NodeEventContext {
  return {
    deps: {} as NodeEventContext["deps"],
    broadcast: vi.fn(),
    nodeSendToSession: vi.fn(),
    nodeSubscribe: vi.fn(),
    nodeUnsubscribe: vi.fn(),
    broadcastVoiceWakeChanged: vi.fn(),
    addChatRun: vi.fn(),
    removeChatRun: vi.fn(),
    chatAbortControllers: new Map(),
    chatAbortedRuns: new Map(),
    chatRunBuffers: new Map(),
    chatDeltaSentAt: new Map(),
    dedupe: new Map(),
    agentRunSeq: new Map(),
    getHealthCache: vi.fn().mockReturnValue(null),
    refreshHealthSnapshot: vi.fn().mockResolvedValue({}),
    loadGatewayModelCatalog: vi.fn().mockResolvedValue([]),
    logGateway: { warn: vi.fn() },
    ...overrides,
  };
}

function directRegisterEvent(token = "ABCD1234ABCD1234ABCD1234ABCD1234"): NodeEvent {
  return {
    event: "push.apns.register",
    payloadJSON: JSON.stringify({
      transport: "direct",
      token,
      topic: "ai.openclaw.ios",
      environment: "sandbox",
    }),
  };
}

function relayRegisterEvent(overrides?: Record<string, unknown>): NodeEvent {
  return {
    event: "push.apns.register",
    payloadJSON: JSON.stringify({
      transport: "relay",
      relayHandle: "relay-handle-123",
      sendGrant: "grant-abc",
      installationId: "install-xyz",
      topic: "ai.openclaw.ios",
      environment: "production",
      distribution: "official",
      gatewayDeviceId: "gateway-device-1",
      ...overrides,
    }),
  };
}

function sha256Fingerprint(token: string): string {
  return createHash("sha256").update(token).digest("hex").slice(0, 16);
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("push.apns.register — rate limiting", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.clearAllMocks();
    _resetApnsRateLimitState();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it("allows up to 6 registrations from the same nodeId within 60s", async () => {
    const ctx = makeCtx();
    const nodeId = "node-rl-1";

    for (let i = 0; i < 6; i++) {
      await handleNodeEvent(ctx, nodeId, directRegisterEvent());
    }

    expect(mockRegister).toHaveBeenCalledTimes(6);
  });

  it("rejects the 7th registration from the same nodeId within 60s (silent drop)", async () => {
    const ctx = makeCtx();
    const nodeId = "node-rl-2";

    for (let i = 0; i < 7; i++) {
      await handleNodeEvent(ctx, nodeId, directRegisterEvent());
    }

    // Only 6 should have reached registerApnsRegistration
    expect(mockRegister).toHaveBeenCalledTimes(6);
  });

  it("allows registrations again after the 60s window expires", async () => {
    const ctx = makeCtx();
    const nodeId = "node-rl-3";

    // Exhaust the window
    for (let i = 0; i < 6; i++) {
      await handleNodeEvent(ctx, nodeId, directRegisterEvent());
    }
    expect(mockRegister).toHaveBeenCalledTimes(6);

    // 7th should be rejected
    await handleNodeEvent(ctx, nodeId, directRegisterEvent());
    expect(mockRegister).toHaveBeenCalledTimes(6);

    // Advance past the 60s window
    vi.advanceTimersByTime(61_000);

    // Should succeed again
    await handleNodeEvent(ctx, nodeId, directRegisterEvent());
    expect(mockRegister).toHaveBeenCalledTimes(7);
  });

  it("rejects the 11th registration from the same connection even with different nodeIds", async () => {
    const ctx = makeCtx();

    // 10 registrations from 10 different nodeIds through the same ctx (connection)
    for (let i = 0; i < 10; i++) {
      await handleNodeEvent(ctx, `conn-node-${i}`, directRegisterEvent());
    }
    expect(mockRegister).toHaveBeenCalledTimes(10);

    // 11th from yet another nodeId — should be rejected by per-connection limit
    await handleNodeEvent(ctx, "conn-node-10", directRegisterEvent());
    expect(mockRegister).toHaveBeenCalledTimes(10);
  });

  it("rejects registrations when global circuit breaker trips at 51 total", async () => {

    // 50 registrations from 50 unique nodes (stays under per-node and per-connection
    // limits by using fresh contexts for each batch of 10)
    let totalCalls = 0;
    for (let batch = 0; batch < 5; batch++) {
      const batchCtx = makeCtx();
      for (let i = 0; i < 10; i++) {
        const uniqueNode = `global-node-${batch}-${i}`;
        await handleNodeEvent(batchCtx, uniqueNode, directRegisterEvent());
        totalCalls++;
      }
    }
    expect(mockRegister).toHaveBeenCalledTimes(50);

    // 51st registration — global circuit breaker should reject
    const freshCtx = makeCtx();
    await handleNodeEvent(freshCtx, "global-node-overflow", directRegisterEvent());
    expect(mockRegister).toHaveBeenCalledTimes(50);
  });

  it("prunes rate limit state after window expiry (no memory leak)", async () => {

    // Create entries for many distinct nodeIds
    for (let i = 0; i < 20; i++) {
      const batchCtx = makeCtx();
      await handleNodeEvent(batchCtx, `prune-node-${i}`, directRegisterEvent());
    }

    // Advance well past the 60s window
    vi.advanceTimersByTime(120_000);

    // After expiry, internal state should have been pruned.
    // We verify this indirectly: a new registration should succeed even though
    // the global counter previously had 20 entries. If state wasn't pruned,
    // continued registrations would hit the global limit sooner.
    const _freshCtx = makeCtx();
    for (let i = 0; i < 50; i++) {
      const innerCtx = makeCtx();
      await handleNodeEvent(innerCtx, `post-prune-node-${i}`, directRegisterEvent());
    }

    // All 50 should succeed because the previous 20 were pruned
    // Total = 20 (before) + 50 (after) = 70 calls to registerApnsRegistration
    expect(mockRegister).toHaveBeenCalledTimes(70);
  });
});

describe("push.apns.register — audit logging", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.clearAllMocks();
    _resetApnsRateLimitState();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it("logs structured audit entry for successful direct registration", async () => {
    const ctx = makeCtx();
    const logWarn = ctx.logGateway.warn as Mock;
    const nodeId = "audit-direct-1";
    const token = "ABCD1234ABCD1234ABCD1234ABCD1234";

    await handleNodeEvent(ctx, nodeId, directRegisterEvent(token));

    // The handler should produce a structured audit log (not just warn-on-failure).
    // We expect an info-level or structured log call that includes these fields.
    // For now, we check that a log was emitted containing the audit fields.
    // The current code only logs on failure, so this test will FAIL (RED phase).
    const allCalls = logWarn.mock.calls.map((c: string[]) => c[0]);
    const auditLog = allCalls.find(
      (msg: string) =>
        msg.includes("apns.register") &&
        msg.includes(nodeId) &&
        msg.includes("direct") &&
        msg.includes("created"),
    );

    expect(auditLog).toBeDefined();
    // Must include token fingerprint (sha256 truncated), NOT raw token
    expect(auditLog).toContain(sha256Fingerprint(token));
    expect(auditLog).not.toContain(token);
    // Must include transport and topic
    expect(auditLog).toContain("direct");
    expect(auditLog).toContain("ai.openclaw.ios");
    // Must include environment for direct transport
    expect(auditLog).toContain("sandbox");
    // Must include registrationAction
    expect(auditLog).toContain("created");
  });

  it("logs structured audit entry for successful relay registration", async () => {
    const ctx = makeCtx();
    const logWarn = ctx.logGateway.warn as Mock;
    const nodeId = "audit-relay-1";

    mockRegister.mockResolvedValueOnce({
      nodeId,
      transport: "relay",
      relayHandle: "relay-handle-123",
      sendGrant: "grant-abc",
      installationId: "install-xyz",
      topic: "ai.openclaw.ios",
      environment: "production",
      distribution: "official",
      updatedAtMs: 1,
    });

    await handleNodeEvent(ctx, nodeId, relayRegisterEvent());

    const allCalls = logWarn.mock.calls.map((c: string[]) => c[0]);
    const auditLog = allCalls.find(
      (msg: string) =>
        msg.includes("apns.register") &&
        msg.includes(nodeId) &&
        msg.includes("relay"),
    );

    expect(auditLog).toBeDefined();
    // Must include transport and topic
    expect(auditLog).toContain("relay");
    expect(auditLog).toContain("ai.openclaw.ios");
    // Must include a token fingerprint (sha256 based)
    // For relay, the fingerprint is derived from relayHandle or installationId
    expect(auditLog).toMatch(/tokenFingerprint=[a-f0-9]{16}/);
  });

  it("logs a warning with nodeId and rejection reason when rate-limited", async () => {
    const ctx = makeCtx();
    const logWarn = ctx.logGateway.warn as Mock;
    const nodeId = "audit-ratelimit-1";

    // Exhaust the per-nodeId limit
    for (let i = 0; i < 6; i++) {
      await handleNodeEvent(ctx, nodeId, directRegisterEvent());
    }
    logWarn.mockClear();

    // 7th should be rate-limited and produce a warning
    await handleNodeEvent(ctx, nodeId, directRegisterEvent());

    const allCalls = logWarn.mock.calls.map((c: string[]) => c[0]);
    const rateLimitLog = allCalls.find(
      (msg: string) => msg.includes(nodeId) && msg.includes("rate"),
    );

    expect(rateLimitLog).toBeDefined();
    expect(rateLimitLog).toContain(nodeId);
    // Should mention why it was rejected (e.g., "per-node limit" or similar)
    expect(rateLimitLog).toMatch(/rate.?limit|exceeded|rejected/i);
  });

  it("uses sha256 truncated hash for token fingerprint, never raw token value", async () => {
    const ctx = makeCtx();
    const logWarn = ctx.logGateway.warn as Mock;
    const token = "DEADBEEF1234567890ABCDEF12345678";

    await handleNodeEvent(ctx, "audit-hash-1", directRegisterEvent(token));

    const allCalls = logWarn.mock.calls.map((c: string[]) => c[0]).join("\n");

    // Raw token must NEVER appear in any log output
    expect(allCalls).not.toContain(token);
    // The last 4 chars of the token should also not appear as a raw suffix
    expect(allCalls).not.toContain(token.slice(-4));

    // The sha256-based fingerprint SHOULD appear
    const expectedFingerprint = sha256Fingerprint(token);
    expect(allCalls).toContain(expectedFingerprint);
  });

  it("logs registrationAction='updated' when re-registering the same token", async () => {
    const ctx = makeCtx();
    const logWarn = ctx.logGateway.warn as Mock;
    const nodeId = "audit-update-1";
    const token = "REREGISTER_TOKEN_ABCD1234ABCD1234";

    // First registration — "created"
    await handleNodeEvent(ctx, nodeId, directRegisterEvent(token));

    const firstCalls = logWarn.mock.calls.map((c: string[]) => c[0]);
    const createdLog = firstCalls.find(
      (msg: string) => msg.includes("apns.register") && msg.includes("created"),
    );
    expect(createdLog).toBeDefined();

    logWarn.mockClear();

    // Mock returns same registration (indicating update, not new)
    mockRegister.mockResolvedValueOnce({
      nodeId,
      transport: "direct",
      token,
      topic: "ai.openclaw.ios",
      environment: "sandbox",
      updatedAtMs: 2,
      _wasUpdate: true, // implementation detail: handler checks if entry existed
    });

    // Second registration with same token — should log "updated"
    await handleNodeEvent(ctx, nodeId, directRegisterEvent(token));

    const secondCalls = logWarn.mock.calls.map((c: string[]) => c[0]);
    const updatedLog = secondCalls.find(
      (msg: string) => msg.includes("apns.register") && msg.includes("updated"),
    );
    expect(updatedLog).toBeDefined();
  });
});

// ---------------------------------------------------------------------------
// Step 5a: Identity binding monitor for direct transport
// ---------------------------------------------------------------------------

describe("push.apns.register — identity binding monitor (direct)", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.clearAllMocks();
    _resetApnsRateLimitState();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it("logs telemetry warning when direct registration omits gatewayDeviceId", async () => {
    const ctx = makeCtx();
    const logWarn = ctx.logGateway.warn as Mock;
    const nodeId = "binding-direct-1";

    // Direct registration without gatewayDeviceId
    await handleNodeEvent(ctx, nodeId, directRegisterEvent());

    const allCalls = logWarn.mock.calls.map((c: string[]) => c[0]);
    const bindingLog = allCalls.find(
      (msg: string) =>
        msg.includes("identity-binding") &&
        msg.includes(nodeId) &&
        msg.includes("missing"),
    );

    // Should log a telemetry warning about missing gatewayDeviceId
    expect(bindingLog).toBeDefined();
    // Registration should still succeed (monitor mode, not reject)
    expect(mockRegister).toHaveBeenCalledTimes(1);
  });

  it("logs telemetry warning when direct registration has mismatched gatewayDeviceId", async () => {
    const ctx = makeCtx();
    const logWarn = ctx.logGateway.warn as Mock;
    const nodeId = "binding-direct-2";

    const evt: NodeEvent = {
      event: "push.apns.register",
      payloadJSON: JSON.stringify({
        transport: "direct",
        token: "ABCD1234ABCD1234ABCD1234ABCD1234",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        gatewayDeviceId: "wrong-device-id",
      }),
    };

    await handleNodeEvent(ctx, nodeId, evt);

    const allCalls = logWarn.mock.calls.map((c: string[]) => c[0]);
    const bindingLog = allCalls.find(
      (msg: string) =>
        msg.includes("identity-binding") &&
        msg.includes(nodeId) &&
        msg.includes("mismatch"),
    );

    // Should log a telemetry warning about mismatched gatewayDeviceId
    expect(bindingLog).toBeDefined();
    // Registration should still succeed (monitor mode, not reject)
    expect(mockRegister).toHaveBeenCalledTimes(1);
  });

  it("does NOT log identity-binding warning when direct registration has correct gatewayDeviceId", async () => {
    const ctx = makeCtx();
    const logWarn = ctx.logGateway.warn as Mock;
    const nodeId = "binding-direct-3";

    const evt: NodeEvent = {
      event: "push.apns.register",
      payloadJSON: JSON.stringify({
        transport: "direct",
        token: "ABCD1234ABCD1234ABCD1234ABCD1234",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        gatewayDeviceId: "gateway-device-1", // matches mock
      }),
    };

    await handleNodeEvent(ctx, nodeId, evt);

    const allCalls = logWarn.mock.calls.map((c: string[]) => c[0]);
    const bindingLog = allCalls.find(
      (msg: string) => msg.includes("identity-binding") && msg.includes(nodeId),
    );

    // No identity-binding warning when gatewayDeviceId matches
    expect(bindingLog).toBeUndefined();
    expect(mockRegister).toHaveBeenCalledTimes(1);
  });
});

// ---------------------------------------------------------------------------
// Step 6a: Signed registration schema — telemetry
// ---------------------------------------------------------------------------

describe("push.apns.register — signed registration telemetry", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.clearAllMocks();
    _resetApnsRateLimitState();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it("logs telemetry when registration includes a deviceSignature", async () => {
    const ctx = makeCtx();
    const logWarn = ctx.logGateway.warn as Mock;
    const nodeId = "signed-1";

    const evt: NodeEvent = {
      event: "push.apns.register",
      payloadJSON: JSON.stringify({
        transport: "direct",
        token: "ABCD1234ABCD1234ABCD1234ABCD1234",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        deviceSignature: "base64-signature-data",
        deviceSignedAtMs: 1712700000000,
      }),
    };

    await handleNodeEvent(ctx, nodeId, evt);

    const allCalls = logWarn.mock.calls.map((c: string[]) => c[0]);
    const signLog = allCalls.find(
      (msg: string) =>
        msg.includes("apns.register") &&
        msg.includes(nodeId) &&
        msg.includes("signed=true"),
    );

    expect(signLog).toBeDefined();
    expect(mockRegister).toHaveBeenCalledTimes(1);

    // Verify the signature fields are passed through to registerApnsRegistration
    const callArgs = mockRegister.mock.calls[0][0];
    expect(callArgs.deviceSignature).toBe("base64-signature-data");
    expect(callArgs.deviceSignedAtMs).toBe(1712700000000);
  });

  it("logs telemetry signed=false when registration omits deviceSignature", async () => {
    const ctx = makeCtx();
    const logWarn = ctx.logGateway.warn as Mock;
    const nodeId = "unsigned-1";

    await handleNodeEvent(ctx, nodeId, directRegisterEvent());

    const allCalls = logWarn.mock.calls.map((c: string[]) => c[0]);
    const signLog = allCalls.find(
      (msg: string) =>
        msg.includes("apns.register") &&
        msg.includes(nodeId) &&
        msg.includes("signed=false"),
    );

    expect(signLog).toBeDefined();
    // Registration should still succeed (telemetry only)
    expect(mockRegister).toHaveBeenCalledTimes(1);
  });

  it("passes deviceSignature fields through to registerApnsRegistration for relay transport", async () => {
    const ctx = makeCtx();
    const nodeId = "signed-relay-1";

    const evt: NodeEvent = {
      event: "push.apns.register",
      payloadJSON: JSON.stringify({
        transport: "relay",
        relayHandle: "relay-handle-123",
        sendGrant: "grant-abc",
        installationId: "install-xyz",
        topic: "ai.openclaw.ios",
        environment: "production",
        distribution: "official",
        gatewayDeviceId: "gateway-device-1",
        deviceSignature: "relay-sig-data",
        deviceSignedAtMs: 1712700000000,
      }),
    };

    await handleNodeEvent(ctx, nodeId, evt);

    expect(mockRegister).toHaveBeenCalledTimes(1);
    const callArgs = mockRegister.mock.calls[0][0];
    expect(callArgs.deviceSignature).toBe("relay-sig-data");
    expect(callArgs.deviceSignedAtMs).toBe(1712700000000);
  });
});
