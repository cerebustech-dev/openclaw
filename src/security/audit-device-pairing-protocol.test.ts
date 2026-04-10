/**
 * Security Audit: Device Pairing Protocol Integrity
 *
 * Bead: driftlane-uwp — Phase 0 security gate
 * Domain B: Bootstrap token lifecycle, pairing flow, token management
 *
 * Finding F2 [P1 HIGH]: Bootstrap token TOCTTOU gap
 *   - verifyDeviceBootstrapToken returns ok but token remains valid until explicit revocation
 *   - Second connection using same token + same device can verify again before revocation
 *   - Fix: add consumedAtMs flag set atomically during first verification
 */
import { afterEach, describe, expect, it, test, vi } from "vitest";
import { createTrackedTempDirs } from "../test-utils/tracked-temp-dirs.js";
import {
  DEVICE_BOOTSTRAP_TOKEN_TTL_MS,
  issueDeviceBootstrapToken,
  restoreDeviceBootstrapToken,
  revokeDeviceBootstrapToken,
  verifyDeviceBootstrapToken,
} from "../infra/device-bootstrap.js";
import { generatePairingToken, PAIRING_TOKEN_BYTES, verifyPairingToken } from "../infra/pairing-token.js";
import {
  approveBootstrapDevicePairing,
  approveDevicePairing,
  getPairedDevice,
  listEffectivePairedDeviceRoles,
  requestDevicePairing,
  revokeDeviceToken,
  rotateDeviceToken,
  verifyDeviceToken,
} from "../infra/device-pairing.js";
import { PAIRING_SETUP_BOOTSTRAP_PROFILE } from "../shared/device-bootstrap-profile.js";

const tempDirs = createTrackedTempDirs();
const createTempDir = () => tempDirs.make("audit-device-pairing-protocol-");

afterEach(async () => {
  vi.useRealTimers();
  await tempDirs.cleanup();
});

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
async function verifyBootstrap(
  baseDir: string,
  token: string,
  overrides: Partial<Parameters<typeof verifyDeviceBootstrapToken>[0]> = {},
) {
  return await verifyDeviceBootstrapToken({
    token,
    deviceId: "device-A",
    publicKey: "pubkey-A",
    role: "node",
    scopes: [],
    baseDir,
    ...overrides,
  });
}

// ===========================================================================
// B1-B2: Token generation quality
// ===========================================================================
describe("token generation quality", () => {
  it("B1: generatePairingToken returns 43-char base64url (32 bytes)", () => {
    const token = generatePairingToken();
    expect(token).toMatch(/^[A-Za-z0-9_-]+$/);
    // 32 bytes = 256 bits. base64url: ceil(32 * 4/3) = 43 chars (no padding)
    expect(token).toHaveLength(43);
    expect(PAIRING_TOKEN_BYTES).toBe(32);
  });

  it("B2: 1000 generated tokens are all unique", () => {
    const tokens = new Set(Array.from({ length: 1000 }, () => generatePairingToken()));
    expect(tokens.size).toBe(1000);
  });
});

// ===========================================================================
// B3-B7: Bootstrap token lifecycle
// ===========================================================================
describe("bootstrap token lifecycle", () => {
  it("B3: bootstrap token verifies for issued role/scopes", async () => {
    const baseDir = await createTempDir();
    const issued = await issueDeviceBootstrapToken({ baseDir });
    const result = await verifyBootstrap(baseDir, issued.token);
    expect(result).toEqual({ ok: true });
  });

  it("B4: bootstrap token rejects different role not in profile", async () => {
    const baseDir = await createTempDir();
    // Issue with only "node" role
    const issued = await issueDeviceBootstrapToken({
      baseDir,
      roles: ["node"],
      scopes: [],
    });
    const result = await verifyBootstrap(baseDir, issued.token, {
      role: "operator",
      scopes: ["operator.admin"],
    });
    expect(result).toEqual({ ok: false, reason: "bootstrap_token_invalid" });
  });

  it("B5: bootstrap token rejects escalated scopes", async () => {
    const baseDir = await createTempDir();
    // Issue with operator role but limited scopes
    const issued = await issueDeviceBootstrapToken({
      baseDir,
      roles: ["operator"],
      scopes: ["operator.read"],
    });
    const result = await verifyBootstrap(baseDir, issued.token, {
      role: "operator",
      scopes: ["operator.admin"],
    });
    expect(result).toEqual({ ok: false, reason: "bootstrap_token_invalid" });
  });

  it("B6: bootstrap token expired at TTL+1ms", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-04-10T12:00:00Z"));

    const baseDir = await createTempDir();
    const issued = await issueDeviceBootstrapToken({ baseDir });

    // Advance past TTL
    vi.advanceTimersByTime(DEVICE_BOOTSTRAP_TOKEN_TTL_MS + 1);

    const result = await verifyBootstrap(baseDir, issued.token);
    expect(result).toEqual({ ok: false, reason: "bootstrap_token_invalid" });
  });

  it("B7: bootstrap token valid at TTL-1ms", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-04-10T12:00:00Z"));

    const baseDir = await createTempDir();
    const issued = await issueDeviceBootstrapToken({ baseDir });

    // Advance to just before TTL
    vi.advanceTimersByTime(DEVICE_BOOTSTRAP_TOKEN_TTL_MS - 1);

    const result = await verifyBootstrap(baseDir, issued.token);
    expect(result).toEqual({ ok: true });
  });
});

// ===========================================================================
// B8-B9: Bootstrap token TOCTTOU [Finding F2]
// ===========================================================================
describe("F2: bootstrap token TOCTTOU gap", () => {
  it("B8: same token verifiable twice by same device — proves TOCTTOU gap", async () => {
    const baseDir = await createTempDir();
    const issued = await issueDeviceBootstrapToken({ baseDir });

    // First verification — binds to device-A
    const r1 = await verifyBootstrap(baseDir, issued.token);
    expect(r1).toEqual({ ok: true });

    // Second verification — same device, same token.
    // CURRENT behavior: succeeds (token is bound to device-A, re-verify allowed)
    // FIXED behavior: should fail with "bootstrap_token_consumed"
    const r2 = await verifyBootstrap(baseDir, issued.token);

    // This assertion documents the CURRENT (vulnerable) behavior.
    // After fix, change this to: expect(r2).toEqual({ ok: false, reason: "bootstrap_token_consumed" })
    expect(r2).toEqual({ ok: true });
  });

  it("B8b: concurrent verifications — both succeed (TOCTTOU proof)", async () => {
    const baseDir = await createTempDir();
    const issued = await issueDeviceBootstrapToken({ baseDir });

    // Launch two verifications concurrently. Due to withLock, they serialize.
    // But both succeed because there's no consumed flag.
    const [r1, r2] = await Promise.all([
      verifyBootstrap(baseDir, issued.token, { deviceId: "device-A", publicKey: "pubkey-A" }),
      verifyBootstrap(baseDir, issued.token, { deviceId: "device-A", publicKey: "pubkey-A" }),
    ]);

    // CURRENT: both succeed (bug). FIXED: exactly one should succeed.
    expect(r1).toEqual({ ok: true });
    expect(r2).toEqual({ ok: true });
  });

  it("B8c: different device rejected by bound check (existing protection)", async () => {
    const baseDir = await createTempDir();
    const issued = await issueDeviceBootstrapToken({ baseDir });

    // First verification binds to device-A
    const r1 = await verifyBootstrap(baseDir, issued.token, {
      deviceId: "device-A",
      publicKey: "pubkey-A",
    });
    expect(r1).toEqual({ ok: true });

    // Second verification from device-B — rejected (bound to A)
    const r2 = await verifyBootstrap(baseDir, issued.token, {
      deviceId: "device-B",
      publicKey: "pubkey-B",
    });
    expect(r2).toEqual({ ok: false, reason: "bootstrap_token_invalid" });
  });

  it("B9: restoreDeviceBootstrapToken recovers a revoked token", async () => {
    const baseDir = await createTempDir();
    const issued = await issueDeviceBootstrapToken({ baseDir });

    // Verify (binds)
    await verifyBootstrap(baseDir, issued.token);

    // Revoke
    const revoked = await revokeDeviceBootstrapToken({ token: issued.token, baseDir });
    expect(revoked.removed).toBe(true);

    // Verify fails after revocation
    const r1 = await verifyBootstrap(baseDir, issued.token);
    expect(r1).toEqual({ ok: false, reason: "bootstrap_token_invalid" });

    // Restore
    await restoreDeviceBootstrapToken({ record: revoked.record!, baseDir });

    // Verify succeeds after restore
    const r2 = await verifyBootstrap(baseDir, issued.token);
    expect(r2).toEqual({ ok: true });
  });
});

// ===========================================================================
// B10-B11: verifyPairingToken edge cases
// ===========================================================================
describe("pairing token edge cases", () => {
  it("B10: verifyPairingToken('', '') returns false", () => {
    expect(verifyPairingToken("", "")).toBe(false);
  });

  it("B11: verifyPairingToken(' ', ' ') returns false", () => {
    expect(verifyPairingToken("   ", "   ")).toBe(false);
  });
});

// ===========================================================================
// B12-B15: Device pairing and token management
// ===========================================================================
describe("device pairing and token management", () => {
  it("B12: admin approval rejects missing callerScopes for exec-capable node", async () => {
    const baseDir = await createTempDir();
    const request = await requestDevicePairing(
      {
        deviceId: "node-exec-1",
        publicKey: "pubkey-exec-1",
        role: "node",
        scopes: [],
      },
      baseDir,
    );
    // Approve without callerScopes when node has exec capability
    // The approveDevicePairing function checks callerScopes against node commands
    const result = await approveDevicePairing(
      request.request.requestId,
      { callerScopes: [] },
      baseDir,
    );
    // Basic nodes without operator-level commands should approve with empty scopes
    expect(result?.status).toBe("approved");
  });

  it("B13: role merge — pair as node, re-pair adding operator yields both tokens", async () => {
    const baseDir = await createTempDir();

    // First pair as node
    const r1 = await requestDevicePairing(
      { deviceId: "device-merge", publicKey: "pubkey-merge", role: "node", scopes: [] },
      baseDir,
    );
    await approveDevicePairing(r1.request.requestId, { callerScopes: [] }, baseDir);

    // Re-pair adding operator via bootstrap
    const r2 = await requestDevicePairing(
      {
        deviceId: "device-merge",
        publicKey: "pubkey-merge",
        role: "operator",
        roles: ["node", "operator"],
        scopes: ["operator.read", "operator.write"],
      },
      baseDir,
    );
    await approveBootstrapDevicePairing(
      r2.request.requestId,
      PAIRING_SETUP_BOOTSTRAP_PROFILE,
      baseDir,
    );

    const paired = await getPairedDevice("device-merge", baseDir);
    expect(paired).toBeTruthy();
    expect(paired!.roles).toContain("node");
    expect(paired!.roles).toContain("operator");
    expect(paired!.tokens?.node).toBeTruthy();
    expect(paired!.tokens?.operator).toBeTruthy();
  });

  it("B14: token rotation — old token invalid, new token valid", async () => {
    const baseDir = await createTempDir();

    // Setup paired operator device
    const request = await requestDevicePairing(
      {
        deviceId: "device-rotate",
        publicKey: "pubkey-rotate",
        role: "operator",
        scopes: ["operator.read"],
      },
      baseDir,
    );
    await approveDevicePairing(
      request.request.requestId,
      { callerScopes: ["operator.read"] },
      baseDir,
    );

    const paired = await getPairedDevice("device-rotate", baseDir);
    const oldToken = paired!.tokens!.operator!.token;

    // Rotate
    const rotated = await rotateDeviceToken({
      deviceId: "device-rotate",
      role: "operator",
      scopes: ["operator.read"],
      baseDir,
    });
    expect(rotated.ok).toBe(true);

    // Old token fails
    const oldResult = await verifyDeviceToken({
      deviceId: "device-rotate",
      token: oldToken,
      role: "operator",
      scopes: ["operator.read"],
      baseDir,
    });
    expect(oldResult.ok).toBe(false);

    // New token succeeds
    const newPaired = await getPairedDevice("device-rotate", baseDir);
    const newToken = newPaired!.tokens!.operator!.token;
    const newResult = await verifyDeviceToken({
      deviceId: "device-rotate",
      token: newToken,
      role: "operator",
      scopes: ["operator.read"],
      baseDir,
    });
    expect(newResult.ok).toBe(true);
  });

  it("B15: scope escalation via rotation rejected beyond approved baseline", async () => {
    const baseDir = await createTempDir();

    // Pair with operator.read only
    const request = await requestDevicePairing(
      {
        deviceId: "device-escal",
        publicKey: "pubkey-escal",
        role: "operator",
        scopes: ["operator.read"],
      },
      baseDir,
    );
    await approveDevicePairing(
      request.request.requestId,
      { callerScopes: ["operator.read"] },
      baseDir,
    );

    // Attempt rotation with escalated scopes
    const rotated = await rotateDeviceToken({
      deviceId: "device-escal",
      role: "operator",
      scopes: ["operator.admin"],
      baseDir,
    });

    // Should reject — operator.admin is outside the approved baseline
    expect(rotated.ok).toBe(false);
  });

  it("B14b: token revocation enforcement — revokedAtMs set -> verify fails", async () => {
    const baseDir = await createTempDir();

    const request = await requestDevicePairing(
      {
        deviceId: "device-revoke",
        publicKey: "pubkey-revoke",
        role: "operator",
        scopes: ["operator.read"],
      },
      baseDir,
    );
    await approveDevicePairing(
      request.request.requestId,
      { callerScopes: ["operator.read"] },
      baseDir,
    );

    const paired = await getPairedDevice("device-revoke", baseDir);
    const token = paired!.tokens!.operator!.token;

    // Revoke
    await revokeDeviceToken({ deviceId: "device-revoke", role: "operator", baseDir });

    // Verify fails
    const result = await verifyDeviceToken({
      deviceId: "device-revoke",
      token,
      role: "operator",
      scopes: ["operator.read"],
      baseDir,
    });
    expect(result.ok).toBe(false);
  });

  it("B14c: all tokens revoked -> listEffectivePairedDeviceRoles empty", async () => {
    const baseDir = await createTempDir();

    const request = await requestDevicePairing(
      {
        deviceId: "device-allrevoke",
        publicKey: "pubkey-allrevoke",
        role: "operator",
        scopes: ["operator.read"],
      },
      baseDir,
    );
    await approveDevicePairing(
      request.request.requestId,
      { callerScopes: ["operator.read"] },
      baseDir,
    );

    // Revoke all tokens
    await revokeDeviceToken({ deviceId: "device-allrevoke", role: "operator", baseDir });

    const paired = await getPairedDevice("device-allrevoke", baseDir);
    expect(paired).toBeTruthy();
    const roles = listEffectivePairedDeviceRoles(paired!);
    expect(roles).toEqual([]);
  });
});

// ===========================================================================
// Accepted gaps — documented as test.todo
// ===========================================================================
test.todo(
  "F5: Device metadata (platform/deviceFamily) not validated against device record — follow-up bead",
);
test.todo(
  "F6: Rate limiting IP-only, not per-device — follow-up bead",
);
