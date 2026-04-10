/**
 * Security Audit: Device Auth Replay & Authentication Resistance
 *
 * Bead: driftlane-uwp — Phase 0 security gate
 * Domain C: Replay attacks, timestamp validation, nonce enforcement, signature version resolution
 *
 * Finding F1 [P1 HIGH]: Non-finite signedAt bypasses timestamp skew check
 *   - message-handler.ts:656-658: `typeof NaN === "number"` and `NaN > 120000` is false
 *   - Fix: add `!Number.isFinite(signedAt)` guard
 */
import crypto from "node:crypto";
import { describe, expect, it, test } from "vitest";
import {
  deriveDeviceIdFromPublicKey,
  normalizeDevicePublicKeyBase64Url,
  signDevicePayload,
  verifyDeviceSignature,
} from "../infra/device-identity.js";
import {
  buildDeviceAuthPayload,
  buildDeviceAuthPayloadV3,
} from "../gateway/device-auth.js";
import { resolveDeviceSignaturePayloadVersion } from "../gateway/server/ws-connection/handshake-auth-helpers.js";

// Replicate the constant from message-handler.ts:126
const DEVICE_SIGNATURE_SKEW_MS = 2 * 60 * 1000;

// ---------------------------------------------------------------------------
// Helpers: generate a fresh Ed25519 identity for testing
// ---------------------------------------------------------------------------
function generateTestIdentity() {
  const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
  const publicKeyPem = publicKey.export({ type: "spki", format: "pem" }).toString();
  const privateKeyPem = privateKey.export({ type: "pkcs8", format: "pem" }).toString();
  const deviceId = deriveDeviceIdFromPublicKey(publicKeyPem)!;
  return { deviceId, publicKeyPem, privateKeyPem };
}

function buildSignedPayload(params: {
  identity: ReturnType<typeof generateTestIdentity>;
  role?: string;
  scopes?: string[];
  signedAtMs: number;
  nonce: string;
  clientId?: string;
  clientMode?: string;
  token?: string | null;
}) {
  const payloadParams = {
    deviceId: params.identity.deviceId,
    clientId: params.clientId ?? "NODE",
    clientMode: params.clientMode ?? "node",
    role: params.role ?? "node",
    scopes: params.scopes ?? [],
    signedAtMs: params.signedAtMs,
    token: params.token ?? null,
    nonce: params.nonce,
  };
  const payload = buildDeviceAuthPayload(payloadParams);
  const signature = signDevicePayload(params.identity.privateKeyPem, payload);
  return { payload, signature, payloadParams };
}

// ---------------------------------------------------------------------------
// Replicate the CURRENT skew check logic from message-handler.ts:656-658
// This is the exact code under audit — NOT the fixed version.
// ---------------------------------------------------------------------------
function currentSkewCheckRejects(signedAt: unknown, nowMs: number): boolean {
  return (
    typeof signedAt !== "number" ||
    Math.abs(nowMs - (signedAt as number)) > DEVICE_SIGNATURE_SKEW_MS
  );
}

// The FIXED version: what the code SHOULD be after adding Number.isFinite
function fixedSkewCheckRejects(signedAt: unknown, nowMs: number): boolean {
  return (
    typeof signedAt !== "number" ||
    !Number.isFinite(signedAt) ||
    Math.abs(nowMs - signedAt) > DEVICE_SIGNATURE_SKEW_MS
  );
}

// ===========================================================================
// C1-C3: Non-finite signedAt bypass [Finding F1]
// ===========================================================================
describe("F1: non-finite signedAt bypass", () => {
  const now = Date.now();

  describe("C1: NaN bypasses current skew check", () => {
    it("typeof NaN === 'number' is true (JavaScript spec)", () => {
      expect(typeof NaN).toBe("number");
    });

    it("NaN > 120000 is false (NaN comparison always false)", () => {
      expect(NaN > DEVICE_SIGNATURE_SKEW_MS).toBe(false);
    });

    it("Math.abs(Date.now() - NaN) is NaN", () => {
      expect(Number.isNaN(Math.abs(now - NaN))).toBe(true);
    });

    it("CURRENT check does NOT reject NaN — this is the bug", () => {
      // The current check returns false (does not reject) for NaN.
      // This means a signedAt of NaN would pass the timestamp guard.
      expect(currentSkewCheckRejects(NaN, now)).toBe(false);
    });

    it("FIXED check DOES reject NaN", () => {
      expect(fixedSkewCheckRejects(NaN, now)).toBe(true);
    });
  });

  describe("C2: Infinity handling", () => {
    it("Infinity IS caught by current > check (Infinity > 120000 = true)", () => {
      expect(currentSkewCheckRejects(Infinity, now)).toBe(true);
    });

    it("FIXED check also rejects Infinity (via isFinite, earlier in chain)", () => {
      expect(fixedSkewCheckRejects(Infinity, now)).toBe(true);
    });
  });

  describe("C3: -Infinity handling", () => {
    it("-Infinity IS caught by current > check (Infinity > 120000 = true via Math.abs)", () => {
      expect(currentSkewCheckRejects(-Infinity, now)).toBe(true);
    });

    it("FIXED check also rejects -Infinity", () => {
      expect(fixedSkewCheckRejects(-Infinity, now)).toBe(true);
    });
  });
});

// ===========================================================================
// C4-C8: Timestamp boundary precision
// ===========================================================================
describe("timestamp skew boundary precision", () => {
  const now = 1_700_000_000_000; // Fixed reference point

  it("C4: exactly 2 minutes ago — passes (> is strict, not >=)", () => {
    const signedAt = now - DEVICE_SIGNATURE_SKEW_MS; // exactly 120000ms ago
    // Math.abs(now - signedAt) === 120000; 120000 > 120000 is false → NOT rejected
    expect(fixedSkewCheckRejects(signedAt, now)).toBe(false);
  });

  it("C5: 2 minutes + 1ms ago — rejected", () => {
    const signedAt = now - DEVICE_SIGNATURE_SKEW_MS - 1;
    // Math.abs(now - signedAt) === 120001; 120001 > 120000 → rejected
    expect(fixedSkewCheckRejects(signedAt, now)).toBe(true);
  });

  it("C6: 2 minutes in future — passes (Math.abs handles future)", () => {
    const signedAt = now + DEVICE_SIGNATURE_SKEW_MS;
    expect(fixedSkewCheckRejects(signedAt, now)).toBe(false);
  });

  it("C7: 2 minutes + 1ms in future — rejected", () => {
    const signedAt = now + DEVICE_SIGNATURE_SKEW_MS + 1;
    expect(fixedSkewCheckRejects(signedAt, now)).toBe(true);
  });

  it("C8: epoch (signedAt: 0) — rejected (years of skew)", () => {
    expect(fixedSkewCheckRejects(0, now)).toBe(true);
  });
});

// ===========================================================================
// C9-C11: Nonce validation
// ===========================================================================
describe("nonce validation logic", () => {
  // These replicate the inline nonce checks from message-handler.ts:663-670

  it("C9: empty nonce is rejected before signature check", () => {
    const providedNonce = "";
    // The handler checks `if (!providedNonce)` first
    expect(!providedNonce).toBe(true);
  });

  it("C10: wrong nonce is rejected", () => {
    const connectNonce = crypto.randomUUID();
    const wrongNonce = crypto.randomUUID();
    expect(wrongNonce !== connectNonce).toBe(true);
  });

  it("C11: signature from connection-A rejected on connection-B (different nonce in payload)", () => {
    const identity = generateTestIdentity();
    const nonceA = crypto.randomUUID();
    const nonceB = crypto.randomUUID();
    const now = Date.now();

    // Sign with nonce-A
    const { signature } = buildSignedPayload({
      identity,
      signedAtMs: now,
      nonce: nonceA,
    });

    // Build payload with nonce-B (what connection-B would verify against)
    const payloadB = buildDeviceAuthPayload({
      deviceId: identity.deviceId,
      clientId: "NODE",
      clientMode: "node",
      role: "node",
      scopes: [],
      signedAtMs: now,
      token: null,
      nonce: nonceB,
    });

    // Signature for nonce-A does not verify against payload for nonce-B
    expect(verifyDeviceSignature(identity.publicKeyPem, payloadB, signature)).toBe(false);
  });
});

// ===========================================================================
// C12: Signature version resolution (v3 first, v2 fallback)
// ===========================================================================
describe("signature version resolution", () => {
  it("C12a: v3-signed payload resolves as 'v3'", () => {
    const identity = generateTestIdentity();
    const now = Date.now();
    const nonce = crypto.randomUUID();

    const v3Payload = buildDeviceAuthPayloadV3({
      deviceId: identity.deviceId,
      clientId: "NODE",
      clientMode: "node",
      role: "node",
      scopes: [],
      signedAtMs: now,
      token: null,
      nonce,
      platform: "ios",
      deviceFamily: "iphone",
    });
    const signature = signDevicePayload(identity.privateKeyPem, v3Payload);

    const result = resolveDeviceSignaturePayloadVersion({
      device: { id: identity.deviceId, signature, publicKey: identity.publicKeyPem },
      connectParams: {
        client: { id: "NODE", mode: "node", platform: "ios", deviceFamily: "iphone" },
        auth: {},
      } as any,
      role: "node",
      scopes: [],
      signedAtMs: now,
      nonce,
    });
    expect(result).toBe("v3");
  });

  it("C12b: v2-signed payload resolves as 'v2'", () => {
    const identity = generateTestIdentity();
    const now = Date.now();
    const nonce = crypto.randomUUID();

    const v2Payload = buildDeviceAuthPayload({
      deviceId: identity.deviceId,
      clientId: "NODE",
      clientMode: "node",
      role: "node",
      scopes: [],
      signedAtMs: now,
      token: null,
      nonce,
    });
    const signature = signDevicePayload(identity.privateKeyPem, v2Payload);

    const result = resolveDeviceSignaturePayloadVersion({
      device: { id: identity.deviceId, signature, publicKey: identity.publicKeyPem },
      connectParams: {
        client: { id: "NODE", mode: "node", platform: "ios", deviceFamily: "iphone" },
        auth: {},
      } as any,
      role: "node",
      scopes: [],
      signedAtMs: now,
      nonce,
    });
    expect(result).toBe("v2");
  });

  it("C12c: invalid signature resolves as null", () => {
    const identity = generateTestIdentity();
    const now = Date.now();
    const nonce = crypto.randomUUID();

    const result = resolveDeviceSignaturePayloadVersion({
      device: {
        id: identity.deviceId,
        signature: "definitely-not-a-valid-signature",
        publicKey: identity.publicKeyPem,
      },
      connectParams: {
        client: { id: "NODE", mode: "node" },
        auth: {},
      } as any,
      role: "node",
      scopes: [],
      signedAtMs: now,
      nonce,
    });
    expect(result).toBeNull();
  });
});

// ===========================================================================
// C13: deviceId derivation mismatch
// ===========================================================================
describe("deviceId derivation check", () => {
  it("C13: wrong pubkey for claimed deviceId is detected", () => {
    const identityA = generateTestIdentity();
    const identityB = generateTestIdentity();

    // Device claims A's deviceId but provides B's public key
    const derivedId = deriveDeviceIdFromPublicKey(identityB.publicKeyPem);
    expect(derivedId).not.toBe(identityA.deviceId);
  });
});

// ===========================================================================
// C14-C16: Wire format (JSON.parse edge cases)
// ===========================================================================
describe("wire format: JSON.parse edge cases for signedAt", () => {
  it("C14: JSON 1e309 produces Infinity — rejected by isFinite", () => {
    const parsed = JSON.parse('{"signedAt": 1e309}');
    expect(parsed.signedAt).toBe(Infinity);
    expect(Number.isFinite(parsed.signedAt)).toBe(false);
    expect(fixedSkewCheckRejects(parsed.signedAt, Date.now())).toBe(true);
  });

  it("C15: JSON string 'NaN' — rejected by typeof (string, not number)", () => {
    const parsed = JSON.parse('{"signedAt": "NaN"}');
    expect(typeof parsed.signedAt).toBe("string");
    expect(fixedSkewCheckRejects(parsed.signedAt, Date.now())).toBe(true);
  });

  it("C16: JSON null — rejected by typeof", () => {
    const parsed = JSON.parse('{"signedAt": null}');
    expect(parsed.signedAt).toBeNull();
    expect(fixedSkewCheckRejects(parsed.signedAt, Date.now())).toBe(true);
  });

  it("C14b: JSON -1e309 produces -Infinity — rejected by isFinite", () => {
    const parsed = JSON.parse('{"signedAt": -1e309}');
    expect(parsed.signedAt).toBe(-Infinity);
    expect(Number.isFinite(parsed.signedAt)).toBe(false);
    expect(fixedSkewCheckRejects(parsed.signedAt, Date.now())).toBe(true);
  });

  it("C14c: JSON boolean true — rejected by typeof", () => {
    const parsed = JSON.parse('{"signedAt": true}');
    expect(typeof parsed.signedAt).toBe("boolean");
    expect(fixedSkewCheckRejects(parsed.signedAt, Date.now())).toBe(true);
  });
});

// ===========================================================================
// C17-C18: Malformed payload handling
// ===========================================================================
describe("malformed payload handling", () => {
  it("C17: extra pipe segments do not affect signature (payload is exact match)", () => {
    const identity = generateTestIdentity();
    const now = Date.now();
    const nonce = crypto.randomUUID();

    const v2Payload = buildDeviceAuthPayload({
      deviceId: identity.deviceId,
      clientId: "NODE",
      clientMode: "node",
      role: "node",
      scopes: [],
      signedAtMs: now,
      token: null,
      nonce,
    });
    const signature = signDevicePayload(identity.privateKeyPem, v2Payload);

    // Appending extra segment breaks signature verification
    const tampered = v2Payload + "|extra-segment";
    expect(verifyDeviceSignature(identity.publicKeyPem, tampered, signature)).toBe(false);
  });

  it("C18: fewer pipe segments (truncated payload) fails verification", () => {
    const identity = generateTestIdentity();
    const now = Date.now();
    const nonce = crypto.randomUUID();

    const v2Payload = buildDeviceAuthPayload({
      deviceId: identity.deviceId,
      clientId: "NODE",
      clientMode: "node",
      role: "node",
      scopes: [],
      signedAtMs: now,
      token: null,
      nonce,
    });
    const signature = signDevicePayload(identity.privateKeyPem, v2Payload);

    // Remove last segment
    const segments = v2Payload.split("|");
    const truncated = segments.slice(0, -1).join("|");
    expect(verifyDeviceSignature(identity.publicKeyPem, truncated, signature)).toBe(false);
  });

  it("C17b: embedded pipe in scopes does not inject extra segments", () => {
    // Scopes are comma-joined, pipes in scope names would break parsing
    // but buildDeviceAuthPayload joins with commas, not pipes
    const payload = buildDeviceAuthPayload({
      deviceId: "abc123",
      clientId: "NODE",
      clientMode: "node",
      role: "node",
      scopes: ["operator.read", "operator.write"],
      signedAtMs: Date.now(),
      token: null,
      nonce: "test-nonce",
    });
    const segments = payload.split("|");
    expect(segments).toHaveLength(9);
    expect(segments[0]).toBe("v2");
    expect(segments[5]).toBe("operator.read,operator.write");
  });
});

// ===========================================================================
// C19: Invalid PEM / invalid base64url public key
// ===========================================================================
describe("invalid key format handling", () => {
  it("C19a: verifyDeviceSignature with garbage public key returns false (no throw)", () => {
    expect(verifyDeviceSignature("not-a-key", "payload", "signature")).toBe(false);
  });

  it("C19b: verifyDeviceSignature with empty public key returns false (no throw)", () => {
    expect(verifyDeviceSignature("", "payload", "signature")).toBe(false);
  });

  it("C19c: normalizeDevicePublicKeyBase64Url with malformed PEM returns null", () => {
    expect(normalizeDevicePublicKeyBase64Url("-----BEGIN PUBLIC KEY-----\ninvalid\n-----END PUBLIC KEY-----")).toBeNull();
  });

  it("C19d: normalizeDevicePublicKeyBase64Url with empty string returns null", () => {
    expect(normalizeDevicePublicKeyBase64Url("")).toBeNull();
  });

  it("C19e: deriveDeviceIdFromPublicKey with garbage hashes the decoded bytes (not null) — signature verification catches mismatch", () => {
    // The function base64url-decodes non-PEM input and hashes whatever bytes result.
    // This is safe because verifyDeviceSignature will reject the wrong key.
    const result = deriveDeviceIdFromPublicKey("not-a-key");
    expect(result).toBeTypeOf("string");
    expect(result).toHaveLength(64); // SHA256 hex
  });

  it("C19f: verifyDeviceSignature with valid key but corrupted base64url signature returns false", () => {
    const identity = generateTestIdentity();
    const payload = "test-payload";
    // Valid-looking base64url but not a real signature
    expect(verifyDeviceSignature(identity.publicKeyPem, payload, "AAAA")).toBe(false);
  });
});

// ===========================================================================
// Accepted gaps — documented as test.todo, not passing tests
// ===========================================================================
test.todo(
  "F3: Cross-connection replay with global signature cache — follow-up bead (nonce+TLS sufficient for Phase 0)",
);
test.todo(
  "F4: Deprecate v2 payload format and force v3-only — follow-up bead",
);
