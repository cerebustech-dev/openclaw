/**
 * Security Audit: Token & Credential Security
 *
 * Bead: driftlane-uwp — Phase 0 security gate
 * Domain D: Constant-time comparison, token revocation, rate limiting
 */
import { afterEach, beforeEach, describe, expect, it, test, vi } from "vitest";
import { safeEqualSecret } from "./secret-equal.js";
import { verifyPairingToken } from "../infra/pairing-token.js";
import { createAuthRateLimiter, type AuthRateLimiter } from "../gateway/auth-rate-limit.js";

// ===========================================================================
// D1-D5: safeEqualSecret constant-time comparison
// ===========================================================================
describe("safeEqualSecret", () => {
  it("D1: correct for matching strings", () => {
    expect(safeEqualSecret("secret-token", "secret-token")).toBe(true);
  });

  it("D2: false for near-miss", () => {
    expect(safeEqualSecret("secret-token", "secret-tokEn")).toBe(false);
  });

  it("D3: false for length mismatch (SHA256 normalization handles this)", () => {
    expect(safeEqualSecret("short", "much-longer-string")).toBe(false);
  });

  it("D4: safeEqualSecret(undefined, ...) = false", () => {
    expect(safeEqualSecret(undefined, "token")).toBe(false);
  });

  it("D5: safeEqualSecret(null, ...) = false", () => {
    expect(safeEqualSecret(null, "token")).toBe(false);
  });

  it("D5b: verifyPairingToken guards against empty-empty comparison", () => {
    // safeEqualSecret("", "") would return true (SHA256 of empty string matches)
    // but verifyPairingToken has an explicit empty guard
    expect(verifyPairingToken("", "")).toBe(false);
    expect(verifyPairingToken("   ", "   ")).toBe(false);
  });
});

// ===========================================================================
// D9-D12: Rate limiting
// ===========================================================================
describe("rate limiter", () => {
  let limiter: AuthRateLimiter;

  beforeEach(() => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-04-10T12:00:00Z"));
    limiter = createAuthRateLimiter({
      maxAttempts: 3,
      windowMs: 1_000,
      lockoutMs: 5_000,
      pruneIntervalMs: 0, // disable auto-prune for test control
    });
  });

  afterEach(() => {
    limiter.dispose();
    vi.useRealTimers();
  });

  it("D9: 3 failures -> lockout; after lockoutMs -> allowed", () => {
    const ip = "10.0.0.1";
    limiter.recordFailure(ip);
    limiter.recordFailure(ip);
    limiter.recordFailure(ip);

    const check = limiter.check(ip);
    expect(check.allowed).toBe(false);
    expect(check.retryAfterMs).toBeGreaterThan(0);
    expect(check.retryAfterMs).toBeLessThanOrEqual(5_000);

    // Advance past lockout
    vi.advanceTimersByTime(5_001);

    const checkAfter = limiter.check(ip);
    expect(checkAfter.allowed).toBe(true);
  });

  it("D10: IP-scoped — IP-A failures do not affect IP-B", () => {
    const ipA = "10.0.0.1";
    const ipB = "10.0.0.2";

    limiter.recordFailure(ipA);
    limiter.recordFailure(ipA);
    limiter.recordFailure(ipA);

    expect(limiter.check(ipA).allowed).toBe(false);
    expect(limiter.check(ipB).allowed).toBe(true);
  });

  it("D11a: loopback 127.0.0.1 exempt", () => {
    const ip = "127.0.0.1";
    limiter.recordFailure(ip);
    limiter.recordFailure(ip);
    limiter.recordFailure(ip);
    expect(limiter.check(ip).allowed).toBe(true);
  });

  it("D11b: loopback ::1 exempt", () => {
    const ip = "::1";
    limiter.recordFailure(ip);
    limiter.recordFailure(ip);
    limiter.recordFailure(ip);
    expect(limiter.check(ip).allowed).toBe(true);
  });

  it("D11c: exemptLoopback=false disables exemption", () => {
    limiter.dispose();
    limiter = createAuthRateLimiter({
      maxAttempts: 3,
      windowMs: 1_000,
      lockoutMs: 5_000,
      exemptLoopback: false,
      pruneIntervalMs: 0,
    });
    const ip = "127.0.0.1";
    limiter.recordFailure(ip);
    limiter.recordFailure(ip);
    limiter.recordFailure(ip);
    expect(limiter.check(ip).allowed).toBe(false);
  });

  it("D12: prune clears expired entries", () => {
    const ip = "10.0.0.1";
    limiter.recordFailure(ip);
    expect(limiter.size()).toBe(1);

    // Advance past window
    vi.advanceTimersByTime(2_000);
    limiter.prune();
    expect(limiter.size()).toBe(0);
  });

  it("D12b: prune retains locked-out entries", () => {
    const ip = "10.0.0.1";
    limiter.recordFailure(ip);
    limiter.recordFailure(ip);
    limiter.recordFailure(ip);

    limiter.prune();
    // Still locked out — entry retained
    expect(limiter.size()).toBe(1);

    // Advance past lockout
    vi.advanceTimersByTime(5_001);
    limiter.prune();
    expect(limiter.size()).toBe(0);
  });

  it("D9b: reset clears lockout immediately", () => {
    const ip = "10.0.0.1";
    limiter.recordFailure(ip);
    limiter.recordFailure(ip);
    limiter.recordFailure(ip);
    expect(limiter.check(ip).allowed).toBe(false);

    limiter.reset(ip);
    expect(limiter.check(ip).allowed).toBe(true);
  });
});

// ===========================================================================
// Accepted gaps — documented as test.todo
// ===========================================================================
test.todo(
  "D13: v3 metadata not cross-checked against device record — follow-up bead",
);
