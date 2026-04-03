import { describe, expect, it } from "vitest";
import { parseConfig, DEFAULT_SCOPES } from "./index.js";

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
});
