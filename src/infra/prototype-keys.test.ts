import { describe, expect, it } from "vitest";
import { assertSafeObjectKey, assertSafePathSegments, isBlockedObjectKey } from "./prototype-keys.js";

describe("isBlockedObjectKey", () => {
  it("blocks prototype-pollution keys and allows ordinary keys", () => {
    for (const key of ["__proto__", "prototype", "constructor"]) {
      expect(isBlockedObjectKey(key)).toBe(true);
    }

    for (const key of ["toString", "value", "constructorName", "__proto__x", "Prototype"]) {
      expect(isBlockedObjectKey(key)).toBe(false);
    }
  });
});

describe("assertSafePathSegments", () => {
  it("throws for __proto__ segment", () => {
    expect(() => assertSafePathSegments(["foo", "__proto__", "bar"])).toThrow(
      "Blocked path segment",
    );
  });

  it("throws for constructor segment", () => {
    expect(() => assertSafePathSegments(["constructor"])).toThrow("Blocked path segment");
  });

  it("throws for prototype segment", () => {
    expect(() => assertSafePathSegments(["a", "prototype"])).toThrow("Blocked path segment");
  });

  it("allows safe path segments", () => {
    expect(() => assertSafePathSegments(["providers", "openai", "apiKey"])).not.toThrow();
  });

  it("allows empty segments array", () => {
    expect(() => assertSafePathSegments([])).not.toThrow();
  });

  it("includes the blocked key name in the error message", () => {
    expect(() => assertSafePathSegments(["__proto__"])).toThrow(/"__proto__"/);
  });
});

describe("assertSafeObjectKey", () => {
  it("throws for __proto__", () => {
    expect(() => assertSafeObjectKey("__proto__")).toThrow("Blocked object key");
  });

  it("throws for constructor", () => {
    expect(() => assertSafeObjectKey("constructor")).toThrow("Blocked object key");
  });

  it("throws for prototype", () => {
    expect(() => assertSafeObjectKey("prototype")).toThrow("Blocked object key");
  });

  it("allows safe keys", () => {
    expect(() => assertSafeObjectKey("apiKey")).not.toThrow();
    expect(() => assertSafeObjectKey("constructorName")).not.toThrow();
  });
});
