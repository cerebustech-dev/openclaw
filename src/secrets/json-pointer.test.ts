import { describe, expect, it, afterEach } from "vitest";
import { readJsonPointer, setJsonPointer } from "./json-pointer.js";

describe("setJsonPointer prototype pollution guard", () => {
  afterEach(() => {
    // Clean up any prototype pollution from canary tests
    for (const key of ["polluted", "test", "x"]) {
      delete (Object.prototype as Record<string, unknown>)[key];
    }
  });

  it("throws when pointer contains __proto__ segment", () => {
    expect(() => setJsonPointer({}, "/__proto__/polluted", true)).toThrow("Blocked path segment");
  });

  it("throws when pointer contains constructor segment", () => {
    expect(() => setJsonPointer({}, "/constructor/polluted", true)).toThrow(
      "Blocked path segment",
    );
  });

  it("throws when pointer contains prototype segment", () => {
    expect(() => setJsonPointer({}, "/prototype/polluted", true)).toThrow("Blocked path segment");
  });

  it("does not pollute Object.prototype via __proto__", () => {
    try {
      setJsonPointer({}, "/__proto__/polluted", true);
    } catch {
      // expected
    }
    expect((({}) as Record<string, unknown>).polluted).toBeUndefined();
  });

  it("sets normal nested values correctly", () => {
    const root: Record<string, unknown> = {};
    setJsonPointer(root, "/a/b", 42);
    expect((root.a as Record<string, unknown>).b).toBe(42);
  });

  it("sets top-level values correctly", () => {
    const root: Record<string, unknown> = {};
    setJsonPointer(root, "/key", "value");
    expect(root.key).toBe("value");
  });
});

describe("readJsonPointer prototype chain safety", () => {
  it("throws for __proto__ path segment (defense-in-depth)", () => {
    expect(() => readJsonPointer({}, "/__proto__")).toThrow("Blocked path segment");
  });

  it("throws for constructor path segment (defense-in-depth)", () => {
    expect(() => readJsonPointer({}, "/constructor")).toThrow("Blocked path segment");
  });

  it("reads normal nested values correctly", () => {
    const root = { a: { b: 42 } };
    expect(readJsonPointer(root, "/a/b")).toBe(42);
  });
});
