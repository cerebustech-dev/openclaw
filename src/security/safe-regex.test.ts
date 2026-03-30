import { describe, expect, it } from "vitest";
import {
  compileSafeRegex,
  compileSafeRegexDetailed,
  hasNestedRepetition,
  testRegexWithBoundedInput,
  buildBoundedGlobRegex,
  assertSafeRegexLiteral,
} from "./safe-regex.js";

describe("safe regex", () => {
  it.each([
    ["(a+)+$", true],
    ["(a|aa)+$", true],
    ["^(?:foo|bar)$", false],
    ["^(ab|cd)+$", false],
  ] as const)("classifies nested repetition for %s", (pattern, expected) => {
    expect(hasNestedRepetition(pattern)).toBe(expected);
  });

  it.each([
    ["(a+)+$", null],
    ["(a|aa)+$", null],
    ["(a|aa){2}$", RegExp],
  ] as const)("compiles %s safely", (pattern, expected) => {
    if (expected === null) {
      expect(compileSafeRegex(pattern)).toBeNull();
      return;
    }
    expect(compileSafeRegex(pattern)).toBeInstanceOf(expected);
  });

  it("compiles common safe filter regex", () => {
    const re = compileSafeRegex("^agent:.*:discord:");
    expect(re).toBeInstanceOf(RegExp);
    expect(re?.test("agent:main:discord:channel:123")).toBe(true);
    expect(re?.test("agent:main:telegram:channel:123")).toBe(false);
  });

  it("supports explicit flags", () => {
    const re = compileSafeRegex("token=([A-Za-z0-9]+)", "gi");
    expect(re).toBeInstanceOf(RegExp);
    expect("TOKEN=abcd1234".replace(re as RegExp, "***")).toBe("***");
  });

  it.each([
    ["   ", "empty"],
    ["(a+)+$", "unsafe-nested-repetition"],
    ["(invalid", "invalid-regex"],
    ["^agent:main$", null],
  ] as const)("returns structured reject reason for %s", (pattern, expected) => {
    expect(compileSafeRegexDetailed(pattern).reason).toBe(expected);
  });

  it.each([
    [/^agent:main:discord:/, `agent:main:discord:${"x".repeat(5000)}`, true],
    [/discord:tail$/, `${"x".repeat(5000)}discord:tail`, true],
    [/discord:tail$/, `${"x".repeat(5000)}telegram:tail`, false],
  ] as const)("checks bounded regex windows for %s", (pattern, input, expected) => {
    expect(testRegexWithBoundedInput(pattern, input)).toBe(expected);
  });
});

describe("buildBoundedGlobRegex", () => {
  it("resists ReDoS on adversarial wildcard pattern", () => {
    const regex = buildBoundedGlobRegex("a" + "*b".repeat(15));
    const adversarial = "a" + "x".repeat(50_000);
    const start = performance.now();
    const result = regex ? regex.test(adversarial) : false;
    const elapsed = performance.now() - start;
    expect(elapsed).toBeLessThan(100);
    expect(result).toBe(false);
  });

  it("returns null for empty pattern", () => {
    expect(buildBoundedGlobRegex("")).toBeNull();
  });

  it("returns null for pattern exceeding max length", () => {
    expect(buildBoundedGlobRegex("x".repeat(2000))).toBeNull();
  });

  it("matches single wildcard correctly", () => {
    const regex = buildBoundedGlobRegex("hello*world");
    expect(regex).not.toBeNull();
    expect(regex!.test("helloFOOworld")).toBe(true);
    expect(regex!.test("helloFOOworldBAR")).toBe(false);
  });

  it("matches double wildcard up to 4096 chars", () => {
    const regex = buildBoundedGlobRegex("**");
    expect(regex).not.toBeNull();
    expect(regex!.test("x".repeat(4096))).toBe(true);
    expect(regex!.test("x".repeat(4097))).toBe(false);
  });

  it("matches glob-style patterns", () => {
    const regex = buildBoundedGlobRegex("*.example.com");
    expect(regex).not.toBeNull();
    expect(regex!.test("foo.example.com")).toBe(true);
    expect(regex!.test("bar.example.com")).toBe(true);
    expect(regex!.test("example.com")).toBe(false);
  });

  it("supports case-insensitive flag", () => {
    const regex = buildBoundedGlobRegex("agent:*", { flags: "i" });
    expect(regex).not.toBeNull();
    expect(regex!.test("AGENT:foo")).toBe(true);
    expect(regex!.test("agent:bar")).toBe(true);
  });

  it("uses separator-aware wildcard when separator provided", () => {
    const regex = buildBoundedGlobRegex("src/*/file.ts", { separator: "/" });
    expect(regex).not.toBeNull();
    expect(regex!.test("src/foo/file.ts")).toBe(true);
    expect(regex!.test("src/foo/bar/file.ts")).toBe(false);
  });

  it("handles double-star with separator (crosses boundaries)", () => {
    const regex = buildBoundedGlobRegex("src/**/file.ts", { separator: "/" });
    expect(regex).not.toBeNull();
    expect(regex!.test("src/foo/bar/file.ts")).toBe(true);
  });

  it("escapes regex special characters in pattern", () => {
    const regex = buildBoundedGlobRegex("foo(bar)*");
    expect(regex).not.toBeNull();
    expect(regex!.test("foo(bar)anything")).toBe(true);
    expect(regex!.test("foobar")).toBe(false);
  });

  it("handles unicode in patterns", () => {
    const regex = buildBoundedGlobRegex("日本語*test");
    expect(regex).not.toBeNull();
    expect(regex!.test("日本語FOOtest")).toBe(true);
  });

  it("respects custom maxLength", () => {
    expect(buildBoundedGlobRegex("short", { maxLength: 3 })).toBeNull();
    expect(buildBoundedGlobRegex("ok", { maxLength: 3 })).not.toBeNull();
  });
});

describe("assertSafeRegexLiteral", () => {
  it("returns value when format matches", () => {
    expect(assertSafeRegexLiteral("12345", /^\d+$/)).toBe("12345");
  });

  it("throws when format does not match", () => {
    expect(() => assertSafeRegexLiteral("'; DROP TABLE", /^\d+$/)).toThrow();
  });

  it("throws when value exceeds maxLength", () => {
    expect(() => assertSafeRegexLiteral("x".repeat(300), /^.+$/, 256)).toThrow();
  });

  it("uses default maxLength of 256", () => {
    expect(() => assertSafeRegexLiteral("x".repeat(257), /^.+$/)).toThrow();
    expect(assertSafeRegexLiteral("x".repeat(256), /^.+$/)).toBe("x".repeat(256));
  });

  it("accepts valid formats", () => {
    expect(assertSafeRegexLiteral("abc-123", /^[\w-]+$/)).toBe("abc-123");
  });
});
