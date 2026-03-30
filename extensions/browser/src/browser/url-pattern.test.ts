import { describe, expect, it } from "vitest";
import { matchBrowserUrlPattern } from "./url-pattern.js";

describe("browser url pattern matching", () => {
  it("matches exact URLs", () => {
    expect(matchBrowserUrlPattern("https://example.com/a", "https://example.com/a")).toBe(true);
    expect(matchBrowserUrlPattern("https://example.com/a", "https://example.com/b")).toBe(false);
  });

  it("matches substring patterns without wildcards", () => {
    expect(matchBrowserUrlPattern("example.com", "https://example.com/a")).toBe(true);
    expect(matchBrowserUrlPattern("/dash", "https://example.com/app/dash")).toBe(true);
    expect(matchBrowserUrlPattern("nope", "https://example.com/a")).toBe(false);
  });

  it("matches glob patterns", () => {
    expect(matchBrowserUrlPattern("**/dash", "https://example.com/app/dash")).toBe(true);
    expect(matchBrowserUrlPattern("https://example.com/*", "https://example.com/a")).toBe(true);
    expect(matchBrowserUrlPattern("https://example.com/*", "https://other.com/a")).toBe(false);
  });

  it("rejects empty patterns", () => {
    expect(matchBrowserUrlPattern("", "https://example.com")).toBe(false);
    expect(matchBrowserUrlPattern("   ", "https://example.com")).toBe(false);
  });
});

describe("ReDoS resistance", () => {
  it("resists ReDoS on adversarial wildcard pattern", () => {
    const adversarial = "a" + "x".repeat(50_000);
    const start = performance.now();
    const result = matchBrowserUrlPattern("a" + "*b".repeat(15), adversarial);
    const elapsed = performance.now() - start;
    expect(elapsed).toBeLessThan(100);
    expect(result).toBe(false);
  });

  it("still matches *.example.com patterns", () => {
    expect(matchBrowserUrlPattern("*.example.com", "foo.example.com")).toBe(true);
    expect(matchBrowserUrlPattern("*.example.com", "bar.example.com")).toBe(true);
    expect(matchBrowserUrlPattern("*.example.com", "other.net")).toBe(false);
  });

  it("still matches https://*.com/* patterns", () => {
    expect(matchBrowserUrlPattern("https://*.com/*", "https://example.com/path")).toBe(true);
    expect(matchBrowserUrlPattern("https://*.com/*", "https://other.org/path")).toBe(false);
  });
});
