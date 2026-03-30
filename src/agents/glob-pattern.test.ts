import { describe, expect, it } from "vitest";
import { compileGlobPattern } from "./glob-pattern.js";

const identity = (v: string) => v;

describe("compileGlobPattern ReDoS resistance", () => {
  it("resists ReDoS on adversarial wildcard pattern", () => {
    const pattern = "a" + "*b".repeat(15);
    const compiled = compileGlobPattern({ raw: pattern, normalize: identity });
    const adversarial = "a" + "x".repeat(50_000);
    const start = performance.now();
    if (compiled.kind === "regex") {
      compiled.value.test(adversarial);
    } else if (compiled.kind === "exact") {
      // exact match is inherently safe
    }
    const elapsed = performance.now() - start;
    expect(elapsed).toBeLessThan(100);
  });

  it("matches *.ts against foo.ts", () => {
    const compiled = compileGlobPattern({ raw: "*.ts", normalize: identity });
    expect(compiled.kind).toBe("regex");
    if (compiled.kind === "regex") {
      expect(compiled.value.test("foo.ts")).toBe(true);
      expect(compiled.value.test("foo.js")).toBe(false);
    }
  });

  it("matches ** against anything", () => {
    const compiled = compileGlobPattern({ raw: "**", normalize: identity });
    // ** should match anything — could be 'all' or 'regex'
    if (compiled.kind === "regex") {
      expect(compiled.value.test("anything/here")).toBe(true);
    } else {
      // 'all' kind also matches everything
      expect(compiled.kind).toBe("all");
    }
  });

  it("exact match for patterns without wildcards", () => {
    const compiled = compileGlobPattern({ raw: "exact", normalize: identity });
    expect(compiled.kind).toBe("exact");
    expect(compiled).toEqual({ kind: "exact", value: "exact" });
  });

  it("falls back gracefully for overly long patterns", () => {
    const longPattern = "a" + "*".repeat(2000);
    const compiled = compileGlobPattern({ raw: longPattern, normalize: identity });
    // Should not throw, and should either be safe regex or fallback to exact
    expect(["regex", "exact", "all"]).toContain(compiled.kind);
  });
});
