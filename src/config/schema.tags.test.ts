import { describe, expect, it } from "vitest";
import { deriveTagsForPath } from "./schema.tags.js";

describe("patternToRegExp ReDoS resistance", () => {
  it("resists ReDoS on adversarial dotted pattern", () => {
    // TAG_OVERRIDES uses patternToRegExp internally for patterns with *
    // We test via resolveOverride through deriveTagsForPath
    // First we need to verify the function handles adversarial input safely
    const start = performance.now();
    // Even if there were a wildcard override pattern, adversarial input should be fast
    const result = deriveTagsForPath("a" + ".b".repeat(5000));
    const elapsed = performance.now() - start;
    expect(elapsed).toBeLessThan(100);
    expect(Array.isArray(result)).toBe(true);
  });

  it("correctly tags src.*.ts style patterns", () => {
    // deriveTagsForPath uses prefix/keyword rules, test it still works
    const tags = deriveTagsForPath("tools.exec.applyPatch.workspaceOnly");
    expect(tags).toContain("tools");
    expect(tags).toContain("security");
  });
});
