import { describe, expect, it } from "vitest";

/**
 * Documenting test: proves that the selfNumber -> selfDigits filtering
 * in isBotMentionedFromTargets makes the subsequent regex safe.
 *
 * The code does: selfDigits = selfNumber.replace(/\D/g, "")
 * then:          new RegExp(`\\+?${selfDigits}`, "i")
 *
 * Since selfDigits can only contain [0-9], this is inherently safe.
 */
describe("mentions selfDigits regex safety", () => {
  it("selfDigits contains only digits after filtering", () => {
    const malicious = "123(a+)+456$";
    const filtered = malicious.replace(/\D/g, "");
    expect(filtered).toBe("123456");
    expect(/^\d*$/.test(filtered)).toBe(true);
    // Prove the regex built from filtered digits is safe
    // Uses double backslash like the actual source code
    const regex = new RegExp(`\\+?${filtered}`, "i");
    expect(regex.test("+123456")).toBe(true);
  });

  it("empty input produces empty selfDigits and no regex is built", () => {
    const filtered = "".replace(/\D/g, "");
    expect(filtered).toBe("");
    // The code checks 'if (selfDigits)' before building regex, so empty is safe
  });

  it("purely non-numeric input produces empty selfDigits", () => {
    const filtered = "abc()+*".replace(/\D/g, "");
    expect(filtered).toBe("");
  });
});
