import { afterEach, describe, expect, it } from "vitest";
import type { OpenClawConfig } from "../config/config.js";
import {
  deletePathStrict,
  getPath,
  setPathCreateStrict,
  setPathExistingStrict,
} from "./path-utils.js";

function asConfig(value: unknown): OpenClawConfig {
  return value as OpenClawConfig;
}

function createAgentListConfig(): OpenClawConfig {
  return asConfig({
    agents: {
      list: [{ id: "a" }],
    },
  });
}

describe("secrets path utils", () => {
  it("deletePathStrict compacts arrays via splice", () => {
    const config = asConfig({});
    setPathCreateStrict(config, ["agents", "list"], [{ id: "a" }, { id: "b" }, { id: "c" }]);
    const changed = deletePathStrict(config, ["agents", "list", "1"]);
    expect(changed).toBe(true);
    expect(getPath(config, ["agents", "list"])).toEqual([{ id: "a" }, { id: "c" }]);
  });

  it("getPath returns undefined for invalid array path segment", () => {
    const config = asConfig({
      agents: {
        list: [{ id: "a" }],
      },
    });
    expect(getPath(config, ["agents", "list", "foo"])).toBeUndefined();
  });

  it("setPathExistingStrict throws when path does not already exist", () => {
    const config = createAgentListConfig();
    expect(() =>
      setPathExistingStrict(
        config,
        ["agents", "list", "0", "memorySearch", "remote", "apiKey"],
        "x",
      ),
    ).toThrow(/Path segment does not exist/);
  });

  it("setPathExistingStrict updates an existing leaf", () => {
    const config = asConfig({
      talk: {
        apiKey: "old", // pragma: allowlist secret
      },
    });
    const changed = setPathExistingStrict(config, ["talk", "apiKey"], "new");
    expect(changed).toBe(true);
    expect(getPath(config, ["talk", "apiKey"])).toBe("new");
  });

  it("setPathCreateStrict creates missing container segments", () => {
    const config = asConfig({});
    const changed = setPathCreateStrict(config, ["talk", "provider", "apiKey"], "x");
    expect(changed).toBe(true);
    expect(getPath(config, ["talk", "provider", "apiKey"])).toBe("x");
  });

  it("setPathCreateStrict leaves value unchanged when equal", () => {
    const config = asConfig({
      talk: {
        apiKey: "same", // pragma: allowlist secret
      },
    });
    const changed = setPathCreateStrict(config, ["talk", "apiKey"], "same");
    expect(changed).toBe(false);
    expect(getPath(config, ["talk", "apiKey"])).toBe("same");
  });
});

describe("prototype pollution guards", () => {
  afterEach(() => {
    for (const key of ["polluted", "test", "evil", "x"]) {
      delete (Object.prototype as Record<string, unknown>)[key];
    }
  });

  describe("setPathCreateStrict", () => {
    it("throws on __proto__ segment", () => {
      expect(() => setPathCreateStrict(asConfig({}), ["__proto__", "polluted"], true)).toThrow(
        "Blocked path segment",
      );
    });

    it("throws on constructor segment", () => {
      expect(() => setPathCreateStrict(asConfig({}), ["constructor"], true)).toThrow(
        "Blocked path segment",
      );
    });

    it("throws on prototype segment", () => {
      expect(() => setPathCreateStrict(asConfig({}), ["a", "prototype", "b"], true)).toThrow(
        "Blocked path segment",
      );
    });

    it("does not pollute Object.prototype after attempted __proto__ write", () => {
      try {
        setPathCreateStrict(asConfig({}), ["providers", "__proto__", "polluted"], "EXPLOIT");
      } catch {
        // expected
      }
      expect((({}) as Record<string, unknown>).polluted).toBeUndefined();
    });
  });

  describe("setPathExistingStrict", () => {
    it("throws on __proto__ segment", () => {
      expect(() =>
        setPathExistingStrict(asConfig({ a: { b: 1 } }), ["__proto__"], "evil"),
      ).toThrow("Blocked path segment");
    });

    it("throws on constructor segment", () => {
      expect(() =>
        setPathExistingStrict(asConfig({ constructor: 1 }), ["constructor"], "evil"),
      ).toThrow("Blocked path segment");
    });
  });

  describe("deletePathStrict", () => {
    it("throws on __proto__ segment", () => {
      expect(() => deletePathStrict(asConfig({}), ["__proto__"])).toThrow("Blocked path segment");
    });

    it("throws on prototype segment", () => {
      expect(() => deletePathStrict(asConfig({}), ["prototype"])).toThrow("Blocked path segment");
    });
  });

  describe("getPath", () => {
    it("throws on __proto__ segment instead of leaking Object.prototype", () => {
      expect(() => getPath({}, ["__proto__"])).toThrow("Blocked path segment");
    });

    it("throws on constructor segment instead of leaking Object constructor", () => {
      expect(() => getPath({}, ["constructor"])).toThrow("Blocked path segment");
    });

    it("still reads normal paths correctly", () => {
      const root = { a: { b: { c: 42 } } };
      expect(getPath(root, ["a", "b", "c"])).toBe(42);
    });
  });
});
