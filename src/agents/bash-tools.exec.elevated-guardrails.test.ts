import { describe, expect, it } from "vitest";
import { checkElevatedFullGuardrail } from "../security/dangerous-command-patterns.js";

const WORKSPACE = "/home/user/project";

describe("elevated-full exec guardrails", () => {
  describe("denylist patterns", () => {
    it("blocks rm -rf /", () => {
      const result = checkElevatedFullGuardrail("rm -rf /", WORKSPACE);
      expect(result.allowed).toBe(false);
    });

    it("blocks mkfs", () => {
      const result = checkElevatedFullGuardrail("mkfs.ext4 /dev/sda1", WORKSPACE);
      expect(result.allowed).toBe(false);
    });

    it("blocks dd to block device", () => {
      const result = checkElevatedFullGuardrail("dd if=/dev/zero of=/dev/sda bs=1M", WORKSPACE);
      expect(result.allowed).toBe(false);
    });

    it("blocks fork bombs", () => {
      const result = checkElevatedFullGuardrail(":() { :|:& }; :", WORKSPACE);
      expect(result.allowed).toBe(false);
    });
  });

  describe("obfuscation variants", () => {
    it("blocks backslash-escaped rm", () => {
      const result = checkElevatedFullGuardrail("r\\m -rf /", WORKSPACE);
      expect(result.allowed).toBe(false);
    });

    it("blocks hex-escaped rm via $'...'", () => {
      const result = checkElevatedFullGuardrail("$'\\x72m' -rf /", WORKSPACE);
      expect(result.allowed).toBe(false);
    });
  });

  describe("context checks", () => {
    it("blocks destructive paths outside workspace (rm -rf /etc)", () => {
      const result = checkElevatedFullGuardrail("rm -rf /etc", WORKSPACE);
      expect(result.allowed).toBe(false);
    });

    it("allows rm -rf inside workspace (rm -rf ./build)", () => {
      const result = checkElevatedFullGuardrail("rm -rf ./build", WORKSPACE);
      expect(result.allowed).toBe(true);
    });
  });

  describe("safe patterns", () => {
    it("allows ls -la", () => {
      const result = checkElevatedFullGuardrail("ls -la", WORKSPACE);
      expect(result.allowed).toBe(true);
    });

    it("allows cat file", () => {
      const result = checkElevatedFullGuardrail("cat /tmp/test.txt", WORKSPACE);
      expect(result.allowed).toBe(true);
    });

    it("allows rm -rf ./build (relative workspace path)", () => {
      const result = checkElevatedFullGuardrail("rm -rf ./build", WORKSPACE);
      expect(result.allowed).toBe(true);
    });
  });

  describe("audit logging", () => {
    it("returns allowed: true for safe commands (audit log emitted internally)", () => {
      const result = checkElevatedFullGuardrail("echo hello", WORKSPACE);
      expect(result.allowed).toBe(true);
    });

    it("returns allowed: false with reason for blocked commands", () => {
      const result = checkElevatedFullGuardrail("rm -rf /", WORKSPACE);
      expect(result.allowed).toBe(false);
      if (!result.allowed) {
        expect(result.reason).toContain("guardrail");
      }
    });
  });
});
