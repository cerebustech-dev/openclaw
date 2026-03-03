/**
 * Centralized security policy regression tests.
 *
 * Each assertion guards a security default introduced during the security
 * remediation audit.  If any of these fail, a security-critical default has
 * regressed and MUST be investigated before merging.
 */
import { readFile } from "node:fs/promises";
import type { ServerResponse } from "node:http";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { DEFAULT_SUBAGENT_MAX_SPAWNS_PER_MINUTE } from "../config/agent-limits.js";
import { buildControlUiCspHeader } from "../gateway/control-ui-csp.js";
import { setDefaultSecurityHeaders } from "../gateway/http-common.js";
import { resolveEnableState } from "../plugins/config-state.js";
import type { NormalizedPluginsConfig } from "../plugins/config-state.js";
import { checkElevatedFullGuardrail } from "./dangerous-command-patterns.js";
import { collectEnabledInsecureOrDangerousFlags } from "./dangerous-config-flags.js";

const repoRoot = resolve(fileURLToPath(new URL(".", import.meta.url)), "..");

describe("security policy assertions", () => {
  // --- Step 1: X-Frame-Options ---
  it("sets X-Frame-Options: DENY in default security headers", () => {
    const headers: Record<string, string> = {};
    const mockRes = {
      setHeader: (name: string, value: string) => {
        headers[name] = value;
      },
    } as unknown as ServerResponse;

    setDefaultSecurityHeaders(mockRes);
    expect(headers["X-Frame-Options"]).toBe("DENY");
  });

  // --- Step 2: CSP connect-src ---
  it("CSP connect-src does not contain ws:/wss: wildcards", () => {
    const csp = buildControlUiCspHeader();
    expect(csp).toContain("connect-src 'self'");
    expect(csp).not.toMatch(/connect-src[^;]*\bws:/);
    expect(csp).not.toMatch(/connect-src[^;]*\bwss:/);
  });

  // --- Step 3: Device identity directory permissions ---
  it("device identity directories are created with mode 0o700", async () => {
    const src = await readFile(join(repoRoot, "infra/device-identity.ts"), "utf8");
    expect(src).toMatch(/mkdirSync\(.*mode:\s*0o700/s);
  });

  // --- Step 4: System prompt anti-extraction ---
  it("system prompt includes anti-extraction instruction", async () => {
    const src = await readFile(join(repoRoot, "agents/system-prompt.ts"), "utf8");
    expect(src).toMatch(/(do not|never).*(reveal|repeat|share).*system prompt/i);
  });

  // --- Step 5: CLAWDBOT_SHOW_SECRETS default ---
  it("CLAWDBOT_SHOW_SECRETS defaults to hidden (opt-in, not opt-out)", async () => {
    const src = await readFile(join(repoRoot, "commands/status.scan.ts"), "utf8");
    expect(src).toContain('=== "1"');
    expect(src).not.toContain('!== "0"');
  });

  // --- Step 6: Tool loop detection enabled ---
  it("tool loop detection is enabled by default", async () => {
    const src = await readFile(join(repoRoot, "agents/tool-loop-detection.ts"), "utf8");
    expect(src).toMatch(/DEFAULT_LOOP_DETECTION_CONFIG\s*=\s*\{[^}]*enabled:\s*true/s);
  });

  // --- Step 7: Synology allowInsecureSsl default ---
  it("Synology Chat allowInsecureSsl defaults to false", async () => {
    const src = await readFile(join(repoRoot, "../extensions/synology-chat/src/client.ts"), "utf8");
    // All function signatures should default to false
    expect(src).not.toMatch(/allowInsecureSsl\s*=\s*true/);
    expect(src).toMatch(/allowInsecureSsl\s*=\s*false/);
  });

  // --- Step 8: Docker Compose loopback bind ---
  it("Docker Compose defaults gateway bind to loopback", async () => {
    const compose = await readFile(join(repoRoot, "../docker-compose.yml"), "utf8");
    expect(compose).toContain("${OPENCLAW_GATEWAY_BIND:-loopback}");
    expect(compose).not.toContain("${OPENCLAW_GATEWAY_BIND:-lan}");
  });

  // --- Step 9: Auth mode "none" flagged as dangerous ---
  it("auth.mode=none is flagged as a dangerous config flag", () => {
    const flags = collectEnabledInsecureOrDangerousFlags({
      gateway: { auth: { mode: "none" } },
    } as never);
    expect(flags).toContain("gateway.auth.mode=none");
  });

  it("auth.mode=token is NOT flagged as dangerous", () => {
    const flags = collectEnabledInsecureOrDangerousFlags({
      gateway: { auth: { mode: "token" } },
    } as never);
    expect(flags).not.toContain("gateway.auth.mode=none");
  });

  // --- Step 12: Non-bundled plugins require explicit allow ---
  it("non-bundled plugins are blocked when allow list is empty", () => {
    const config: NormalizedPluginsConfig = {
      enabled: true,
      allow: [],
      deny: [],
      loadPaths: [],
      slots: {},
      entries: {},
    };
    const result = resolveEnableState("some-external-plugin", "global", config);
    expect(result.enabled).toBe(false);
    expect(result.reason).toContain("non-bundled");
  });

  // --- Step 13: Subagent spawn rate limit ---
  it("subagent spawn rate limit is configured", () => {
    expect(DEFAULT_SUBAGENT_MAX_SPAWNS_PER_MINUTE).toBeGreaterThan(0);
    expect(DEFAULT_SUBAGENT_MAX_SPAWNS_PER_MINUTE).toBeLessThanOrEqual(60);
  });

  // --- Step 17: dangerously prefix on allowUnsafeExternalContent ---
  it("hook types use dangerouslyAllowUnsafeExternalContent (not unprefixed)", async () => {
    const src = await readFile(join(repoRoot, "config/types.hooks.ts"), "utf8");
    expect(src).toContain("dangerouslyAllowUnsafeExternalContent");
    // Ensure the OLD unprefixed field name is not present as a type field
    expect(src).not.toMatch(/^\s+allowUnsafeExternalContent\??:/m);
  });

  // --- Step 18: Elevated exec guardrails ---
  it("elevated exec guardrails block rm -rf /", () => {
    const result = checkElevatedFullGuardrail("rm -rf /", "/tmp");
    expect(result.allowed).toBe(false);
  });

  it("elevated exec guardrails allow safe commands", () => {
    const result = checkElevatedFullGuardrail("ls -la", "/tmp");
    expect(result.allowed).toBe(true);
  });

  // --- Step 15/16: Dockerfile supply-chain hardening ---
  it("Dockerfile does not contain unpinned curl | bash", async () => {
    const dockerfile = await readFile(join(repoRoot, "../Dockerfile"), "utf8");
    expect(dockerfile).not.toMatch(/curl\s.*\|\s*bash/);
    expect(dockerfile).toContain("ARG BUN_VERSION=");
  });

  it("GitHub Actions workflow pins all uses: to SHA hashes", async () => {
    const workflow = await readFile(
      join(repoRoot, "../.github/workflows/docker-release.yml"),
      "utf8",
    );
    const usesLines = workflow
      .split("\n")
      .filter((line) => line.trim().startsWith("uses:"))
      .map((line) => line.trim());

    expect(usesLines.length).toBeGreaterThan(0);
    for (const line of usesLines) {
      expect(line).toMatch(/@[a-f0-9]{40}/);
    }
  });
});
