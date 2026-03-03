import { describe, expect, it, beforeEach } from "vitest";
import { addSubagentRunForTests, resetSubagentRegistryForTests } from "./subagent-registry.js";
import type { SubagentRunRecord } from "./subagent-registry.types.js";
import { checkSubagentSpawnRateLimit } from "./subagent-spawn.js";

function makeRunRecord(
  overrides: Partial<SubagentRunRecord> & { runId: string; createdAt: number },
): SubagentRunRecord {
  return {
    childSessionKey: `agent:test:subagent:${overrides.runId}`,
    requesterSessionKey: "agent:test:main",
    requesterDisplayKey: "test:main",
    task: "test task",
    cleanup: "delete",
    cleanupHandled: false,
    ...overrides,
  };
}

describe("checkSubagentSpawnRateLimit", () => {
  beforeEach(() => {
    resetSubagentRegistryForTests({ persist: false });
  });

  it("rejects when rate limit is exceeded within window", () => {
    const now = Date.now();
    // Add 10 runs created within the last minute.
    for (let i = 0; i < 10; i++) {
      addSubagentRunForTests(
        makeRunRecord({
          runId: `run-${i}`,
          createdAt: now - 30_000 + i * 100,
          requesterSessionKey: "agent:test:main",
        }),
      );
    }

    const error = checkSubagentSpawnRateLimit("agent:test:main", {
      maxSpawnsPerMinute: 10,
      nowMs: now,
    });
    expect(error).toBeDefined();
    expect(error).toMatch(/rate limit exceeded/);
    expect(error).toContain("10 spawns");
  });

  it("allows when under the limit", () => {
    const now = Date.now();
    // Add only 3 runs, well under the limit of 10.
    for (let i = 0; i < 3; i++) {
      addSubagentRunForTests(
        makeRunRecord({
          runId: `run-${i}`,
          createdAt: now - 10_000 + i * 100,
          requesterSessionKey: "agent:test:main",
        }),
      );
    }

    const error = checkSubagentSpawnRateLimit("agent:test:main", {
      maxSpawnsPerMinute: 10,
      nowMs: now,
    });
    expect(error).toBeUndefined();
  });

  it("ignores old spawns outside the 60s window", () => {
    const now = Date.now();
    // Add 15 runs, but all created more than 60s ago.
    for (let i = 0; i < 15; i++) {
      addSubagentRunForTests(
        makeRunRecord({
          runId: `old-run-${i}`,
          createdAt: now - 120_000 + i * 100,
          requesterSessionKey: "agent:test:main",
        }),
      );
    }

    const error = checkSubagentSpawnRateLimit("agent:test:main", {
      maxSpawnsPerMinute: 10,
      nowMs: now,
    });
    expect(error).toBeUndefined();
  });

  it("counts only runs within the window when mixed with old runs", () => {
    const now = Date.now();
    // 8 old runs outside the window.
    for (let i = 0; i < 8; i++) {
      addSubagentRunForTests(
        makeRunRecord({
          runId: `old-run-${i}`,
          createdAt: now - 120_000 + i * 100,
          requesterSessionKey: "agent:test:main",
        }),
      );
    }
    // 5 recent runs inside the window.
    for (let i = 0; i < 5; i++) {
      addSubagentRunForTests(
        makeRunRecord({
          runId: `new-run-${i}`,
          createdAt: now - 30_000 + i * 100,
          requesterSessionKey: "agent:test:main",
        }),
      );
    }

    const error = checkSubagentSpawnRateLimit("agent:test:main", {
      maxSpawnsPerMinute: 10,
      nowMs: now,
    });
    // Only 5 are within the window, so it should be allowed.
    expect(error).toBeUndefined();
  });

  it("isolates rate limits per session key", () => {
    const now = Date.now();
    // Saturate session A to the limit.
    for (let i = 0; i < 10; i++) {
      addSubagentRunForTests(
        makeRunRecord({
          runId: `a-run-${i}`,
          createdAt: now - 10_000 + i * 100,
          requesterSessionKey: "agent:a:main",
        }),
      );
    }

    // Session A should be rate-limited.
    expect(
      checkSubagentSpawnRateLimit("agent:a:main", { maxSpawnsPerMinute: 10, nowMs: now }),
    ).toBeDefined();

    // Session B has no runs and should be allowed.
    expect(
      checkSubagentSpawnRateLimit("agent:b:main", { maxSpawnsPerMinute: 10, nowMs: now }),
    ).toBeUndefined();
  });
});
