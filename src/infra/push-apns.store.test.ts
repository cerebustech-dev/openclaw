import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  clearApnsRegistration,
  clearApnsRegistrationIfCurrent,
  loadApnsRegistration,
  loadApnsRegistrations,
  registerApnsRegistration,
  registerApnsToken,
} from "./push-apns.js";

const tempDirs: string[] = [];

async function makeTempDir(): Promise<string> {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "openclaw-push-apns-store-test-"));
  tempDirs.push(dir);
  return dir;
}

afterEach(async () => {
  while (tempDirs.length > 0) {
    const dir = tempDirs.pop();
    if (dir) {
      await fs.rm(dir, { recursive: true, force: true });
    }
  }
});

describe("push APNs registration store", () => {
  it("stores and reloads direct APNs registrations", async () => {
    const baseDir = await makeTempDir();
    const saved = await registerApnsToken({
      nodeId: "ios-node-1",
      token: "ABCD1234ABCD1234ABCD1234ABCD1234",
      topic: "ai.openclaw.ios",
      environment: "sandbox",
      baseDir,
    });

    const loaded = await loadApnsRegistration("ios-node-1", baseDir);
    expect(loaded).toMatchObject({
      nodeId: "ios-node-1",
      transport: "direct",
      topic: "ai.openclaw.ios",
      environment: "sandbox",
      updatedAtMs: saved.updatedAtMs,
    });
    expect(loaded && loaded.transport === "direct" ? loaded.token : null).toBe(
      "abcd1234abcd1234abcd1234abcd1234",
    );
  });

  it("stores relay-backed registrations without a raw token", async () => {
    const baseDir = await makeTempDir();
    const saved = await registerApnsRegistration({
      nodeId: "ios-node-relay",
      transport: "relay",
      relayHandle: "relay-handle-123",
      sendGrant: "send-grant-123",
      installationId: "install-123",
      topic: "ai.openclaw.ios",
      environment: "production",
      distribution: "official",
      tokenDebugSuffix: " abcd-1234 ",
      baseDir,
    });

    const loaded = await loadApnsRegistration("ios-node-relay", baseDir);
    expect(saved.transport).toBe("relay");
    expect(loaded).toMatchObject({
      nodeId: "ios-node-relay",
      transport: "relay",
      relayHandle: "relay-handle-123",
      sendGrant: "send-grant-123",
      installationId: "install-123",
      topic: "ai.openclaw.ios",
      environment: "production",
      distribution: "official",
      tokenDebugSuffix: "abcd1234",
    });
    expect(loaded && "token" in loaded).toBe(false);
  });

  it("normalizes legacy direct records from disk and ignores invalid entries", async () => {
    const baseDir = await makeTempDir();
    const statePath = path.join(baseDir, "push", "apns-registrations.json");
    await fs.mkdir(path.dirname(statePath), { recursive: true });
    await fs.writeFile(
      statePath,
      `${JSON.stringify(
        {
          registrationsByNodeId: {
            " ios-node-legacy ": {
              nodeId: " ios-node-legacy ",
              token: "<ABCD1234ABCD1234ABCD1234ABCD1234>",
              topic: " ai.openclaw.ios ",
              environment: " PRODUCTION ",
              updatedAtMs: 3,
            },
            "   ": {
              nodeId: " ios-node-fallback ",
              token: "<ABCD1234ABCD1234ABCD1234ABCD1234>",
              topic: " ai.openclaw.ios ",
              updatedAtMs: 2,
            },
            "ios-node-bad-relay": {
              transport: "relay",
              nodeId: "ios-node-bad-relay",
              relayHandle: "relay-handle-123",
              sendGrant: "send-grant-123",
              installationId: "install-123",
              topic: "ai.openclaw.ios",
              environment: "production",
              distribution: "beta",
              updatedAtMs: 1,
            },
          },
        },
        null,
        2,
      )}\n`,
      "utf8",
    );

    await expect(loadApnsRegistration("ios-node-legacy", baseDir)).resolves.toMatchObject({
      nodeId: "ios-node-legacy",
      transport: "direct",
      token: "abcd1234abcd1234abcd1234abcd1234",
      topic: "ai.openclaw.ios",
      environment: "production",
      updatedAtMs: 3,
    });
    await expect(loadApnsRegistration("ios-node-fallback", baseDir)).resolves.toMatchObject({
      nodeId: "ios-node-fallback",
      transport: "direct",
      topic: "ai.openclaw.ios",
      environment: "sandbox",
      updatedAtMs: 2,
    });
    await expect(loadApnsRegistration("ios-node-bad-relay", baseDir)).resolves.toBeNull();
  });

  it("falls back cleanly for malformed or missing registration state", async () => {
    const baseDir = await makeTempDir();
    const statePath = path.join(baseDir, "push", "apns-registrations.json");
    await fs.mkdir(path.dirname(statePath), { recursive: true });
    await fs.writeFile(statePath, "[]", "utf8");

    await expect(loadApnsRegistration("ios-node-missing", baseDir)).resolves.toBeNull();
    await expect(loadApnsRegistration("   ", baseDir)).resolves.toBeNull();
    await expect(clearApnsRegistration("   ", baseDir)).resolves.toBe(false);
    await expect(clearApnsRegistration("ios-node-missing", baseDir)).resolves.toBe(false);
  });

  it("rejects invalid direct and relay registration inputs", async () => {
    const baseDir = await makeTempDir();
    const oversized = "x".repeat(257);

    await expect(
      registerApnsToken({
        nodeId: "ios-node-1",
        token: "not-a-token",
        topic: "ai.openclaw.ios",
        baseDir,
      }),
    ).rejects.toThrow("invalid APNs token");
    await expect(
      registerApnsToken({
        nodeId: "n".repeat(257),
        token: "ABCD1234ABCD1234ABCD1234ABCD1234",
        topic: "ai.openclaw.ios",
        baseDir,
      }),
    ).rejects.toThrow("nodeId required");
    await expect(
      registerApnsToken({
        nodeId: "ios-node-1",
        token: "A".repeat(513),
        topic: "ai.openclaw.ios",
        baseDir,
      }),
    ).rejects.toThrow("invalid APNs token");
    await expect(
      registerApnsToken({
        nodeId: "ios-node-1",
        token: "ABCD1234ABCD1234ABCD1234ABCD1234",
        topic: "a".repeat(256),
        baseDir,
      }),
    ).rejects.toThrow("topic required");
    await expect(
      registerApnsRegistration({
        nodeId: "ios-node-relay",
        transport: "relay",
        relayHandle: "relay-handle-123",
        sendGrant: "send-grant-123",
        installationId: "install-123",
        topic: "ai.openclaw.ios",
        environment: "staging",
        distribution: "official",
        baseDir,
      }),
    ).rejects.toThrow("relay registrations must use production environment");
    await expect(
      registerApnsRegistration({
        nodeId: "ios-node-relay",
        transport: "relay",
        relayHandle: "relay-handle-123",
        sendGrant: "send-grant-123",
        installationId: "install-123",
        topic: "ai.openclaw.ios",
        environment: "production",
        distribution: "beta",
        baseDir,
      }),
    ).rejects.toThrow("relay registrations must use official distribution");
    await expect(
      registerApnsRegistration({
        nodeId: "ios-node-relay",
        transport: "relay",
        relayHandle: oversized,
        sendGrant: "send-grant-123",
        installationId: "install-123",
        topic: "ai.openclaw.ios",
        environment: "production",
        distribution: "official",
        baseDir,
      }),
    ).rejects.toThrow("relayHandle too long");
    await expect(
      registerApnsRegistration({
        nodeId: "ios-node-relay",
        transport: "relay",
        relayHandle: "relay-handle-123",
        sendGrant: "send-grant-123",
        installationId: oversized,
        topic: "ai.openclaw.ios",
        environment: "production",
        distribution: "official",
        baseDir,
      }),
    ).rejects.toThrow("installationId too long");
    await expect(
      registerApnsRegistration({
        nodeId: "ios-node-relay",
        transport: "relay",
        relayHandle: "relay-handle-123",
        sendGrant: "x".repeat(1025),
        installationId: "install-123",
        topic: "ai.openclaw.ios",
        environment: "production",
        distribution: "official",
        baseDir,
      }),
    ).rejects.toThrow("sendGrant too long");
  });

  it("persists with a trailing newline and clears registrations", async () => {
    const baseDir = await makeTempDir();
    await registerApnsToken({
      nodeId: "ios-node-1",
      token: "ABCD1234ABCD1234ABCD1234ABCD1234",
      topic: "ai.openclaw.ios",
      baseDir,
    });

    const statePath = path.join(baseDir, "push", "apns-registrations.json");
    await expect(fs.readFile(statePath, "utf8")).resolves.toMatch(/\n$/);
    await expect(clearApnsRegistration("ios-node-1", baseDir)).resolves.toBe(true);
    await expect(loadApnsRegistration("ios-node-1", baseDir)).resolves.toBeNull();
  });

  it("only clears a registration when the stored entry still matches", async () => {
    vi.useFakeTimers();
    try {
      const baseDir = await makeTempDir();
      vi.setSystemTime(new Date("2026-03-11T00:00:00Z"));
      const stale = await registerApnsToken({
        nodeId: "ios-node-1",
        token: "ABCD1234ABCD1234ABCD1234ABCD1234",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      vi.setSystemTime(new Date("2026-03-11T00:00:01Z"));
      const fresh = await registerApnsToken({
        nodeId: "ios-node-1",
        token: "ABCD1234ABCD1234ABCD1234ABCD1234",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      await expect(
        clearApnsRegistrationIfCurrent({
          nodeId: "ios-node-1",
          registration: stale,
          baseDir,
        }),
      ).resolves.toBe(false);
      await expect(loadApnsRegistration("ios-node-1", baseDir)).resolves.toEqual(fresh);
    } finally {
      vi.useRealTimers();
    }
  });
});

// ---------------------------------------------------------------------------
// Multi-device registration tests (Step 3)
// ---------------------------------------------------------------------------
describe("multi-device APNs registration", () => {
  it("stores up to 3 different tokens for the same nodeId", async () => {
    vi.useFakeTimers();
    try {
      const baseDir = await makeTempDir();
      const nodeId = "ios-multi-1";

      vi.setSystemTime(new Date("2026-04-01T00:00:00Z"));
      await registerApnsToken({
        nodeId,
        token: "AAAA1111AAAA1111AAAA1111AAAA1111",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      vi.setSystemTime(new Date("2026-04-01T00:01:00Z"));
      await registerApnsToken({
        nodeId,
        token: "BBBB2222BBBB2222BBBB2222BBBB2222",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      vi.setSystemTime(new Date("2026-04-01T00:02:00Z"));
      await registerApnsToken({
        nodeId,
        token: "CCCC3333CCCC3333CCCC3333CCCC3333",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      const all = await loadApnsRegistrations(nodeId, baseDir);
      expect(all).toHaveLength(3);
      const tokens = all.map((r) => (r.transport === "direct" ? r.token : null));
      expect(tokens).toContain("aaaa1111aaaa1111aaaa1111aaaa1111");
      expect(tokens).toContain("bbbb2222bbbb2222bbbb2222bbbb2222");
      expect(tokens).toContain("cccc3333cccc3333cccc3333cccc3333");
    } finally {
      vi.useRealTimers();
    }
  });

  it("evicts the oldest registration when a 4th token is registered", async () => {
    vi.useFakeTimers();
    try {
      const baseDir = await makeTempDir();
      const nodeId = "ios-multi-evict";

      vi.setSystemTime(new Date("2026-04-01T00:00:00Z"));
      await registerApnsToken({
        nodeId,
        token: "AAAA1111AAAA1111AAAA1111AAAA1111",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      vi.setSystemTime(new Date("2026-04-01T00:01:00Z"));
      await registerApnsToken({
        nodeId,
        token: "BBBB2222BBBB2222BBBB2222BBBB2222",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      vi.setSystemTime(new Date("2026-04-01T00:02:00Z"));
      await registerApnsToken({
        nodeId,
        token: "CCCC3333CCCC3333CCCC3333CCCC3333",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      vi.setSystemTime(new Date("2026-04-01T00:03:00Z"));
      await registerApnsToken({
        nodeId,
        token: "DDDD4444DDDD4444DDDD4444DDDD4444",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      const all = await loadApnsRegistrations(nodeId, baseDir);
      expect(all).toHaveLength(3);
      const tokens = all.map((r) => (r.transport === "direct" ? r.token : null));
      // Oldest (AAAA) should be evicted
      expect(tokens).not.toContain("aaaa1111aaaa1111aaaa1111aaaa1111");
      expect(tokens).toContain("bbbb2222bbbb2222bbbb2222bbbb2222");
      expect(tokens).toContain("cccc3333cccc3333cccc3333cccc3333");
      expect(tokens).toContain("dddd4444dddd4444dddd4444dddd4444");
    } finally {
      vi.useRealTimers();
    }
  });

  it("deduplicates by direct dedupe key and updates in place", async () => {
    vi.useFakeTimers();
    try {
      const baseDir = await makeTempDir();
      const nodeId = "ios-multi-dedup";

      vi.setSystemTime(new Date("2026-04-01T00:00:00Z"));
      await registerApnsToken({
        nodeId,
        token: "AAAA1111AAAA1111AAAA1111AAAA1111",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      vi.setSystemTime(new Date("2026-04-01T00:01:00Z"));
      await registerApnsToken({
        nodeId,
        token: "BBBB2222BBBB2222BBBB2222BBBB2222",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      // Re-register the same token (same dedupe key: direct:ai.openclaw.ios:sandbox:aaaa...)
      vi.setSystemTime(new Date("2026-04-01T00:05:00Z"));
      await registerApnsToken({
        nodeId,
        token: "AAAA1111AAAA1111AAAA1111AAAA1111",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      const all = await loadApnsRegistrations(nodeId, baseDir);
      // Should still be 2, not 3 — the first was updated in place
      expect(all).toHaveLength(2);

      // The re-registered entry should have the newer timestamp
      const reregistered = all.find(
        (r) => r.transport === "direct" && r.token === "aaaa1111aaaa1111aaaa1111aaaa1111",
      );
      expect(reregistered).toBeDefined();
      expect(reregistered!.updatedAtMs).toBe(new Date("2026-04-01T00:05:00Z").getTime());
    } finally {
      vi.useRealTimers();
    }
  });

  it("uses direct:{topic}:{environment}:{token} as dedupe key for direct registrations", async () => {
    vi.useFakeTimers();
    try {
      const baseDir = await makeTempDir();
      const nodeId = "ios-dedup-key-direct";

      vi.setSystemTime(new Date("2026-04-01T00:00:00Z"));
      // Same token but different environment → different dedupe key → separate entries
      await registerApnsToken({
        nodeId,
        token: "AAAA1111AAAA1111AAAA1111AAAA1111",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      vi.setSystemTime(new Date("2026-04-01T00:01:00Z"));
      await registerApnsToken({
        nodeId,
        token: "AAAA1111AAAA1111AAAA1111AAAA1111",
        topic: "ai.openclaw.ios",
        environment: "production",
        baseDir,
      });

      const all = await loadApnsRegistrations(nodeId, baseDir);
      // Different environments = different dedupe keys = 2 entries
      expect(all).toHaveLength(2);
    } finally {
      vi.useRealTimers();
    }
  });

  it("uses relay:{installationId}:{topic}:{environment} as dedupe key for relay registrations", async () => {
    vi.useFakeTimers();
    try {
      const baseDir = await makeTempDir();
      const nodeId = "ios-dedup-key-relay";

      vi.setSystemTime(new Date("2026-04-01T00:00:00Z"));
      await registerApnsRegistration({
        nodeId,
        transport: "relay",
        relayHandle: "relay-handle-1",
        sendGrant: "send-grant-1",
        installationId: "install-A",
        topic: "ai.openclaw.ios",
        environment: "production",
        distribution: "official",
        baseDir,
      });

      vi.setSystemTime(new Date("2026-04-01T00:01:00Z"));
      await registerApnsRegistration({
        nodeId,
        transport: "relay",
        relayHandle: "relay-handle-2",
        sendGrant: "send-grant-2",
        installationId: "install-B",
        topic: "ai.openclaw.ios",
        environment: "production",
        distribution: "official",
        baseDir,
      });

      const all = await loadApnsRegistrations(nodeId, baseDir);
      // Different installationIds = different dedupe keys = 2 entries
      expect(all).toHaveLength(2);

      // Re-register install-A with different handle → updates in place (same dedupe key)
      vi.setSystemTime(new Date("2026-04-01T00:02:00Z"));
      await registerApnsRegistration({
        nodeId,
        transport: "relay",
        relayHandle: "relay-handle-1-updated",
        sendGrant: "send-grant-1-updated",
        installationId: "install-A",
        topic: "ai.openclaw.ios",
        environment: "production",
        distribution: "official",
        baseDir,
      });

      const allAfter = await loadApnsRegistrations(nodeId, baseDir);
      expect(allAfter).toHaveLength(2);
      const updatedEntry = allAfter.find(
        (r) => r.transport === "relay" && r.installationId === "install-A",
      );
      expect(updatedEntry).toBeDefined();
      expect(updatedEntry!.transport === "relay" && updatedEntry!.relayHandle).toBe(
        "relay-handle-1-updated",
      );
    } finally {
      vi.useRealTimers();
    }
  });

  it("breaks eviction ties by dedupe key sort order", async () => {
    vi.useFakeTimers();
    try {
      const baseDir = await makeTempDir();
      const nodeId = "ios-evict-tiebreak";

      // Register 3 tokens all at the exact same time
      const sameTime = new Date("2026-04-01T00:00:00Z");
      vi.setSystemTime(sameTime);

      await registerApnsToken({
        nodeId,
        token: "CCCC3333CCCC3333CCCC3333CCCC3333",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });
      await registerApnsToken({
        nodeId,
        token: "AAAA1111AAAA1111AAAA1111AAAA1111",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });
      await registerApnsToken({
        nodeId,
        token: "BBBB2222BBBB2222BBBB2222BBBB2222",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      // Now add a 4th — one of the tied entries should be evicted deterministically
      vi.setSystemTime(new Date("2026-04-01T00:01:00Z"));
      await registerApnsToken({
        nodeId,
        token: "DDDD4444DDDD4444DDDD4444DDDD4444",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      const all = await loadApnsRegistrations(nodeId, baseDir);
      expect(all).toHaveLength(3);
      const tokens = all
        .map((r) => (r.transport === "direct" ? r.token : null))
        .toSorted((a, b) => String(a).localeCompare(String(b)));

      // The entry with the lowest dedupe key sort should be evicted when timestamps tie.
      // Dedupe keys: direct:ai.openclaw.ios:sandbox:aaaa... < direct:...:bbbb... < direct:...:cccc...
      // So aaaa entry is evicted (first in sort = oldest tiebreaker)
      expect(tokens).not.toContain("aaaa1111aaaa1111aaaa1111aaaa1111");
      expect(tokens).toContain("dddd4444dddd4444dddd4444dddd4444");
    } finally {
      vi.useRealTimers();
    }
  });

  it("loadApnsRegistration (singular) returns the first/primary registration for backwards compatibility", async () => {
    vi.useFakeTimers();
    try {
      const baseDir = await makeTempDir();
      const nodeId = "ios-compat-singular";

      vi.setSystemTime(new Date("2026-04-01T00:00:00Z"));
      await registerApnsToken({
        nodeId,
        token: "AAAA1111AAAA1111AAAA1111AAAA1111",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      vi.setSystemTime(new Date("2026-04-01T00:01:00Z"));
      await registerApnsToken({
        nodeId,
        token: "BBBB2222BBBB2222BBBB2222BBBB2222",
        topic: "ai.openclaw.ios",
        environment: "sandbox",
        baseDir,
      });

      // Singular should still return a single registration (not null)
      const single = await loadApnsRegistration(nodeId, baseDir);
      expect(single).not.toBeNull();
      expect(single).toHaveProperty("nodeId", nodeId);

      // Plural should return the full array
      const all = await loadApnsRegistrations(nodeId, baseDir);
      expect(all).toHaveLength(2);
    } finally {
      vi.useRealTimers();
    }
  });

  it("loadApnsRegistrations returns full array of registrations", async () => {
    const baseDir = await makeTempDir();
    const nodeId = "ios-load-plural";

    await registerApnsToken({
      nodeId,
      token: "AAAA1111AAAA1111AAAA1111AAAA1111",
      topic: "ai.openclaw.ios",
      environment: "sandbox",
      baseDir,
    });

    const all = await loadApnsRegistrations(nodeId, baseDir);
    expect(Array.isArray(all)).toBe(true);
    expect(all).toHaveLength(1);
    expect(all[0]).toMatchObject({
      nodeId,
      transport: "direct",
      token: "aaaa1111aaaa1111aaaa1111aaaa1111",
    });
  });

  it("loadApnsRegistrations returns empty array for unknown nodeId", async () => {
    const baseDir = await makeTempDir();
    const all = await loadApnsRegistrations("nonexistent-node", baseDir);
    expect(Array.isArray(all)).toBe(true);
    expect(all).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// Legacy migration tests (Step 3)
// ---------------------------------------------------------------------------
describe("legacy migration to multi-device format", () => {
  it("normalizes old single-object format to an array of 1 on load", async () => {
    const baseDir = await makeTempDir();
    const statePath = path.join(baseDir, "push", "apns-registrations.json");
    await fs.mkdir(path.dirname(statePath), { recursive: true });

    // Write the OLD format: registrationsByNodeId maps to single objects
    await fs.writeFile(
      statePath,
      `${JSON.stringify(
        {
          registrationsByNodeId: {
            "ios-legacy-node": {
              nodeId: "ios-legacy-node",
              transport: "direct",
              token: "abcd1234abcd1234abcd1234abcd1234",
              topic: "ai.openclaw.ios",
              environment: "sandbox",
              updatedAtMs: 1000,
            },
          },
        },
        null,
        2,
      )}\n`,
      "utf8",
    );

    const all = await loadApnsRegistrations("ios-legacy-node", baseDir);
    expect(Array.isArray(all)).toBe(true);
    expect(all).toHaveLength(1);
    expect(all[0]).toMatchObject({
      nodeId: "ios-legacy-node",
      transport: "direct",
      token: "abcd1234abcd1234abcd1234abcd1234",
    });
  });

  it("re-save after migration load preserves array format", async () => {
    const baseDir = await makeTempDir();
    const statePath = path.join(baseDir, "push", "apns-registrations.json");
    await fs.mkdir(path.dirname(statePath), { recursive: true });

    // Write old single-object format
    await fs.writeFile(
      statePath,
      `${JSON.stringify(
        {
          registrationsByNodeId: {
            "ios-migrate-node": {
              nodeId: "ios-migrate-node",
              transport: "direct",
              token: "abcd1234abcd1234abcd1234abcd1234",
              topic: "ai.openclaw.ios",
              environment: "sandbox",
              updatedAtMs: 1000,
            },
          },
        },
        null,
        2,
      )}\n`,
      "utf8",
    );

    // Trigger a load + re-save by registering a new token for a different node
    await registerApnsToken({
      nodeId: "ios-other-node",
      token: "EEEE5555EEEE5555EEEE5555EEEE5555",
      topic: "ai.openclaw.ios",
      baseDir,
    });

    // Read the raw file and verify the migrated node is now stored as an array
    const rawJson = JSON.parse(await fs.readFile(statePath, "utf8"));
    const migratedEntry = rawJson.registrationsByNodeId["ios-migrate-node"];
    expect(Array.isArray(migratedEntry)).toBe(true);
  });

  it("normalizes mixed file with some arrays and some single objects", async () => {
    const baseDir = await makeTempDir();
    const statePath = path.join(baseDir, "push", "apns-registrations.json");
    await fs.mkdir(path.dirname(statePath), { recursive: true });

    // Write a mixed format: one node has array, another has single object
    await fs.writeFile(
      statePath,
      `${JSON.stringify(
        {
          registrationsByNodeId: {
            "ios-array-node": [
              {
                nodeId: "ios-array-node",
                transport: "direct",
                token: "aaaa1111aaaa1111aaaa1111aaaa1111",
                topic: "ai.openclaw.ios",
                environment: "sandbox",
                updatedAtMs: 1000,
              },
            ],
            "ios-object-node": {
              nodeId: "ios-object-node",
              transport: "direct",
              token: "bbbb2222bbbb2222bbbb2222bbbb2222",
              topic: "ai.openclaw.ios",
              environment: "sandbox",
              updatedAtMs: 2000,
            },
          },
        },
        null,
        2,
      )}\n`,
      "utf8",
    );

    // Both should be loadable as arrays
    const arrayNodeRegs = await loadApnsRegistrations("ios-array-node", baseDir);
    expect(Array.isArray(arrayNodeRegs)).toBe(true);
    expect(arrayNodeRegs).toHaveLength(1);

    const objectNodeRegs = await loadApnsRegistrations("ios-object-node", baseDir);
    expect(Array.isArray(objectNodeRegs)).toBe(true);
    expect(objectNodeRegs).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// Activity-based expiry tests (Step 4)
// ---------------------------------------------------------------------------
describe("activity-based expiry", () => {
  const DAY_MS = 24 * 60 * 60 * 1000;

  it("prunes registrations older than 90 days (hard max age)", async () => {
    vi.useFakeTimers();
    try {
      const baseDir = await makeTempDir();
      const statePath = path.join(baseDir, "push", "apns-registrations.json");
      await fs.mkdir(path.dirname(statePath), { recursive: true });

      const now = new Date("2026-04-01T00:00:00Z").getTime();

      // Write a registration with updatedAtMs 91 days ago
      await fs.writeFile(
        statePath,
        `${JSON.stringify(
          {
            registrationsByNodeId: {
              "ios-expired": [
                {
                  nodeId: "ios-expired",
                  transport: "direct",
                  token: "aaaa1111aaaa1111aaaa1111aaaa1111",
                  topic: "ai.openclaw.ios",
                  environment: "sandbox",
                  updatedAtMs: now - 91 * DAY_MS,
                  registeredAtMs: now - 91 * DAY_MS,
                },
              ],
            },
          },
          null,
          2,
        )}\n`,
        "utf8",
      );

      vi.setSystemTime(now);
      const all = await loadApnsRegistrations("ios-expired", baseDir);
      expect(all).toHaveLength(0);
    } finally {
      vi.useRealTimers();
    }
  });

  it("keeps registrations younger than 90 days", async () => {
    vi.useFakeTimers();
    try {
      const baseDir = await makeTempDir();
      const statePath = path.join(baseDir, "push", "apns-registrations.json");
      await fs.mkdir(path.dirname(statePath), { recursive: true });

      const now = new Date("2026-04-01T00:00:00Z").getTime();

      await fs.writeFile(
        statePath,
        `${JSON.stringify(
          {
            registrationsByNodeId: {
              "ios-fresh": [
                {
                  nodeId: "ios-fresh",
                  transport: "direct",
                  token: "aaaa1111aaaa1111aaaa1111aaaa1111",
                  topic: "ai.openclaw.ios",
                  environment: "sandbox",
                  updatedAtMs: now - 89 * DAY_MS,
                  registeredAtMs: now - 89 * DAY_MS,
                },
              ],
            },
          },
          null,
          2,
        )}\n`,
        "utf8",
      );

      vi.setSystemTime(now);
      const all = await loadApnsRegistrations("ios-fresh", baseDir);
      expect(all).toHaveLength(1);
    } finally {
      vi.useRealTimers();
    }
  });

  it("prunes registrations with lastConfirmedAtMs older than 30 days (soft stale)", async () => {
    vi.useFakeTimers();
    try {
      const baseDir = await makeTempDir();
      const statePath = path.join(baseDir, "push", "apns-registrations.json");
      await fs.mkdir(path.dirname(statePath), { recursive: true });

      const now = new Date("2026-04-01T00:00:00Z").getTime();

      await fs.writeFile(
        statePath,
        `${JSON.stringify(
          {
            registrationsByNodeId: {
              "ios-stale": [
                {
                  nodeId: "ios-stale",
                  transport: "direct",
                  token: "aaaa1111aaaa1111aaaa1111aaaa1111",
                  topic: "ai.openclaw.ios",
                  environment: "sandbox",
                  updatedAtMs: now - 50 * DAY_MS,
                  registeredAtMs: now - 50 * DAY_MS,
                  lastConfirmedAtMs: now - 31 * DAY_MS,
                },
              ],
            },
          },
          null,
          2,
        )}\n`,
        "utf8",
      );

      vi.setSystemTime(now);
      const all = await loadApnsRegistrations("ios-stale", baseDir);
      expect(all).toHaveLength(0);
    } finally {
      vi.useRealTimers();
    }
  });

  it("keeps registrations with lastConfirmedAtMs within 30 days", async () => {
    vi.useFakeTimers();
    try {
      const baseDir = await makeTempDir();
      const statePath = path.join(baseDir, "push", "apns-registrations.json");
      await fs.mkdir(path.dirname(statePath), { recursive: true });

      const now = new Date("2026-04-01T00:00:00Z").getTime();

      await fs.writeFile(
        statePath,
        `${JSON.stringify(
          {
            registrationsByNodeId: {
              "ios-recent-confirm": [
                {
                  nodeId: "ios-recent-confirm",
                  transport: "direct",
                  token: "aaaa1111aaaa1111aaaa1111aaaa1111",
                  topic: "ai.openclaw.ios",
                  environment: "sandbox",
                  updatedAtMs: now - 50 * DAY_MS,
                  registeredAtMs: now - 50 * DAY_MS,
                  lastConfirmedAtMs: now - 29 * DAY_MS,
                },
              ],
            },
          },
          null,
          2,
        )}\n`,
        "utf8",
      );

      vi.setSystemTime(now);
      const all = await loadApnsRegistrations("ios-recent-confirm", baseDir);
      expect(all).toHaveLength(1);
    } finally {
      vi.useRealTimers();
    }
  });

  it("keeps old registration if lastConfirmedAtMs is recent (activity resets soft expiry)", async () => {
    vi.useFakeTimers();
    try {
      const baseDir = await makeTempDir();
      const statePath = path.join(baseDir, "push", "apns-registrations.json");
      await fs.mkdir(path.dirname(statePath), { recursive: true });

      const now = new Date("2026-04-01T00:00:00Z").getTime();

      // updatedAtMs is 80 days old (within 90-day hard limit)
      // lastConfirmedAtMs is 5 days ago (within 30-day soft limit)
      await fs.writeFile(
        statePath,
        `${JSON.stringify(
          {
            registrationsByNodeId: {
              "ios-active-old": [
                {
                  nodeId: "ios-active-old",
                  transport: "direct",
                  token: "aaaa1111aaaa1111aaaa1111aaaa1111",
                  topic: "ai.openclaw.ios",
                  environment: "sandbox",
                  updatedAtMs: now - 80 * DAY_MS,
                  registeredAtMs: now - 80 * DAY_MS,
                  lastConfirmedAtMs: now - 5 * DAY_MS,
                },
              ],
            },
          },
          null,
          2,
        )}\n`,
        "utf8",
      );

      vi.setSystemTime(now);
      const all = await loadApnsRegistrations("ios-active-old", baseDir);
      expect(all).toHaveLength(1);
    } finally {
      vi.useRealTimers();
    }
  });

  it("prunes lazily on loadRegistrationsState, no background timer", async () => {
    vi.useFakeTimers();
    try {
      const baseDir = await makeTempDir();
      const statePath = path.join(baseDir, "push", "apns-registrations.json");
      await fs.mkdir(path.dirname(statePath), { recursive: true });

      const now = new Date("2026-04-01T00:00:00Z").getTime();

      // Write an expired entry
      await fs.writeFile(
        statePath,
        `${JSON.stringify(
          {
            registrationsByNodeId: {
              "ios-lazy-prune": [
                {
                  nodeId: "ios-lazy-prune",
                  transport: "direct",
                  token: "aaaa1111aaaa1111aaaa1111aaaa1111",
                  topic: "ai.openclaw.ios",
                  environment: "sandbox",
                  updatedAtMs: now - 91 * DAY_MS,
                  registeredAtMs: now - 91 * DAY_MS,
                },
                {
                  nodeId: "ios-lazy-prune",
                  transport: "direct",
                  token: "bbbb2222bbbb2222bbbb2222bbbb2222",
                  topic: "ai.openclaw.ios",
                  environment: "sandbox",
                  updatedAtMs: now - 10 * DAY_MS,
                  registeredAtMs: now - 10 * DAY_MS,
                },
              ],
            },
          },
          null,
          2,
        )}\n`,
        "utf8",
      );

      vi.setSystemTime(now);

      // Before load: raw file still has both entries
      const rawBefore = JSON.parse(await fs.readFile(statePath, "utf8"));
      expect(rawBefore.registrationsByNodeId["ios-lazy-prune"]).toHaveLength(2);

      // Load triggers pruning
      const all = await loadApnsRegistrations("ios-lazy-prune", baseDir);
      expect(all).toHaveLength(1);
      expect(all[0]).toMatchObject({
        token: "bbbb2222bbbb2222bbbb2222bbbb2222",
      });
    } finally {
      vi.useRealTimers();
    }
  });

  it("prunes per-entry within the array, not the whole node's registrations", async () => {
    vi.useFakeTimers();
    try {
      const baseDir = await makeTempDir();
      const statePath = path.join(baseDir, "push", "apns-registrations.json");
      await fs.mkdir(path.dirname(statePath), { recursive: true });

      const now = new Date("2026-04-01T00:00:00Z").getTime();

      await fs.writeFile(
        statePath,
        `${JSON.stringify(
          {
            registrationsByNodeId: {
              "ios-partial-prune": [
                {
                  nodeId: "ios-partial-prune",
                  transport: "direct",
                  token: "aaaa1111aaaa1111aaaa1111aaaa1111",
                  topic: "ai.openclaw.ios",
                  environment: "sandbox",
                  updatedAtMs: now - 91 * DAY_MS,
                  registeredAtMs: now - 91 * DAY_MS,
                },
                {
                  nodeId: "ios-partial-prune",
                  transport: "direct",
                  token: "bbbb2222bbbb2222bbbb2222bbbb2222",
                  topic: "ai.openclaw.ios",
                  environment: "sandbox",
                  updatedAtMs: now - 5 * DAY_MS,
                  registeredAtMs: now - 5 * DAY_MS,
                },
                {
                  nodeId: "ios-partial-prune",
                  transport: "direct",
                  token: "cccc3333cccc3333cccc3333cccc3333",
                  topic: "ai.openclaw.ios",
                  environment: "sandbox",
                  updatedAtMs: now - 91 * DAY_MS,
                  registeredAtMs: now - 91 * DAY_MS,
                },
              ],
            },
          },
          null,
          2,
        )}\n`,
        "utf8",
      );

      vi.setSystemTime(now);
      const all = await loadApnsRegistrations("ios-partial-prune", baseDir);
      // Only the non-expired entry should survive
      expect(all).toHaveLength(1);
      expect(all[0]).toMatchObject({
        token: "bbbb2222bbbb2222bbbb2222bbbb2222",
      });
    } finally {
      vi.useRealTimers();
    }
  });
});

// ---------------------------------------------------------------------------
// Concurrency tests (Step 3)
// ---------------------------------------------------------------------------
describe("concurrent APNs registration operations", () => {
  it("handles 10 parallel registerApnsRegistration calls without data corruption", async () => {
    const baseDir = await makeTempDir();
    const nodeId = "ios-concurrent";
    const tokens = Array.from({ length: 10 }, (_, i) => {
      const hex = (i + 1).toString(16).padStart(2, "0");
      return `${hex}${hex}${hex}${hex}${hex}${hex}${hex}${hex}${hex}${hex}${hex}${hex}${hex}${hex}${hex}${hex}`;
    });

    // Fire all 10 in parallel
    await Promise.all(
      tokens.map((token) =>
        registerApnsToken({
          nodeId,
          token,
          topic: "ai.openclaw.ios",
          environment: "sandbox",
          baseDir,
        }),
      ),
    );

    const all = await loadApnsRegistrations(nodeId, baseDir);
    // Must not exceed MAX_REGISTRATIONS_PER_NODE (3)
    expect(all.length).toBeLessThanOrEqual(3);
    // Must have at least 1 entry (no data loss to empty)
    expect(all.length).toBeGreaterThanOrEqual(1);
    // All entries should be valid registrations
    for (const reg of all) {
      expect(reg).toHaveProperty("nodeId", nodeId);
      expect(reg).toHaveProperty("transport", "direct");
    }
  });

  it("handles parallel register + clear without crash", async () => {
    const baseDir = await makeTempDir();
    const nodeId = "ios-concurrent-clear";

    // Seed an initial registration
    await registerApnsToken({
      nodeId,
      token: "AAAA1111AAAA1111AAAA1111AAAA1111",
      topic: "ai.openclaw.ios",
      environment: "sandbox",
      baseDir,
    });

    // Fire register and clear in parallel — should not crash
    await expect(
      Promise.all([
        registerApnsToken({
          nodeId,
          token: "BBBB2222BBBB2222BBBB2222BBBB2222",
          topic: "ai.openclaw.ios",
          environment: "sandbox",
          baseDir,
        }),
        clearApnsRegistration(nodeId, baseDir),
        registerApnsToken({
          nodeId,
          token: "CCCC3333CCCC3333CCCC3333CCCC3333",
          topic: "ai.openclaw.ios",
          environment: "sandbox",
          baseDir,
        }),
      ]),
    ).resolves.toBeDefined();

    // After all operations, state should be consistent (not corrupted)
    const all = await loadApnsRegistrations(nodeId, baseDir);
    expect(Array.isArray(all)).toBe(true);
    expect(all.length).toBeLessThanOrEqual(3);
  });
});
