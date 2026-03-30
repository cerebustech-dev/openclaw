import { describe, expect, it, vi } from "vitest";

vi.mock("../../../../src/media/fetch.js", () => ({
  fetchRemoteMedia: vi.fn(),
}));

vi.mock("../../../../src/media/store.js", () => ({
  saveMediaBuffer: vi.fn(),
}));

vi.mock("../../../../src/globals.js", () => ({
  logVerbose: () => {},
}));

/**
 * ReDoS safety tests for resolveDiscordMentions.
 *
 * The resolveDiscordMentions function (internal to message-utils.ts) builds
 * a regex from user.id. If user.id contains a regex-special string,
 * this could cause ReDoS or incorrect matches.
 */
describe("resolveDiscordMentions ReDoS safety", () => {
  it("handles non-numeric user.id without hanging or crashing", async () => {
    const { resolveDiscordMessageText } = await import("./message-utils.js");

    const maliciousId = "(a+)+$";
    const start = performance.now();
    const result = resolveDiscordMessageText({
      content: "Hello <@" + maliciousId + "> world",
      mentionedUsers: [
        { id: maliciousId, username: "hacker", globalName: null, discriminator: "0" },
      ],
    } as any);
    const elapsed = performance.now() - start;
    expect(elapsed).toBeLessThan(100);
    // The malicious user.id should be skipped (not resolved), so the original text stays
    expect(result).not.toContain("@hacker");
  });

  it("resolves mentions with a normal numeric user.id", async () => {
    const { resolveDiscordMessageText } = await import("./message-utils.js");

    const result = resolveDiscordMessageText({
      content: "Hello <@123456789> world",
      mentionedUsers: [
        { id: "123456789", username: "alice", globalName: "Alice", discriminator: "0" },
      ],
    } as any);
    expect(result).toBe("Hello @Alice world");
  });
});
