// TODO(driftlane-bzn): pass-B reconciliation — restore this rate-limit test
// suite once `checkSubagentSpawnRateLimit` is re-introduced into
// subagent-spawn.ts. The original 141-line suite (with addSubagentRunForTests
// + SubagentRunRecord fixtures) is preserved in git history at
// security/upstream-2026.4.27-integration~ before the v2026.4.27 merge.
import { describe, it } from "vitest";

describe.skip("checkSubagentSpawnRateLimit (deferred to driftlane-bzn)", () => {
  it.skip("placeholder", () => {
    // intentionally empty — rate-limit production wiring pending
  });
});
