import { describe, expect, it } from "vitest";
import type { OpenClawConfig } from "../config/config.js";
import { collectEnabledInsecureOrDangerousFlags } from "./dangerous-config-flags.js";

describe("collectEnabledInsecureOrDangerousFlags", () => {
  it("flags auth.mode=none as dangerous", () => {
    const cfg = { gateway: { auth: { mode: "none" } } } as OpenClawConfig;
    const flags = collectEnabledInsecureOrDangerousFlags(cfg);
    expect(flags).toContain("gateway.auth.mode=none");
  });

  it("does not flag auth.mode=token", () => {
    const cfg = { gateway: { auth: { mode: "token" } } } as OpenClawConfig;
    const flags = collectEnabledInsecureOrDangerousFlags(cfg);
    expect(flags).not.toContain("gateway.auth.mode=none");
  });

  it("does not flag when auth config is absent", () => {
    const cfg = {} as OpenClawConfig;
    const flags = collectEnabledInsecureOrDangerousFlags(cfg);
    expect(flags).not.toContain("gateway.auth.mode=none");
  });
});
