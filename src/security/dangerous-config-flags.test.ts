import { beforeEach, describe, expect, it, vi } from "vitest";
import type { OpenClawConfig } from "../config/config.js";
import { collectEnabledInsecureOrDangerousFlags } from "./dangerous-config-flags.js";

const { loadPluginManifestRegistryMock } = vi.hoisted(() => ({
  loadPluginManifestRegistryMock: vi.fn(),
}));

vi.mock("../plugins/manifest-registry.js", () => ({
  loadPluginManifestRegistry: loadPluginManifestRegistryMock,
}));

function asConfig(value: unknown): OpenClawConfig {
  return value as OpenClawConfig;
}

describe("collectEnabledInsecureOrDangerousFlags", () => {
  beforeEach(() => {
    loadPluginManifestRegistryMock.mockReset();
  });

  it("flags auth.mode=none as dangerous", () => {
    loadPluginManifestRegistryMock.mockReturnValue({ plugins: [], diagnostics: [] });
    const cfg = { gateway: { auth: { mode: "none" } } } as OpenClawConfig;
    const flags = collectEnabledInsecureOrDangerousFlags(cfg);
    expect(flags).toContain("gateway.auth.mode=none");
  });

  it("does not flag auth.mode=token", () => {
    loadPluginManifestRegistryMock.mockReturnValue({ plugins: [], diagnostics: [] });
    const cfg = { gateway: { auth: { mode: "token" } } } as OpenClawConfig;
    const flags = collectEnabledInsecureOrDangerousFlags(cfg);
    expect(flags).not.toContain("gateway.auth.mode=none");
  });

  it("does not flag when auth config is absent", () => {
    loadPluginManifestRegistryMock.mockReturnValue({ plugins: [], diagnostics: [] });
    const cfg = {} as OpenClawConfig;
    const flags = collectEnabledInsecureOrDangerousFlags(cfg);
    expect(flags).not.toContain("gateway.auth.mode=none");
  });

  it("collects manifest-declared dangerous plugin config values", () => {
    loadPluginManifestRegistryMock.mockReturnValue({
      plugins: [
        {
          id: "acpx",
          configContracts: {
            dangerousFlags: [{ path: "permissionMode", equals: "approve-all" }],
          },
        },
      ],
      diagnostics: [],
    });

    expect(
      collectEnabledInsecureOrDangerousFlags(
        asConfig({
          plugins: {
            entries: {
              acpx: {
                config: {
                  permissionMode: "approve-all",
                },
              },
            },
          },
        }),
      ),
    ).toContain("plugins.entries.acpx.config.permissionMode=approve-all");
  });

  it("ignores plugin config values that are not declared as dangerous", () => {
    loadPluginManifestRegistryMock.mockReturnValue({
      plugins: [
        {
          id: "other",
          configContracts: {
            dangerousFlags: [{ path: "mode", equals: "danger" }],
          },
        },
      ],
      diagnostics: [],
    });

    expect(
      collectEnabledInsecureOrDangerousFlags(
        asConfig({
          plugins: {
            entries: {
              other: {
                config: {
                  mode: "safe",
                },
              },
            },
          },
        }),
      ),
    ).toEqual([]);
  });
});
