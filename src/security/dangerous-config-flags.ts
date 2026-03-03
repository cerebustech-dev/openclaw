import type { OpenClawConfig } from "../config/config.js";

export function collectEnabledInsecureOrDangerousFlags(cfg: OpenClawConfig): string[] {
  const enabledFlags: string[] = [];
  if (cfg.gateway?.auth?.mode === "none") {
    enabledFlags.push("gateway.auth.mode=none");
  }
  if (cfg.gateway?.controlUi?.allowInsecureAuth === true) {
    enabledFlags.push("gateway.controlUi.allowInsecureAuth=true");
  }
  if (cfg.gateway?.controlUi?.dangerouslyAllowHostHeaderOriginFallback === true) {
    enabledFlags.push("gateway.controlUi.dangerouslyAllowHostHeaderOriginFallback=true");
  }
  if (cfg.gateway?.controlUi?.dangerouslyDisableDeviceAuth === true) {
    enabledFlags.push("gateway.controlUi.dangerouslyDisableDeviceAuth=true");
  }
  if (cfg.hooks?.gmail?.dangerouslyAllowUnsafeExternalContent === true) {
    enabledFlags.push("hooks.gmail.dangerouslyAllowUnsafeExternalContent=true");
  }
  if (Array.isArray(cfg.hooks?.mappings)) {
    for (const [index, mapping] of cfg.hooks.mappings.entries()) {
      if (mapping?.dangerouslyAllowUnsafeExternalContent === true) {
        enabledFlags.push(`hooks.mappings[${index}].dangerouslyAllowUnsafeExternalContent=true`);
      }
    }
  }
  if (cfg.tools?.exec?.applyPatch?.workspaceOnly === false) {
    enabledFlags.push("tools.exec.applyPatch.workspaceOnly=false");
  }
  return enabledFlags;
}
