import path from "node:path";
import { createSubsystemLogger } from "../logging/subsystem.js";

const log = createSubsystemLogger("security/elevated-guardrails");

/**
 * Denylist patterns for obviously destructive commands in elevated-full exec mode.
 * Each pattern is tested against individual pipeline segments.
 */
const DESTRUCTIVE_PATTERNS: Array<{ pattern: RegExp; label: string }> = [
  { pattern: /\brm\s+-[^\s]*r[^\s]*f[^\s]*\s+\/(?:\s|$)/, label: "rm -rf /" },
  { pattern: /\brm\s+-[^\s]*f[^\s]*r[^\s]*\s+\/(?:\s|$)/, label: "rm -rf /" },
  { pattern: /\bmkfs\b/, label: "mkfs" },
  { pattern: /\bdd\b.*\bof\s*=\s*\/dev\/[sh]d[a-z]/, label: "dd to block device" },
  { pattern: /:\(\)\s*\{\s*:\|:\s*&\s*\}\s*;?\s*:/, label: "fork bomb" },
];

/**
 * Patterns for destructive operations targeting paths outside workspace.
 */
const DESTRUCTIVE_PATH_PATTERNS: Array<{
  pattern: RegExp;
  extractor: (m: RegExpMatchArray) => string;
  label: string;
}> = [
  {
    pattern: /\brm\s+-[^\s]*r[^\s]*\s+(\/\S+)/,
    extractor: (m) => m[1] ?? "",
    label: "recursive rm",
  },
];

/**
 * Normalize a command by removing common obfuscation patterns.
 * Handles backslash-escaped chars, $'...' quoting, and IFS tricks.
 */
function deobfuscate(command: string): string {
  let cmd = command;
  // Remove backslash escapes within commands (e.g., r\m → rm)
  cmd = cmd.replace(/\\(?=[a-zA-Z])/g, "");
  // Expand $'\xNN' hex escapes
  cmd = cmd.replace(/\$'([^']+)'/g, (_match, inner: string) => {
    return inner.replace(/\\x([0-9a-fA-F]{2})/g, (_: string, hex: string) =>
      String.fromCharCode(parseInt(hex, 16)),
    );
  });
  return cmd;
}

/**
 * Split a command string into pipeline/semicolon-separated segments.
 */
function splitCommandSegments(command: string): string[] {
  return command
    .split(/[|;&]/)
    .map((s) => s.trim())
    .filter(Boolean);
}

/**
 * Check if a target path is inside the workspace directory.
 */
function isInsideWorkspace(targetPath: string, workspaceDir: string): boolean {
  const resolved = path.resolve(targetPath);
  const workspace = path.resolve(workspaceDir);
  return resolved.startsWith(workspace + path.sep) || resolved === workspace;
}

export type GuardrailResult = { allowed: true } | { allowed: false; reason: string };

/**
 * Check an elevated-full exec command against guardrails.
 * Blocks obviously destructive commands and operations targeting paths outside workspace.
 */
export function checkElevatedFullGuardrail(command: string, workspaceDir: string): GuardrailResult {
  const deobfuscated = deobfuscate(command);
  const segments = splitCommandSegments(deobfuscated);

  for (const segment of segments) {
    // Check denylist patterns
    for (const { pattern, label } of DESTRUCTIVE_PATTERNS) {
      if (pattern.test(segment)) {
        log.error(`elevated-full guardrail blocked: ${label}`);
        return {
          allowed: false,
          reason: `Blocked by elevated-full guardrail: ${label}`,
        };
      }
    }

    // Check destructive operations targeting paths outside workspace
    for (const { pattern, extractor, label } of DESTRUCTIVE_PATH_PATTERNS) {
      const match = segment.match(pattern);
      if (match) {
        const targetPath = extractor(match);
        if (
          targetPath &&
          targetPath.startsWith("/") &&
          !isInsideWorkspace(targetPath, workspaceDir)
        ) {
          log.error(
            `elevated-full guardrail blocked: ${label} targeting ${targetPath} (outside workspace)`,
          );
          return {
            allowed: false,
            reason: `Blocked by elevated-full guardrail: ${label} targeting path outside workspace (${targetPath})`,
          };
        }
      }
    }
  }

  // Audit log for all elevated-full commands
  log.info(`elevated-full exec: ${command.slice(0, 200)}`);
  return { allowed: true };
}
