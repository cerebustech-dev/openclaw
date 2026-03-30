import { buildBoundedGlobRegex } from "openclaw/plugin-sdk/security-runtime";

export function matchBrowserUrlPattern(pattern: string, url: string): boolean {
  const trimmedPattern = pattern.trim();
  if (!trimmedPattern) {
    return false;
  }
  if (trimmedPattern === url) {
    return true;
  }
  if (trimmedPattern.includes("*")) {
    const regex = buildBoundedGlobRegex(trimmedPattern);
    if (!regex) return false;
    return regex.test(url);
  }
  return url.includes(trimmedPattern);
}
