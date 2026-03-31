const BLOCKED_OBJECT_KEYS = new Set(["__proto__", "prototype", "constructor"]);

export function isBlockedObjectKey(key: string): boolean {
  return BLOCKED_OBJECT_KEYS.has(key);
}

export function assertSafePathSegments(segments: readonly string[]): void {
  for (const s of segments) {
    if (BLOCKED_OBJECT_KEYS.has(s)) {
      throw new Error(`Blocked path segment: "${s}"`);
    }
  }
}

export function assertSafeObjectKey(key: string): void {
  if (BLOCKED_OBJECT_KEYS.has(key)) {
    throw new Error(`Blocked object key: "${key}"`);
  }
}
