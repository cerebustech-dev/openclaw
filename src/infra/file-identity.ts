export type FileIdentityStat = {
  dev: number | bigint;
  ino: number | bigint;
};

export function sameFileIdentity(
  left: FileIdentityStat,
  right: FileIdentityStat,
  platform: NodeJS.Platform = process.platform,
): boolean {
  if (left.ino !== right.ino) {
    return false;
  }

  if (left.dev === right.dev) {
    return true;
  }
  // On Windows, path-based stat (lstat/stat) and fd-based stat (fstat) can
  // report different volume identifiers for the same file: lstat queries the
  // Win32 path-volume API and fstat queries the NT-handle volume API, which
  // return distinct serial-number forms (legacy 32-bit DOS serial vs full NT
  // serial). Either side may also be zero on older Node.js builds. NTFS file
  // index (ino) uniquely identifies a file within a volume, and the open()
  // boundary check is already pinned to a single resolved path, so an
  // ino match on win32 is sufficient identity.
  return platform === "win32";
}
