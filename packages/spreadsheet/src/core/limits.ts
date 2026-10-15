/** Limits applied before a package is decompressed, so a hostile file can't exhaust memory. */
export interface PackageLimits {
  /** Largest file accepted. Default 50 MB. */
  maxFileBytes?: number;
  /** Largest total size once unzipped. Default 500 MB. */
  maxUnzippedBytes?: number;
  /** Highest unzipped:zipped ratio for a large part, the signature of a zip bomb. Default 200. */
  maxRatio?: number;
  /** Most parts in the package. Default 20,000. */
  maxParts?: number;
}

/** Sized for a browser tab, phones included: larger workbooks can be allowed through OpenOptions.limits. */
export const DEFAULT_LIMITS: Required<PackageLimits> = {
  maxFileBytes: 50 * 1024 * 1024,
  maxUnzippedBytes: 500 * 1024 * 1024,
  maxRatio: 200,
  maxParts: 20_000,
};
