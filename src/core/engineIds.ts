/**
 * The calculation engines a workbook can be worked out with. Each is loaded
 * into the worker only when a workbook is opened with it, so this list is
 * all the page itself knows of them.
 */
export const ENGINES = [
  { id: 'native', name: 'Native' },
  { id: 'hyperformula', name: 'HyperFormula' },
] as const;

export type EngineId = (typeof ENGINES)[number]['id'];

/** How an engine was loaded: how long it took, and how much was downloaded for it (0 once it was cached). */
export interface EngineInfo {
  id: EngineId;
  name: string;
  loadMs: number;
  /** Bytes over the network, compressed as served. */
  transferBytes: number;
  /** Bytes of code once uncompressed. */
  decodedBytes: number;
}

/** What the worker measured of an open or an edit: how long it took, and its JavaScript heap where the browser says. */
export interface WorkStats {
  ms: number;
  heapBytes?: number;
}
