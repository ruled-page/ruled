/**
 * The calculation engine, as the page knows it: the engine itself is loaded
 * into the worker only when a workbook is opened.
 */
export type EngineId = 'native';

/** How the engine was loaded: how long it took, and how much was downloaded for it (0 once it was cached). */
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
