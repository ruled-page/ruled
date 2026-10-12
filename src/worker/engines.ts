import type { EngineKind } from '@sitepen/formula-engine';
import type { EngineInfo } from '../core/engineIds';

let loaded: { kind: EngineKind; info: EngineInfo } | undefined;

/**
 * Loads the calculation engine, a chunk of the worker's own fetched when a
 * workbook is first opened, noting how long that took and what was
 * downloaded for it.
 */
export async function loadEngine(): Promise<{ kind: EngineKind; info: EngineInfo }> {
  if (loaded) return loaded;
  const seen = performance.getEntriesByType('resource').length;
  const started = performance.now();
  const kind = await import('@sitepen/formula-engine/native').then((m) => m.nativeEngine);
  const loadMs = Math.round(performance.now() - started);
  const fetched = performance.getEntriesByType('resource').slice(seen) as PerformanceResourceTiming[];
  const total = (key: 'transferSize' | 'decodedBodySize') => fetched.reduce((n, e) => n + (e[key] ?? 0), 0);
  loaded = { kind, info: { id: 'native', name: kind.name, loadMs, transferBytes: total('transferSize'), decodedBytes: total('decodedBodySize') } };
  return loaded;
}

/** The worker's JavaScript heap in use, where the browser tells (Chrome); undefined elsewhere. */
export function heapBytes(): number | undefined {
  return (performance as Performance & { memory?: { usedJSHeapSize: number } }).memory?.usedJSHeapSize;
}
