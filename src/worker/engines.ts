import type { EngineKind } from '../core/engine';
import type { EngineId, EngineInfo } from '../core/engineIds';

/** Whether this build has the editor's own engine; the library builds don't yet, keeping their one worker file small. */
declare const __WBE_NATIVE_ENGINE__: boolean;

/** Each engine, imported when a workbook is first opened with it: its code is a chunk of its own. */
const LOADERS: Partial<Record<EngineId, () => Promise<EngineKind>>> = {
  hyperformula: () => import('../core/calc').then((m) => m.hyperFormulaEngine),
};
if (__WBE_NATIVE_ENGINE__) LOADERS.native = () => import('../engine/native').then((m) => m.nativeEngine);

const loaded = new Map<EngineId, { kind: EngineKind; info: EngineInfo }>();

/**
 * Loads an engine (HyperFormula for one this build doesn't have), noting
 * how long that took and what was downloaded for it the first time.
 */
export async function loadEngine(requested: string | undefined): Promise<{ kind: EngineKind; info: EngineInfo }> {
  const id: EngineId = requested && LOADERS[requested as EngineId] ? (requested as EngineId) : 'hyperformula';
  const known = loaded.get(id);
  if (known) return known;
  const seen = performance.getEntriesByType('resource').length;
  const started = performance.now();
  const kind = await LOADERS[id]!();
  const loadMs = Math.round(performance.now() - started);
  const fetched = performance.getEntriesByType('resource').slice(seen) as PerformanceResourceTiming[];
  const total = (key: 'transferSize' | 'decodedBodySize') => fetched.reduce((n, e) => n + (e[key] ?? 0), 0);
  const result = { kind, info: { id, name: kind.name, loadMs, transferBytes: total('transferSize'), decodedBytes: total('decodedBodySize') } };
  loaded.set(id, result);
  return result;
}

/** The worker's JavaScript heap in use, where the browser tells (Chrome); undefined elsewhere. */
export function heapBytes(): number | undefined {
  return (performance as Performance & { memory?: { usedJSHeapSize: number } }).memory?.usedJSHeapSize;
}
