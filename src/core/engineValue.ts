// Kept apart from calc.ts so that code on the page can read engine values
// without loading the calculation engine itself, which runs in the worker.

export type EngineValue = number | string | boolean | { error: string; message?: string } | null;

/** Errors the engine reports when it cannot interpret a formula at all. */
export function isUnsupportedError(v: EngineValue): v is { error: string; message?: string } {
  if (!v || typeof v !== 'object') return false;
  if (v.error === '#ERROR!') return true;
  return v.error === '#NAME?' && /not recognized|not supported/i.test(v.message ?? '');
}
