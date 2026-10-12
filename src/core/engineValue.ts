// Kept apart from the engine so that code on the page can read engine values
// without loading the calculation engine itself, which runs in the worker.

export type EngineValue = number | string | boolean | { error: string; message?: string } | null;

/** The editor's mark for a formula the engine can't calculate, left for Excel to judge: #ERROR!, an error Excel never shows. */
export function isUnsupportedError(v: EngineValue): v is { error: string; message?: string } {
  return !!v && typeof v === 'object' && v.error === '#ERROR!';
}
