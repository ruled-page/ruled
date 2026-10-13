import type { StoreOptions } from '../state/store';
import type { EditorHandle } from './handle';

/** Options an open editor takes as they change, without closing its workbook. */
export const LIVE_OPTIONS: readonly (keyof StoreOptions)[] = ['theme', 'language', 'user', 'schemas', 'customLists'];

/** Equal, or arrays and plain objects whose entries are (lists of lists one level deeper), so a new literal of the same values isn't a change. */
function same(a: unknown, b: unknown, depth = 0): boolean {
  if (a === b) return true;
  if (Array.isArray(a) && Array.isArray(b)) return a.length === b.length && a.every((x, i) => (depth < 1 ? same(x, b[i], depth + 1) : x === b[i]));
  if (a && b && typeof a === 'object' && typeof b === 'object' && !Array.isArray(a) && !Array.isArray(b)) {
    const keys = Object.keys(a);
    return keys.length === Object.keys(b).length && keys.every((k) => (a as Record<string, unknown>)[k] === (b as Record<string, unknown>)[k]);
  }
  return false;
}

/** The options that differ between two sets. */
export function changedOptions(before: StoreOptions, next: StoreOptions): (keyof StoreOptions)[] {
  return (Object.keys({ ...before, ...next }) as (keyof StoreOptions)[]).filter((k) => !same(before[k], next[k]));
}

/**
 * Gives an open editor the theme, language, user, schemas and custom lists
 * of `next` where they differ from `before`; schemas no longer listed are
 * removed. Other options aren't touched: they're read when an editor is made.
 */
export function applyLiveOptions(handle: EditorHandle, before: StoreOptions, next: StoreOptions, changed = changedOptions(before, next)): void {
  if (changed.includes('theme')) handle.setTheme(next.theme);
  if (changed.includes('language') && next.language) handle.setLanguage(next.language);
  if (changed.includes('user')) handle.setUser(next.user);
  if (changed.includes('customLists')) handle.setCustomLists(next.customLists);
  if (changed.includes('schemas')) {
    const ids = new Set((next.schemas ?? []).map((s) => s.id));
    for (const s of before.schemas ?? []) if (!ids.has(s.id)) handle.removeSchema(s.id);
    // Registered last to first, so the first is tried first, as in `schemas`.
    for (const s of [...(next.schemas ?? [])].reverse()) handle.registerSchema(s);
  }
}
