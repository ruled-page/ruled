import type { Text } from '../i18n/message';
import type { CellUpdate, EditResult } from './types';

/** Texts without repeats; messages are compared by content, as each edit makes its own. */
export function distinctTexts(texts: Text[]): Text[] {
  return [...new Map(texts.map((t) => [JSON.stringify(t), t])).values()];
}

/**
 * Combines results of edits applied one after another into one result that
 * undoes and redoes as a single step. Restoring applies snapshots in order
 * with the last one winning, so the earliest "before" goes last in undo.
 */
export function mergeEditResults(results: EditResult[]): EditResult {
  const updates = new Map<string, CellUpdate>();
  for (const r of results) for (const u of r.updates) updates.set(`${u.sheet}:${u.row}:${u.col}`, u);
  const layout = results.filter((r) => r.layout?.applied.length);
  return {
    ...(layout.length && {
      layout: { applied: layout.flatMap((r) => r.layout!.applied), undo: layout.reverse().flatMap((r) => r.layout!.undo) },
    }),
    updates: [...updates.values()],
    undo: results.slice().reverse().flatMap((r) => r.undo),
    redo: results.flatMap((r) => r.redo),
    skipped: results.reduce((n, r) => n + r.skipped, 0),
    skippedReasons: distinctTexts(results.flatMap((r) => r.skippedReasons)),
    warnings: results.flatMap((r) => r.warnings),
    newStyles: results.flatMap((r) => r.newStyles),
    editedCount: results.length ? results[results.length - 1].editedCount : 0,
    rejected: results.find((r) => r.rejected)?.rejected,
  };
}
