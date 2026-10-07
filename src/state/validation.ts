import { inRect } from '../core/refs';
import { msg, type Message } from '../i18n/message';
import { VT, type DisplayStyle, type ListValidation, type SheetInfo } from '../core/types';
import { displayText } from './format';
import type { SheetView } from './sheetView';

/** Most entries a range-based list offers, to keep the dropdown usable. */
const MAX_OPTIONS = 1000;

export function listValidationAt(info: SheetInfo, row: number, col: number): ListValidation | undefined {
  return info.listValidations.find((v) => v.rects.some((r) => inRect(r, row, col)));
}

/**
 * The values a list validation allows, as Excel shows them in its dropdown:
 * the displayed text of the source cells, without blanks or repeats.
 * Undefined when the source can't be resolved.
 */
export function listOptions(dv: ListValidation, views: SheetView[], styles: DisplayStyle[], date1904: boolean): string[] | undefined {
  const src = dv.source;
  if (src.kind === 'values') return src.values;
  if (src.kind !== 'range') return undefined;
  const view = views[src.sheet];
  if (!view) return undefined;
  const r2 = Math.min(src.rect.r2, Math.max(view.info.maxRow, src.rect.r1) + 200);
  const c2 = Math.min(src.rect.c2, Math.max(view.info.maxCol, src.rect.c1) + 20);
  const seen = new Set<string>();
  const out: string[] = [];
  for (let row = src.rect.r1; row <= r2 && out.length < MAX_OPTIONS; row++) {
    for (let col = src.rect.c1; col <= c2 && out.length < MAX_OPTIONS; col++) {
      const rec = view.get(row, col);
      if (!rec || rec.t === VT.None) continue;
      const text = displayText(rec, styles[rec.style] ?? styles[0], date1904).text.trim();
      if (!text || seen.has(text)) continue;
      seen.add(text);
      out.push(text);
    }
  }
  return out;
}

/** Whether typed text is one of the options: ignoring case, and numerically for numbers. */
export function listAllows(options: readonly string[], text: string): boolean {
  const t = text.trim().toLowerCase();
  const n = Number(t.replace(/,/g, ''));
  return options.some((o) => {
    const v = o.trim().toLowerCase();
    if (v === t) return true;
    const m = Number(v.replace(/,/g, ''));
    return t !== '' && v !== '' && !Number.isNaN(n) && !Number.isNaN(m) && n === m;
  });
}

/** Excel's own wording when a value isn't allowed and the rule has no message. */
export const DEFAULT_LIST_ERROR: Message = msg('store.validation.defaultError');
