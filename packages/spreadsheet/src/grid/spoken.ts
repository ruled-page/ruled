import { rangeToA1, toA1 } from '@ruled/formula/refs';
import { linkAddress } from '../core/links';
import { CellFlag, type Rect } from '../core/types';
import { effectiveStyle } from '../state/cellStyle';
import { displayText } from '../state/format';
import { cellLink } from '../state/links';
import type { SheetView } from '../state/sheetView';
import type { WorkbookStore } from '../state/store';
import { noteSummary } from '../ui/FormulaBar';

/**
 * What a screen reader says for a cell: its reference, what it shows, and
 * then what the grid only shows by colour or marks: a merge, a formula and
 * whether its value is out of date or Excel's, a change, a list, a note
 * or comments, a link and where it goes, and any problem with it. `shown`
 * is its text as drawn; `list` says it edits with a dropdown.
 */
export function spokenText(store: WorkbookStore, view: SheetView, row: number, col: number, merge: Rect | undefined, shown: string, list: boolean): string {
  const t = store.t;
  const at = merge ? { row: merge.r1, col: merge.c1 } : { row, col };
  const rec = view.get(at.row, at.col);
  // The cells a merge covers are read as the merge, whose top-left cell holds its value.
  if (merge && !shown && rec) shown = displayText(rec, effectiveStyle(view, store.styles, at.row, at.col, rec.style), store.date1904, store.numbers.display === 'full', store.locale).text;
  const parts = [toA1(at.row, at.col), shown || t('grid.empty')];
  const link = cellLink(view, at.row, at.col);
  if (link) {
    const address = linkAddress(link);
    parts[1] = address ? t('links.spoken', { value: parts[1], address }) : t('links.spokenUnknown', { value: parts[1] });
  }
  if (merge) parts.push(t('grid.flag.merged', { range: rangeToA1(merge) }));
  if (rec && rec.flags & CellFlag.Formula) parts.push(t('grid.flag.formula'));
  // Where a spilled value comes from, which the grid outlines only while it is selected.
  if (rec?.spill && rec.flags & CellFlag.Spilled) parts.push(t('grid.flag.spilled', { ref: toA1(rec.spill.r1, rec.spill.c1) }));
  if (rec && rec.flags & CellFlag.Stale) parts.push(t('grid.flag.stale'));
  else if (rec && rec.flags & (CellFlag.Unsupported | CellFlag.Mismatch)) parts.push(t('grid.flag.savedValue'));
  if (rec && store.highlightEdits && rec.flags & CellFlag.Edited) parts.push(t('grid.flag.edited'));
  if (list) parts.push(t('grid.flag.list'));
  const note = view.noteAt(at.row, at.col);
  if (note) parts.push(noteSummary(t, note));
  for (const i of store.schema?.cellIssues(view.info.index, row, col) ?? []) {
    parts.push(t(i.severity === 'error' ? 'grid.spokenError' : 'grid.spokenWarning', { message: t.text(i.message) }));
  }
  return parts.join(', ');
}
