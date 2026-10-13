import type { Rect } from '../core/types';
import type { SheetView } from '../state/sheetView';
import type { Selection, WorkbookStore } from '../state/store';
import type { GridPalette } from '../ui/gridPalette';
import { withAlpha } from './color';

/** A range the grid marks: filled in `color`, which may be see-through, and outlined in it, solid or dashed. */
export interface Mark {
  rect: Rect;
  color: string;
  style: 'solid' | 'dashed';
}

/**
 * What the grid marks besides the selection: a cut waiting to be pasted,
 * where cells dragged by the selection's border would land, the range the
 * selected cell's formula spills into (dashed while something is in its
 * way), and the references of a formula being typed, in their colours.
 */
export function sheetMarks(store: WorkbookStore, view: SheetView, selection: Selection, palette: GridPalette, refColors: readonly string[], landing?: Rect): Mark[] {
  const marks: Mark[] = [];
  const cut = store.cutRange;
  if (cut && cut.sheet === view.info.index) marks.push({ rect: cut.rect, color: palette.cutFill, style: 'dashed' });
  if (landing) marks.push({ rect: landing, color: palette.selectionFill, style: 'dashed' });
  const spill = view.get(selection.row, selection.col)?.spill;
  if (spill) marks.push({ rect: spill, color: withAlpha(palette.spillLine, 0), style: view.get(spill.r1, spill.c1)?.s === '#SPILL!' ? 'dashed' : 'solid' });
  for (const ref of store.formulaRefs) {
    if (ref.sheet === view.info.index) marks.push({ rect: ref.rect, color: withAlpha(refColors[ref.color % refColors.length], 0.1), style: 'solid' });
  }
  return marks;
}
