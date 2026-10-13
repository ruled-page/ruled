import type { WorkbookStore } from '../state/store';
import { APPLE, commandKey } from '../ui/keys';
import { objectKey } from '../ui/objectGeometry';

/** A key pressed on the grid. */
export interface GridKey {
  key: string;
  keyCode: number;
  shiftKey: boolean;
  ctrlKey: boolean;
  metaKey: boolean;
  altKey: boolean;
}

/** Where a heading's filter button is, for its menu to open under. */
export type FilterButton = NonNullable<WorkbookStore['filterPicker']>;

/**
 * The editor's own keys on the grid, Excel's where Excel has them, taken
 * before the grid moves the selection; says whether the key was one.
 * `filterButton` gives a cell's filter button, if it has one.
 */
export function editorKey(store: WorkbookStore, e: GridKey, filterButton: (row: number, col: number) => FilterButton | undefined): boolean {
  const sel = store.selection;
  const plain = !e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey;
  // Space ticks or unticks the selection's checkboxes.
  if (e.key === ' ' && plain && store.checkboxAt(sel.row, sel.col) !== undefined) {
    void store.toggleCheckboxes();
    return true;
  }
  // Ctrl+Enter (⌘Enter on a Mac) follows the active cell's link, as Ctrl+click does.
  if (e.key === 'Enter' && commandKey(e) && !e.shiftKey && !e.altKey) {
    void store.followLink();
    return true;
  }
  // Alt+Shift+Right and Left group and ungroup the selection.
  if ((e.key === 'ArrowRight' || e.key === 'ArrowLeft') && e.altKey && e.shiftKey && !e.ctrlKey && !e.metaKey) {
    void store.groupSelection(e.key === 'ArrowRight' ? 1 : -1);
    return true;
  }
  // Whole lines: Shift+Space the selection's rows, Ctrl+Space its columns, Ctrl+A (⌘A) the sheet.
  if (e.key === ' ' && !e.altKey && !e.metaKey && e.shiftKey !== e.ctrlKey) {
    const axis = e.shiftKey ? 'rows' : 'cols';
    store.selectLines(axis, sel.rects.map((r) => (axis === 'rows' ? [r.r1, r.r2] : [r.c1, r.c2])), { row: sel.row, col: sel.col });
    return true;
  }
  if (e.key.toLowerCase() === 'a' && (APPLE ? e.metaKey : e.ctrlKey) && !e.shiftKey && !e.altKey) {
    store.selectAll();
    return true;
  }
  // Alt+Down on a heading with a filter button opens its menu.
  if (e.key === 'ArrowDown' && e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey) {
    const button = filterButton(sel.row, sel.col);
    if (button) {
      store.set('filterPicker', button);
      return true;
    }
  }
  // Ctrl+Alt+5 selects the sheet's first chart, picture or shape; Tab then goes through them.
  if (e.keyCode === 53 && e.ctrlKey && e.altKey && !e.metaKey && !e.shiftKey) {
    const first = store.editableObjects()[0];
    if (first) {
      store.selectObject(objectKey(first));
      return true;
    }
  }
  // Alt+F5 refreshes the sheet's pivot tables and Ctrl+Alt+F5 every sheet's.
  if (e.key === 'F5' && e.altKey && !e.metaKey && !e.shiftKey) {
    void store.refreshPivots(e.ctrlKey);
    return true;
  }
  // Alt+F1 inserts a column chart of the data.
  if (e.key === 'F1' && e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey && store.features.charts) {
    void store.insertChart('column');
    return true;
  }
  // Ctrl+Shift+L, or ⇧⌘F on a Mac, turns the filter buttons on or off.
  const letter = e.key.toLowerCase();
  if (commandKey(e) && e.shiftKey && !e.altKey && (letter === 'l' || (APPLE && letter === 'f'))) {
    void store.toggleFilter();
    return true;
  }
  // Ctrl++ (Ctrl+Shift+= on a US keyboard) inserts and Ctrl+- deletes, in place of the browser's zoom.
  const insert = e.key === '+' || (e.key === '=' && e.shiftKey);
  if ((insert || e.key === '-') && commandKey(e) && !e.altKey) {
    store.changeCells(insert ? 'insert' : 'delete');
    return true;
  }
  return false;
}
