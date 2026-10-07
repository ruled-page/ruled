import { MAX_COLS, MAX_ROWS } from '../core/refs';
import type { WorkbookStore } from '../state/store';
import { ContextMenu, type MenuItem } from './ContextMenu';

const mod = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform) ? '⌘' : 'Ctrl+';
const plural = (n: number, one: string) => `${n} ${one}${n === 1 ? '' : 's'}`;

/** The grid's right-click menu: clipboard, rows and columns, sorting, filtering and find. */
export function GridMenu({ store }: { store: WorkbookStore }) {
  const menu = store.menu;
  const view = store.view;
  if (!menu || !view) return null;
  const sel = store.selection;
  const rect = sel.rects[0] ?? { r1: sel.row, r2: sel.row, c1: sel.col, c2: sel.col };
  const rows = rect.r2 - rect.r1 + 1;
  const cols = rect.c2 - rect.c1 + 1;
  const wholeRows = rect.c1 === 0 && rect.c2 >= Math.max(view.info.maxCol, 0);
  const wholeCols = rect.r1 === 0 && rect.r2 >= Math.max(view.info.maxRow, 0);
  const locked = view.info.kind !== 'worksheet' || view.info.protected;
  const rowItems: MenuItem[] = [
    { label: `Insert ${plural(rows, 'row')} above`, onSelect: () => void store.insertRows('above'), disabled: locked || rows >= MAX_ROWS },
    { label: `Insert ${plural(rows, 'row')} below`, onSelect: () => void store.insertRows('below'), disabled: locked || rows >= MAX_ROWS },
    { label: `Delete ${plural(rows, 'row')}`, onSelect: () => void store.deleteRows(), disabled: locked, danger: true },
  ];
  const colItems: MenuItem[] = [
    { label: `Insert ${plural(cols, 'column')} left`, onSelect: () => void store.insertCols('left'), disabled: locked || cols >= MAX_COLS },
    { label: `Insert ${plural(cols, 'column')} right`, onSelect: () => void store.insertCols('right'), disabled: locked || cols >= MAX_COLS },
    { label: `Delete ${plural(cols, 'column')}`, onSelect: () => void store.deleteCols(), disabled: locked, danger: true },
  ];
  const filtered = !!view.filter;
  const dataItems: MenuItem[] = [
    { label: 'Sort A to Z', onSelect: () => void store.sort(false), disabled: locked },
    { label: 'Sort Z to A', onSelect: () => void store.sort(true), disabled: locked },
    { label: 'Filter by values…', onSelect: () => store.set('filterPicker', { col: sel.col, x: menu.x, y: menu.y }) },
    ...(filtered ? [{ label: 'Show all rows (clear filter)', onSelect: () => store.clearFilter() }] : []),
  ];
  const items: MenuItem[] = [
    { label: 'Cut', hint: `${mod}X`, onSelect: () => void store.copySelection(true), disabled: locked },
    { label: 'Copy', hint: `${mod}C`, onSelect: () => void store.copySelection(false) },
    { label: 'Paste', hint: `${mod}V`, onSelect: () => void store.pasteFromClipboard(), disabled: locked },
    { separator: true },
    ...(menu.kind === 'column' || (wholeCols && !wholeRows) ? colItems : menu.kind === 'row' || wholeRows ? rowItems : [...rowItems, { separator: true } as const, ...colItems]),
    { separator: true },
    ...dataItems,
    { separator: true },
    { label: 'Find and replace…', hint: `${mod}F`, onSelect: () => store.set('findOpen', true) },
  ];
  return <ContextMenu x={menu.x} y={menu.y} items={items} onClose={() => store.set('menu', undefined)} />;
}
