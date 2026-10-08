import { MAX_COLS, MAX_ROWS } from '../core/refs';
import type { WorkbookStore } from '../state/store';
import { ContextMenu, type MenuItem } from './ContextMenu';
import { mod } from './keys';
import { useT } from '../i18n/react';

/** The grid's right-click menu: clipboard, rows and columns, sorting, filtering and find. */
export function GridMenu({ store }: { store: WorkbookStore }) {
  const t = useT();
  const menu = store.menu;
  const view = store.view;
  if (!menu || !view) return null;
  const sel = store.selection;
  const rect = sel.rects[0] ?? { r1: sel.row, r2: sel.row, c1: sel.col, c2: sel.col };
  const rows = rect.r2 - rect.r1 + 1;
  const cols = rect.c2 - rect.c1 + 1;
  const wholeRows = rect.c1 === 0 && rect.c2 >= Math.max(view.info.maxCol, 0);
  const wholeCols = rect.r1 === 0 && rect.r2 >= Math.max(view.info.maxRow, 0);
  // Rows, columns and sorting are protected with the sheet; cells only where their format locks them.
  const locked = view.info.kind !== 'worksheet' || view.info.protected;
  const unlocked = () => {
    if (rows * cols > 10_000) return false;
    for (let r = rect.r1; r <= rect.r2; r++) for (let c = rect.c1; c <= rect.c2; c++) if (!view.isUnlocked(view.styleAt(r, c))) return false;
    return true;
  };
  const cellsLocked = view.info.kind !== 'worksheet' || (view.info.protected && !unlocked());
  const rowItems: MenuItem[] = [
    { label: t('menu.insertRowsAbove', { count: rows }), onSelect: () => void store.insertRows('above'), disabled: locked || rows >= MAX_ROWS },
    { label: t('menu.insertRowsBelow', { count: rows }), onSelect: () => void store.insertRows('below'), disabled: locked || rows >= MAX_ROWS },
    { label: t('menu.deleteRows', { count: rows }), onSelect: () => void store.deleteRows(), disabled: locked, danger: true },
  ];
  const colItems: MenuItem[] = [
    { label: t('menu.insertColumnsLeft', { count: cols }), onSelect: () => void store.insertCols('left'), disabled: locked || cols >= MAX_COLS },
    { label: t('menu.insertColumnsRight', { count: cols }), onSelect: () => void store.insertCols('right'), disabled: locked || cols >= MAX_COLS },
    { label: t('menu.deleteColumns', { count: cols }), onSelect: () => void store.deleteCols(), disabled: locked, danger: true },
    { label: t('menu.columnWidth'), onSelect: () => store.set('columnWidthFor', sel.col) },
  ];
  const filtered = !!view.filter;
  const dataItems: MenuItem[] = [
    { label: t('menu.sortAscending'), onSelect: () => void store.sort(false), disabled: locked },
    { label: t('menu.sortDescending'), onSelect: () => void store.sort(true), disabled: locked },
    { label: t('menu.filter'), onSelect: () => store.set('filterPicker', { col: sel.col, x: menu.x, y: menu.y }) },
    ...(filtered ? [{ label: t('menu.clearFilter'), onSelect: () => store.clearFilter() }] : []),
  ];
  const items: MenuItem[] = [
    { label: t('menu.cut'), hint: mod(t, 'X'), onSelect: () => void store.copySelection(true), disabled: locked },
    { label: t('menu.copy'), hint: mod(t, 'C'), onSelect: () => void store.copySelection(false) },
    { label: t('menu.paste'), hint: mod(t, 'V'), onSelect: () => void store.pasteFromClipboard(), disabled: cellsLocked },
    { label: t('menu.fillDown'), onSelect: () => void store.fillDown(), disabled: cellsLocked },
    { label: t('menu.fillRight'), onSelect: () => void store.fillDown(true), disabled: cellsLocked },
    { separator: true },
    ...(menu.kind === 'column' || (wholeCols && !wholeRows) ? colItems : menu.kind === 'row' || wholeRows ? rowItems : [...rowItems, { separator: true } as const, ...colItems]),
    { separator: true },
    ...dataItems,
    { separator: true },
    { label: t('menu.find'), hint: mod(t, 'F'), onSelect: () => store.set('findOpen', true) },
  ];
  return <ContextMenu x={menu.x} y={menu.y} items={items} onClose={() => store.set('menu', undefined)} />;
}
