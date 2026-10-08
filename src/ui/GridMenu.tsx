import { MAX_COLS, MAX_ROWS, toA1 } from '../core/refs';
import type { WorkbookStore } from '../state/store';
import { outlineDepth } from '../state/outline';
import { ContextMenu, type MenuItem } from './ContextMenu';
import { APPLE, mod } from './keys';
import { openDefineName } from '../state/names';
import { useT } from '../i18n/react';

/** The grid's right-click menu: clipboard, notes and comments, rows and columns, sorting, filtering and find. */
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
  // Grouping, ungrouping and the detail and levels of the sheet's groups.
  const outlineItems = (axis: 'rows' | 'cols'): MenuItem[] => {
    const rows = axis === 'rows';
    const disabled = store.outlineLocked(axis);
    const depth = outlineDepth(view.info.outline?.[axis]);
    const levels = Array.from({ length: depth ? depth + 1 : 0 }, (_, i) => ({
      label: t('outline.menu.showLevel', { level: i + 1 }),
      onSelect: () => void store.showOutlineLevel(axis, i + 1),
      disabled,
    }));
    return [
      { separator: true },
      { label: t(rows ? 'outline.menu.groupRows' : 'outline.menu.groupColumns'), hint: APPLE ? '⌥⇧→' : t('outline.keys.group'), onSelect: () => void store.groupLines(axis), disabled },
      ...(depth
        ? [
            { label: t(rows ? 'outline.menu.ungroupRows' : 'outline.menu.ungroupColumns'), hint: APPLE ? '⌥⇧←' : t('outline.keys.ungroup'), onSelect: () => void store.groupLines(axis, -1), disabled },
            ...(store.detailGroup(axis, false) ? [{ label: t('outline.menu.hideDetail'), onSelect: () => void store.showDetail(axis, false), disabled }] : []),
            ...(store.detailGroup(axis, true) ? [{ label: t('outline.menu.showDetail'), onSelect: () => void store.showDetail(axis, true), disabled }] : []),
            ...levels,
            { label: t(rows ? 'outline.menu.clearRows' : 'outline.menu.clearColumns'), onSelect: () => void store.clearOutline(axis), disabled },
          ]
        : []),
    ];
  };
  const rowItems: MenuItem[] = [
    { label: t('menu.insertRowsAbove', { count: rows }), onSelect: () => void store.insertRows('above'), disabled: locked || rows >= MAX_ROWS },
    { label: t('menu.insertRowsBelow', { count: rows }), onSelect: () => void store.insertRows('below'), disabled: locked || rows >= MAX_ROWS },
    { label: t('menu.deleteRows', { count: rows }), onSelect: () => void store.deleteRows(), disabled: locked, danger: true },
  ];
  const rowLayout: MenuItem[] = [
    { label: t('menu.rowHeight'), onSelect: () => store.set('rowHeightFor', sel.row), disabled: locked },
    { label: t('menu.autofitRowHeight'), onSelect: () => void store.autofit('rows'), disabled: locked },
    { label: t('menu.hideRows', { count: rows }), onSelect: () => void store.hideLines('rows'), disabled: store.outlineLocked('rows') },
    ...(store.hasHiddenNear('rows') ? [{ label: t('menu.unhideRows'), onSelect: () => void store.unhideLines('rows'), disabled: store.outlineLocked('rows') }] : []),
    ...outlineItems('rows'),
  ];
  const colItems: MenuItem[] = [
    { label: t('menu.insertColumnsLeft', { count: cols }), onSelect: () => void store.insertCols('left'), disabled: locked || cols >= MAX_COLS },
    { label: t('menu.insertColumnsRight', { count: cols }), onSelect: () => void store.insertCols('right'), disabled: locked || cols >= MAX_COLS },
    { label: t('menu.deleteColumns', { count: cols }), onSelect: () => void store.deleteCols(), disabled: locked, danger: true },
  ];
  const colLayout: MenuItem[] = [
    { label: t('menu.columnWidth'), onSelect: () => store.set('columnWidthFor', sel.col), disabled: locked },
    { label: t('menu.autofitColumnWidth'), onSelect: () => void store.autofit('cols'), disabled: locked },
    { label: t('menu.hideColumns', { count: cols }), onSelect: () => void store.hideLines('cols'), disabled: store.outlineLocked('cols') },
    ...(store.hasHiddenNear('cols') ? [{ label: t('menu.unhideColumns'), onSelect: () => void store.unhideLines('cols'), disabled: store.outlineLocked('cols') }] : []),
    ...outlineItems('cols'),
  ];
  const frozen = view.info.frozenRows > 0 || view.info.frozenCols > 0;
  const merged = store.mergesInSelection().length > 0;
  const freezeItems: MenuItem[] = [
    ...(sel.row > 0 || sel.col > 0 ? [{ label: t('menu.freezePanes', { cell: toA1(sel.row, sel.col) }), onSelect: () => void store.freezePanes('cell') }] : []),
    { label: t('menu.freezeTopRow'), onSelect: () => void store.freezePanes('topRow') },
    { label: t('menu.freezeFirstColumn'), onSelect: () => void store.freezePanes('firstColumn') },
    ...(frozen ? [{ label: t('menu.unfreeze'), onSelect: () => void store.freezePanes('none') }] : []),
  ];
  // On a cell, rows, columns and freezing are submenus; on a heading, its own items are the menu.
  const lineItems: MenuItem[] =
    menu.kind === 'column' || (wholeCols && !wholeRows)
      ? [...colItems, ...colLayout]
      : menu.kind === 'row' || wholeRows
        ? [...rowItems, ...rowLayout]
        : [
            { label: t('menu.rows'), items: [...rowItems, { separator: true }, ...rowLayout] },
            { label: t('menu.columns'), items: [...colItems, { separator: true }, ...colLayout] },
            ...(rows * cols > 1 ? [{ label: t('menu.merge'), onSelect: () => void store.mergeCells(), disabled: locked }] : []),
            ...(merged ? [{ label: t('menu.unmerge'), onSelect: () => void store.unmergeCells(), disabled: locked }] : []),
            { label: t('menu.freeze'), items: freezeItems },
          ];
  // The active cell's note or comments; read-only ones can still be read.
  const note = view.noteAt(sel.row, sel.col);
  const notesLocked = !!view.commentReadOnlyReason(sel.row, sel.col);
  const openComment = (start: 'note' | 'thread') => () => store.openComment(sel.row, sel.col, { x: menu.x, y: menu.y }, start);
  const deleteComment = () => void store.deleteComment(sel.row, sel.col);
  const noteItems: MenuItem[] = !note
    ? [
        { label: t('menu.newNote'), hint: t('keys.shiftF2'), onSelect: openComment('note'), disabled: notesLocked },
        { label: t('menu.newComment'), onSelect: openComment('thread'), disabled: notesLocked },
      ]
    : note.kind === 'note'
      ? [
          { label: t(notesLocked ? 'menu.showNote' : 'menu.editNote'), hint: t('keys.shiftF2'), onSelect: openComment('note') },
          { label: t('menu.deleteNote'), onSelect: deleteComment, disabled: notesLocked, danger: true },
        ]
      : [
          { label: t('menu.showComments'), hint: t('keys.shiftF2'), onSelect: openComment('thread') },
          { label: t('menu.deleteThread'), onSelect: deleteComment, disabled: notesLocked, danger: true },
        ];
  // Links: open the cell's, then give the selection one, or change or take away theirs.
  const link = menu.kind === 'cell' ? store.linkAt() : undefined;
  const linked = menu.kind === 'cell' ? store.linksInSelection() : 0;
  const openDialog = () => store.set('linkDialog', true);
  const linkItems: MenuItem[] =
    menu.kind !== 'cell'
      ? []
      : [
          { separator: true },
          ...(link && store.followsLinks ? [{ label: t('links.menu.open'), hint: mod(t, t('keys.enter')), onSelect: () => void store.followLink() }] : []),
          ...(linked
            ? [
                {
                  label: t('links.menu.link'),
                  items: [
                    { label: t('links.menu.edit'), onSelect: openDialog, disabled: locked },
                    { label: t('links.menu.remove', { count: linked }), onSelect: () => void store.removeLinks(), disabled: locked },
                  ],
                },
              ]
            : [{ label: t('links.menu.insert'), onSelect: openDialog, disabled: locked }]),
        ];
  const filtered = !!view.filter;
  const dataItems: MenuItem[] = [
    { label: t('menu.sortAscending'), onSelect: () => void store.sort(false), disabled: locked },
    { label: t('menu.sortDescending'), onSelect: () => void store.sort(true), disabled: locked },
    { label: t('menu.filter'), onSelect: () => store.set('filterPicker', { col: sel.col, x: menu.x, y: menu.y }) },
    ...(filtered ? [{ label: t('menu.clearFilter'), onSelect: () => void store.clearFilter() }] : []),
  ];
  const items: MenuItem[] = [
    { label: t('menu.cut'), hint: mod(t, 'X'), onSelect: () => void store.copySelection(true), disabled: locked },
    { label: t('menu.copy'), hint: mod(t, 'C'), onSelect: () => void store.copySelection(false) },
    { label: t('menu.paste'), hint: mod(t, 'V'), onSelect: () => void store.pasteFromClipboard(), disabled: cellsLocked },
    {
      label: t('menu.pasteSpecial'),
      disabled: cellsLocked,
      items: [
        { label: t('menu.pasteValues'), onSelect: () => void store.pasteSpecial('values') },
        { label: t('menu.pasteFormulas'), onSelect: () => void store.pasteSpecial('formulas') },
        { label: t('menu.pasteFormats'), onSelect: () => void store.pasteSpecial('formats') },
        { label: t('menu.pasteTransposed'), onSelect: () => void store.pasteSpecial('all', true) },
      ],
    },
    { label: t('menu.fillDown'), onSelect: () => void store.fillDown(), disabled: cellsLocked },
    { label: t('menu.fillRight'), onSelect: () => void store.fillDown(true), disabled: cellsLocked },
    { label: t('fill.menu.series'), onSelect: () => store.set('seriesDialog', true), disabled: cellsLocked },
    { separator: true },
    ...(menu.kind === 'cell' ? [...noteItems, { separator: true } as const] : []),
    ...lineItems,
    { separator: true },
    ...dataItems,
    ...linkItems,
    { separator: true },
    { label: t('names.menu.define'), onSelect: () => openDefineName(store), disabled: view.info.kind !== 'worksheet' },
    { label: t('names.menu.manager'), onSelect: () => store.set('nameManagerOpen', true) },
    { separator: true },
    { label: t('menu.find'), hint: mod(t, 'F'), onSelect: () => store.set('findOpen', true) },
  ];
  return <ContextMenu x={menu.x} y={menu.y} items={items} label={t('menu.label')} onClose={() => store.set('menu', undefined)} />;
}
