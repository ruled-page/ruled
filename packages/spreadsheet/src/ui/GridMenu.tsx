import { MAX_COLS, MAX_ROWS, toA1 } from '@ruled/formula/refs';
import { useEffect } from 'react';
import type { WorkbookStore } from '../state/store';
import type { Text } from '../i18n/message';
import { outlineDepth } from '../state/outline';
import { ContextMenu, keepItems, type MenuItem } from './ContextMenu';
import { APPLE, mod } from './keys';
import { openDefineName } from '../state/names';
import { useT } from '../i18n/react';
import type { CfQuick } from '../state/cfRules';
import type { NewChartKind } from '../core/chartWriter';

/** The charts Insert Chart offers, in Excel's order. */
const CHART_KINDS: NewChartKind[] = ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter'];

/**
 * The grid's right-click menu: clipboard, notes and comments, rows and columns, sorting, filtering and find;
 * on a heading, its rows' or columns'. Each entry's id is in MENU_COMMANDS, and the `menus` option keeps some only.
 */
export function GridMenu({ store }: { store: WorkbookStore }) {
  const t = useT();
  const menu = store.menu;
  const view = store.view;
  const plugins = store.plugins;
  const commands = plugins?.commands.filter((c) => c.where === 'cellMenu') ?? [];
  // Plugins load what their commands need as the menu opens, so a popup a command opens isn't blocked.
  useEffect(() => {
    if (menu) plugins?.preload(plugins.commands.filter((c) => c.where === 'cellMenu'));
  }, [plugins, menu]);
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
    const ids = rows
      ? ({ group: 'outline.menu.groupRows', ungroup: 'outline.menu.ungroupRows', clear: 'outline.menu.clearRows' } as const)
      : ({ group: 'outline.menu.groupColumns', ungroup: 'outline.menu.ungroupColumns', clear: 'outline.menu.clearColumns' } as const);
    const disabled = store.outlineLocked(axis);
    const depth = outlineDepth(view.info.outline?.[axis]);
    const levels = Array.from({ length: depth ? depth + 1 : 0 }, (_, i) => ({
      id: 'outline.menu.showLevel',
      label: t('outline.menu.showLevel', { level: i + 1 }),
      onSelect: () => void store.showOutlineLevel(axis, i + 1),
      disabled,
    }));
    return [
      { separator: true },
      { id: ids.group, label: t(ids.group), hint: APPLE ? '⌥⇧→' : t('outline.keys.group'), onSelect: () => void store.groupLines(axis), disabled },
      ...(depth
        ? [
            { id: ids.ungroup, label: t(ids.ungroup), hint: APPLE ? '⌥⇧←' : t('outline.keys.ungroup'), onSelect: () => void store.groupLines(axis, -1), disabled },
            ...(store.detailGroup(axis, false) ? [{ id: 'outline.menu.hideDetail', label: t('outline.menu.hideDetail'), onSelect: () => void store.showDetail(axis, false), disabled }] : []),
            ...(store.detailGroup(axis, true) ? [{ id: 'outline.menu.showDetail', label: t('outline.menu.showDetail'), onSelect: () => void store.showDetail(axis, true), disabled }] : []),
            ...levels,
            { id: ids.clear, label: t(ids.clear), onSelect: () => void store.clearOutline(axis), disabled },
          ]
        : []),
    ];
  };
  const rowItems: MenuItem[] = [
    { id: 'menu.insertRowsAbove', label: t('menu.insertRowsAbove', { count: rows }), onSelect: () => void store.insertRows('above'), disabled: locked || rows >= MAX_ROWS },
    { id: 'menu.insertRowsBelow', label: t('menu.insertRowsBelow', { count: rows }), onSelect: () => void store.insertRows('below'), disabled: locked || rows >= MAX_ROWS },
    { id: 'menu.deleteRows', label: t('menu.deleteRows', { count: rows }), onSelect: () => void store.deleteRows(), disabled: locked, danger: true },
  ];
  const rowLayout: MenuItem[] = [
    { id: 'menu.rowHeight', label: t('menu.rowHeight'), onSelect: () => store.set('rowHeightFor', sel.row), disabled: locked },
    { id: 'menu.autofitRowHeight', label: t('menu.autofitRowHeight'), onSelect: () => void store.autofit('rows'), disabled: locked },
    { id: 'menu.hideRows', label: t('menu.hideRows', { count: rows }), onSelect: () => void store.hideLines('rows'), disabled: store.outlineLocked('rows') },
    ...(store.hasHiddenNear('rows') ? [{ id: 'menu.unhideRows', label: t('menu.unhideRows'), onSelect: () => void store.unhideLines('rows'), disabled: store.outlineLocked('rows') }] : []),
    ...outlineItems('rows'),
  ];
  const colItems: MenuItem[] = [
    { id: 'menu.insertColumnsLeft', label: t('menu.insertColumnsLeft', { count: cols }), onSelect: () => void store.insertCols('left'), disabled: locked || cols >= MAX_COLS },
    { id: 'menu.insertColumnsRight', label: t('menu.insertColumnsRight', { count: cols }), onSelect: () => void store.insertCols('right'), disabled: locked || cols >= MAX_COLS },
    { id: 'menu.deleteColumns', label: t('menu.deleteColumns', { count: cols }), onSelect: () => void store.deleteCols(), disabled: locked, danger: true },
  ];
  const colLayout: MenuItem[] = [
    { id: 'menu.columnWidth', label: t('menu.columnWidth'), onSelect: () => store.set('columnWidthFor', sel.col), disabled: locked },
    { id: 'menu.autofitColumnWidth', label: t('menu.autofitColumnWidth'), onSelect: () => void store.autofit('cols'), disabled: locked },
    { id: 'menu.hideColumns', label: t('menu.hideColumns', { count: cols }), onSelect: () => void store.hideLines('cols'), disabled: store.outlineLocked('cols') },
    ...(store.hasHiddenNear('cols') ? [{ id: 'menu.unhideColumns', label: t('menu.unhideColumns'), onSelect: () => void store.unhideLines('cols'), disabled: store.outlineLocked('cols') }] : []),
    ...outlineItems('cols'),
  ];
  const frozen = view.info.frozenRows > 0 || view.info.frozenCols > 0;
  const merged = store.mergesInSelection().length > 0;
  const freezeItems: MenuItem[] = [
    ...(sel.row > 0 || sel.col > 0 ? [{ id: 'menu.freezePanes', label: t('menu.freezePanes', { cell: toA1(sel.row, sel.col) }), onSelect: () => void store.freezePanes('cell') }] : []),
    { id: 'menu.freezeTopRow', label: t('menu.freezeTopRow'), onSelect: () => void store.freezePanes('topRow') },
    { id: 'menu.freezeFirstColumn', label: t('menu.freezeFirstColumn'), onSelect: () => void store.freezePanes('firstColumn') },
    ...(frozen ? [{ id: 'menu.unfreeze', label: t('menu.unfreeze'), onSelect: () => void store.freezePanes('none') }] : []),
  ];
  // On a cell, rows, columns and freezing are submenus; on a heading, its own items are the menu.
  // Shown on the cell menu in place of the Rows or Columns submenu, they're kept with it.
  const within = (id: string, list: MenuItem[]) => (menu.kind === 'cell' ? list.map((item) => ('separator' in item ? item : { ...item, within: id })) : list);
  const lineItems: MenuItem[] =
    menu.kind === 'column' || (wholeCols && !wholeRows)
      ? within('menu.columns', [...colItems, ...colLayout])
      : menu.kind === 'row' || wholeRows
        ? within('menu.rows', [...rowItems, ...rowLayout])
        : [
            { id: 'menu.insertCells', label: t('menu.insertCells'), hint: mod(t, '+'), onSelect: () => store.set('cellsDialog', 'insert'), disabled: locked || sel.rects.length !== 1 },
            { id: 'menu.deleteCells', label: t('menu.deleteCells'), hint: mod(t, '-'), onSelect: () => store.set('cellsDialog', 'delete'), disabled: locked || sel.rects.length !== 1, danger: true },
            { id: 'menu.rows', label: t('menu.rows'), items: [...rowItems, { separator: true }, ...rowLayout] },
            { id: 'menu.columns', label: t('menu.columns'), items: [...colItems, { separator: true }, ...colLayout] },
            ...(rows * cols > 1 ? [{ id: 'menu.merge', label: t('menu.merge'), onSelect: () => void store.mergeCells(), disabled: locked }] : []),
            ...(merged ? [{ id: 'menu.unmerge', label: t('menu.unmerge'), onSelect: () => void store.unmergeCells(), disabled: locked }] : []),
            { id: 'menu.freeze', label: t('menu.freeze'), items: freezeItems },
          ];
  // The active cell's note or comments; read-only ones can still be read.
  const note = view.noteAt(sel.row, sel.col);
  const notesLocked = !!view.commentReadOnlyReason(sel.row, sel.col);
  const openComment = (start: 'note' | 'thread') => () => store.openComment(sel.row, sel.col, { x: menu.x, y: menu.y }, start);
  const deleteComment = () => void store.deleteComment(sel.row, sel.col);
  const noteItems: MenuItem[] = !note
    ? [
        { id: 'menu.newNote', label: t('menu.newNote'), hint: t('keys.shiftF2'), onSelect: openComment('note'), disabled: notesLocked },
        { id: 'menu.newComment', label: t('menu.newComment'), onSelect: openComment('thread'), disabled: notesLocked },
      ]
    : note.kind === 'note'
      ? [
          { id: 'menu.editNote', label: t(notesLocked ? 'menu.showNote' : 'menu.editNote'), hint: t('keys.shiftF2'), onSelect: openComment('note') },
          { id: 'menu.deleteNote', label: t('menu.deleteNote'), onSelect: deleteComment, disabled: notesLocked, danger: true },
        ]
      : [
          { id: 'menu.showComments', label: t('menu.showComments'), hint: t('keys.shiftF2'), onSelect: openComment('thread') },
          { id: 'menu.deleteThread', label: t('menu.deleteThread'), onSelect: deleteComment, disabled: notesLocked, danger: true },
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
          ...(link && store.followsLinks ? [{ id: 'links.menu.open', label: t('links.menu.open'), hint: mod(t, t('keys.enter')), onSelect: () => void store.followLink() }] : []),
          ...(linked
            ? [
                {
                  id: 'links.menu.link',
                  label: t('links.menu.link'),
                  items: [
                    { id: 'links.menu.edit', label: t('links.menu.edit'), onSelect: openDialog, disabled: locked },
                    { id: 'links.menu.remove', label: t('links.menu.remove', { count: linked }), onSelect: () => void store.removeLinks(), disabled: locked },
                  ],
                },
              ]
            : [{ id: 'links.menu.insert', label: t('links.menu.insert'), onSelect: openDialog, disabled: locked }]),
        ];
  const filtered = view.filtering;
  // Excel's Conditional Formatting menu, in its groups: quick rules, the galleries, a formula rule, clearing and the Rules Manager.
  const quick = (kinds: CfQuick[]): MenuItem[] => kinds.map((kind) => ({ id: `cf.quick.${kind}`, label: t(`cf.quick.${kind}`), onSelect: () => store.openCfDialog(kind) }));
  const cfItems: MenuItem[] = [
    ...quick(['greaterThan', 'lessThan', 'between', 'equal', 'textContains', 'duplicate']),
    { separator: true },
    ...quick(['top', 'topPercent', 'bottom', 'bottomPercent', 'aboveAverage', 'belowAverage']),
    { separator: true },
    { id: 'cf.menu.gallery', label: t('cf.menu.gallery'), onSelect: () => store.openCfGallery() },
    ...quick(['formula']),
    { separator: true },
    { id: 'cf.menu.clearSelection', label: t('cf.menu.clearSelection'), onSelect: () => void store.clearConditionalFormats('selection') },
    { id: 'cf.menu.clearSheet', label: t('cf.menu.clearSheet'), onSelect: () => void store.clearConditionalFormats('sheet') },
    { id: 'cf.menu.manage', label: t('cf.menu.manage'), onSelect: () => store.openCfManager() },
  ];
  const pivotItems: MenuItem[] = store.inPivot
    ? [
        { id: 'pivot.menu.refresh', label: t('pivot.menu.refresh'), hint: APPLE ? '⌥F5' : t('keys.altF5'), onSelect: () => void store.refreshPivots() },
        { id: 'pivot.menu.refreshAll', label: t('pivot.menu.refreshAll'), hint: APPLE ? '⌃⌥F5' : t('keys.ctrlAltF5'), onSelect: () => void store.refreshPivots(true) },
        { id: 'pivot.menu.fieldList', label: t('pivot.menu.fieldList'), onSelect: () => store.showFieldList() },
        { separator: true },
      ]
    : [];
  const dataItems: MenuItem[] = [
    { id: 'menu.sortAscending', label: t('menu.sortAscending'), onSelect: () => void store.sort(false), disabled: locked },
    { id: 'menu.sortDescending', label: t('menu.sortDescending'), onSelect: () => void store.sort(true), disabled: locked },
    { id: 'sort.menu.custom', label: t('sort.menu.custom'), onSelect: () => store.openSortDialog(), disabled: locked },
    { id: 'menu.filter', label: t('menu.filter'), onSelect: () => store.set('filterPicker', { col: sel.col, x: menu.x, y: menu.y }) },
    { id: 'validation.menu', label: t('validation.menu'), onSelect: () => void store.openValidationDialog(), disabled: locked },
    { id: 'cf.menu', label: t('cf.menu'), items: cfItems, disabled: locked },
    ...(store.features.charts
      ? [
          {
            id: 'charts.menu.insert',
            label: t('charts.menu.insert'),
            disabled: view.info.kind !== 'worksheet' || !!view.info.protected,
            items: CHART_KINDS.map((k) => ({ id: `core.chartType.${k}`, label: t(`core.chartType.${k}`), onSelect: () => void store.insertChart(k) })),
          },
        ]
      : []),
    { id: 'pivot.menu.insert', label: t('pivot.menu.insert'), onSelect: () => store.openInsertPivot(), disabled: view.info.kind !== 'worksheet' },
    {
      id: 'menu.filterOn',
      label: t(view.filter && !view.filter.table ? 'menu.filterOff' : 'menu.filterOn'),
      hint: APPLE ? '⇧⌘F' : t('keys.filterToggle'),
      onSelect: () => void store.toggleFilter(),
      disabled: view.info.kind !== 'worksheet',
    },
    ...(filtered ? [{ id: 'menu.clearFilter', label: t('menu.clearFilter'), onSelect: () => void store.clearFilter() }] : []),
  ];
  const items: MenuItem[] = [
    { id: 'menu.cut', label: t('menu.cut'), hint: mod(t, 'X'), onSelect: () => void store.copySelection(true), disabled: locked },
    { id: 'menu.copy', label: t('menu.copy'), hint: mod(t, 'C'), onSelect: () => void store.copySelection(false) },
    { id: 'menu.paste', label: t('menu.paste'), hint: mod(t, 'V'), onSelect: () => void store.pasteFromClipboard(), disabled: cellsLocked },
    {
      id: 'menu.pasteSpecial',
      label: t('menu.pasteSpecial'),
      disabled: cellsLocked,
      items: [
        { id: 'menu.pasteValues', label: t('menu.pasteValues'), onSelect: () => void store.pasteSpecial('values') },
        { id: 'menu.pasteFormulas', label: t('menu.pasteFormulas'), onSelect: () => void store.pasteSpecial('formulas') },
        { id: 'menu.pasteFormats', label: t('menu.pasteFormats'), onSelect: () => void store.pasteSpecial('formats') },
        { id: 'menu.pasteTransposed', label: t('menu.pasteTransposed'), onSelect: () => void store.pasteSpecial('all', true) },
      ],
    },
    ...(store.touchMode ? [{ id: 'menu.fillDrag', label: t('menu.fillDrag'), onSelect: () => store.set('touchFill', true), disabled: cellsLocked }] : []),
    { id: 'menu.fillDown', label: t('menu.fillDown'), onSelect: () => void store.fillDown(), disabled: cellsLocked },
    { id: 'menu.fillRight', label: t('menu.fillRight'), onSelect: () => void store.fillDown(true), disabled: cellsLocked },
    { id: 'fill.menu.series', label: t('fill.menu.series'), onSelect: () => store.set('seriesDialog', true), disabled: cellsLocked },
    { separator: true },
    ...(menu.kind === 'cell' ? [...noteItems, { separator: true } as const] : []),
    ...lineItems,
    { separator: true },
    ...pivotItems,
    ...dataItems,
    ...linkItems,
    { separator: true },
    { id: 'names.menu.define', label: t('names.menu.define'), onSelect: () => openDefineName(store), disabled: view.info.kind !== 'worksheet' },
    { id: 'names.menu.manager', label: t('names.menu.manager'), onSelect: () => store.set('nameManagerOpen', true) },
    { separator: true },
    { id: 'menu.find', label: t('menu.find'), hint: mod(t, 'F'), onSelect: () => store.set('findOpen', true) },
  ];
  // Plugins' commands, in a group after the editor's own, given the selection.
  const context = { selection: store.commandSelection() };
  const added = commands.flatMap((c): MenuItem[] => {
    const state = plugins!.commandState(c, context);
    return state.visible ? [{ id: c.key, label: t.text(c.label as Text), onSelect: () => plugins!.run(c, context), disabled: !state.enabled }] : [];
  });
  if (added.length) items.push({ separator: true }, ...added);
  const kept = keepItems(items, menu.kind === 'cell' ? store.menus?.cell : store.menus?.header);
  return <ContextMenu x={menu.x} y={menu.y} items={kept} label={t('menu.label')} onClose={() => store.set('menu', undefined)} />;
}
