import type { WorkbookStore } from '../state/store';

/**
 * The ids of the cell and heading menus' entries, for hosts that keep only
 * some of them, and the shortcuts that go with those entries.
 */

/**
 * Where a built-in menu entry shows: in the cell menu, in the menu of the row
 * and column headings, or in both; and whether choosing it can change the
 * workbook (its cells, formats, layout or anything else saved with it).
 */
export interface MenuCommandInfo {
  menus: 'cell' | 'header' | 'both';
  edits: boolean;
}

/**
 * Every built-in entry of the cell menu and the heading menu, by id. An
 * entry's id is the message key of its label, and stays the same whatever the
 * language, the selection (Insert 3 Rows Above is `menu.insertRowsAbove`) or
 * the state the label shows: Filter on and off are both `menu.filterOn`, Edit
 * Note and Show Note `menu.editNote`. Submenus have ids of their own; Show
 * Detail, Hide Detail and Show Level are in both the Rows and the Columns submenus.
 */
export const MENU_COMMANDS = {
  // Clipboard and filling
  'menu.cut': { menus: 'both', edits: true },
  'menu.copy': { menus: 'both', edits: false },
  'menu.paste': { menus: 'both', edits: true },
  'menu.pasteSpecial': { menus: 'both', edits: true },
  'menu.pasteValues': { menus: 'both', edits: true },
  'menu.pasteFormulas': { menus: 'both', edits: true },
  'menu.pasteFormats': { menus: 'both', edits: true },
  'menu.pasteTransposed': { menus: 'both', edits: true },
  'menu.fillDrag': { menus: 'both', edits: true },
  'menu.fillDown': { menus: 'both', edits: true },
  'menu.fillRight': { menus: 'both', edits: true },
  'fill.menu.series': { menus: 'both', edits: true },
  // Notes and comments; opening one that can't be changed only shows it.
  'menu.newNote': { menus: 'cell', edits: true },
  'menu.newComment': { menus: 'cell', edits: true },
  'menu.editNote': { menus: 'cell', edits: false },
  'menu.deleteNote': { menus: 'cell', edits: true },
  'menu.showComments': { menus: 'cell', edits: false },
  'menu.deleteThread': { menus: 'cell', edits: true },
  // Cells, rows and columns
  'menu.insertCells': { menus: 'cell', edits: true },
  'menu.deleteCells': { menus: 'cell', edits: true },
  'menu.rows': { menus: 'cell', edits: true },
  'menu.insertRowsAbove': { menus: 'both', edits: true },
  'menu.insertRowsBelow': { menus: 'both', edits: true },
  'menu.deleteRows': { menus: 'both', edits: true },
  'menu.rowHeight': { menus: 'both', edits: true },
  'menu.autofitRowHeight': { menus: 'both', edits: true },
  'menu.hideRows': { menus: 'both', edits: true },
  'menu.unhideRows': { menus: 'both', edits: true },
  'outline.menu.groupRows': { menus: 'both', edits: true },
  'outline.menu.ungroupRows': { menus: 'both', edits: true },
  'outline.menu.clearRows': { menus: 'both', edits: true },
  'menu.columns': { menus: 'cell', edits: true },
  'menu.insertColumnsLeft': { menus: 'both', edits: true },
  'menu.insertColumnsRight': { menus: 'both', edits: true },
  'menu.deleteColumns': { menus: 'both', edits: true },
  'menu.columnWidth': { menus: 'both', edits: true },
  'menu.autofitColumnWidth': { menus: 'both', edits: true },
  'menu.hideColumns': { menus: 'both', edits: true },
  'menu.unhideColumns': { menus: 'both', edits: true },
  'outline.menu.groupColumns': { menus: 'both', edits: true },
  'outline.menu.ungroupColumns': { menus: 'both', edits: true },
  'outline.menu.clearColumns': { menus: 'both', edits: true },
  'outline.menu.hideDetail': { menus: 'both', edits: true },
  'outline.menu.showDetail': { menus: 'both', edits: true },
  'outline.menu.showLevel': { menus: 'both', edits: true },
  'menu.merge': { menus: 'cell', edits: true },
  'menu.unmerge': { menus: 'cell', edits: true },
  // Freezing panes is saved with the workbook.
  'menu.freeze': { menus: 'cell', edits: true },
  'menu.freezePanes': { menus: 'cell', edits: true },
  'menu.freezeTopRow': { menus: 'cell', edits: true },
  'menu.freezeFirstColumn': { menus: 'cell', edits: true },
  'menu.unfreeze': { menus: 'cell', edits: true },
  // Pivot tables, shown on one
  'pivot.menu.refresh': { menus: 'both', edits: true },
  'pivot.menu.refreshAll': { menus: 'both', edits: true },
  'pivot.menu.fieldList': { menus: 'both', edits: true },
  // Data
  'menu.sortAscending': { menus: 'both', edits: true },
  'menu.sortDescending': { menus: 'both', edits: true },
  'sort.menu.custom': { menus: 'both', edits: true },
  'menu.filter': { menus: 'both', edits: true },
  'validation.menu': { menus: 'both', edits: true },
  'cf.menu': { menus: 'both', edits: true },
  'cf.quick.greaterThan': { menus: 'both', edits: true },
  'cf.quick.lessThan': { menus: 'both', edits: true },
  'cf.quick.between': { menus: 'both', edits: true },
  'cf.quick.equal': { menus: 'both', edits: true },
  'cf.quick.textContains': { menus: 'both', edits: true },
  'cf.quick.duplicate': { menus: 'both', edits: true },
  'cf.quick.top': { menus: 'both', edits: true },
  'cf.quick.topPercent': { menus: 'both', edits: true },
  'cf.quick.bottom': { menus: 'both', edits: true },
  'cf.quick.bottomPercent': { menus: 'both', edits: true },
  'cf.quick.aboveAverage': { menus: 'both', edits: true },
  'cf.quick.belowAverage': { menus: 'both', edits: true },
  'cf.menu.gallery': { menus: 'both', edits: true },
  'cf.quick.formula': { menus: 'both', edits: true },
  'cf.menu.clearSelection': { menus: 'both', edits: true },
  'cf.menu.clearSheet': { menus: 'both', edits: true },
  'cf.menu.manage': { menus: 'both', edits: true },
  'charts.menu.insert': { menus: 'both', edits: true },
  'core.chartType.column': { menus: 'both', edits: true },
  'core.chartType.bar': { menus: 'both', edits: true },
  'core.chartType.line': { menus: 'both', edits: true },
  'core.chartType.area': { menus: 'both', edits: true },
  'core.chartType.pie': { menus: 'both', edits: true },
  'core.chartType.doughnut': { menus: 'both', edits: true },
  'core.chartType.scatter': { menus: 'both', edits: true },
  'pivot.menu.insert': { menus: 'both', edits: true },
  'menu.filterOn': { menus: 'both', edits: true },
  'menu.clearFilter': { menus: 'both', edits: true },
  // Links, in the cell menu only
  'links.menu.open': { menus: 'cell', edits: false },
  'links.menu.link': { menus: 'cell', edits: true },
  'links.menu.edit': { menus: 'cell', edits: true },
  'links.menu.remove': { menus: 'cell', edits: true },
  'links.menu.insert': { menus: 'cell', edits: true },
  // Names and finding
  'names.menu.define': { menus: 'both', edits: true },
  'names.menu.manager': { menus: 'both', edits: true },
  'menu.find': { menus: 'both', edits: false },
} as const satisfies Record<string, MenuCommandInfo>;

/** The id of a built-in entry of the cell or heading menu; see MENU_COMMANDS. */
export type MenuCommandId = keyof typeof MENU_COMMANDS;

/** A menu entry's id: a built-in one, or a plugin command's, written `pluginId/commandId`. */
export type MenuEntryId = MenuCommandId | (string & {});

/**
 * Which entries of the cell and heading menus to keep, by id, in the order
 * the editor has them; leave a list out to keep that menu whole. Keeping a
 * submenu keeps all of it; keeping some of its entries keeps the submenu
 * with just those. A plugin command is kept by `pluginId/commandId`.
 * Separators between groups left empty are dropped, and a menu left empty
 * doesn't open.
 *
 * An entry's shortcut (Ctrl+C for Copy, Shift+F2 for a note) works while
 * either menu keeps the entry, or `shortcuts` lists its id. Shortcuts with no
 * entry, such as undo or save, aren't changed.
 */
export interface MenuOptions {
  /** The cell menu's entries to keep. */
  cell?: readonly MenuEntryId[];
  /** The row and column headings' menu's entries to keep. */
  header?: readonly MenuEntryId[];
  /** Entries whose shortcuts stay on although neither menu keeps them. */
  shortcuts?: readonly MenuCommandId[];
}

/**
 * The built-in entries that have shortcuts, and where they are: in a
 * submenu (by its id) of the cell or heading menu, or (false) not in the heading menu.
 */
const SHORTCUT_ENTRIES: Partial<Record<MenuCommandId, { cell?: MenuCommandId; header?: MenuCommandId | false }>> = {
  'menu.cut': {},
  'menu.copy': {},
  'menu.paste': {},
  'menu.find': {},
  'menu.filterOn': {},
  'pivot.menu.refresh': {},
  'pivot.menu.refreshAll': {},
  'menu.insertCells': { header: false },
  'menu.deleteCells': { header: false },
  'menu.insertRowsAbove': { cell: 'menu.rows' },
  'menu.deleteRows': { cell: 'menu.rows' },
  'menu.insertColumnsLeft': { cell: 'menu.columns' },
  'menu.deleteColumns': { cell: 'menu.columns' },
  'outline.menu.groupRows': { cell: 'menu.rows' },
  'outline.menu.ungroupRows': { cell: 'menu.rows' },
  'outline.menu.groupColumns': { cell: 'menu.columns' },
  'outline.menu.ungroupColumns': { cell: 'menu.columns' },
  'menu.newNote': { header: false },
  'menu.editNote': { header: false },
  'menu.showComments': { header: false },
  'links.menu.open': { header: false },
  'core.chartType.column': { cell: 'charts.menu.insert', header: 'charts.menu.insert' },
};

/** The entries whose shortcuts `shortcutOn` answers for. */
export const SHORTCUT_COMMANDS = Object.keys(SHORTCUT_ENTRIES) as MenuCommandId[];

const keeps = (list: readonly string[] | undefined, id: string, within?: string) => !list || list.includes(id) || (!!within && list.includes(within));

/** Whether a menu entry stays in a read-only workbook: a built-in one that doesn't edit it, or a plugin's. */
export function viewOnly(id: string): boolean {
  return !(MENU_COMMANDS as Record<string, MenuCommandInfo>)[id]?.edits;
}

/** Whether the shortcut of a menu entry works with these menu options; in a read-only workbook, only those of entries that don't edit it do. */
export function shortcutOn(menus: MenuOptions | undefined, id: MenuCommandId, readOnly = false): boolean {
  if (readOnly && !viewOnly(id)) return false;
  const at = SHORTCUT_ENTRIES[id];
  if (!menus || !at) return true;
  return !!menus.shortcuts?.includes(id) || keeps(menus.cell, id, at.cell) || (at.header !== false && keeps(menus.header, id, at.header));
}

/** The entry Shift+F2 stands for at a cell: a new note, the cell's note, or its comments. */
export function noteCommand(store: WorkbookStore, row: number, col: number): MenuCommandId {
  const note = store.view?.noteAt(row, col);
  return !note ? 'menu.newNote' : note.kind === 'note' ? 'menu.editNote' : 'menu.showComments';
}
