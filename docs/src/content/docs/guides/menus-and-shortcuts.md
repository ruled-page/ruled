---
title: Menus and shortcuts
description: Keep only some entries of the cell and heading menus, by ids that are the same in every language, and the shortcuts that go with them.
sidebar:
  order: 9
---

Right-clicking a cell opens the cell menu; right-clicking a row or column heading opens the heading menu (Shift+F10 or the menu key opens the cell menu from the keyboard). Every entry has an id, the same whatever the language, the selection or the state its label shows, and the `menus` option keeps only the entries it lists:

```ts
createWorkbookEditor({
  menus: {
    cell: ['menu.copy', 'menu.paste', 'menu.pasteSpecial', 'menu.find'],
    header: ['menu.copy', 'menu.hideRows', 'menu.hideColumns', 'menu.unhideRows', 'menu.unhideColumns'],
  },
});
```

- **Leave a list out** to keep that menu whole; `[]` empties it, and an empty menu doesn't open.
- **Entries stay in the editor's order**, whatever the order of the list.
- **A submenu listed keeps all of it**: `'menu.pasteSpecial'` keeps Values, Formulas, Formats and Transposed. Listing some of its entries instead keeps the submenu with just those. The cell menu shows the Rows or Columns entries in place of their submenu when whole rows or columns are selected; listing `'menu.rows'` or `'menu.columns'` keeps them there too.
- **Separators** between groups left empty are dropped.
- **Plugin commands** are kept by `pluginId/commandId`, such as `'drive/share'`.
- **Unknown ids are ignored**, so a list written for a later version still works. In TypeScript, `MenuCommandId` lists the built-in ones.

The option can change while the editor is open, as `theme` and `language` can, or through the handle's `setMenus`.

## Shortcuts

An entry's shortcut works only while one of the menus keeps the entry: drop Copy from both menus and Ctrl+C copies nothing; drop Find and Ctrl+F is the browser's again. To keep a shortcut without its entry, list the entry's id in `shortcuts`:

```ts
menus: { cell: ['menu.copy'], header: [], shortcuts: ['menu.find'] }
```

A list left out keeps every entry of that menu, so its shortcuts too, but only for entries that menu has: notes, links, and Insert and Delete cells are in the cell menu only.

| Shortcut | Entry ids |
| --- | --- |
| Ctrl+X, Ctrl+C, Ctrl+V (⌘ on a Mac) | `menu.cut`, `menu.copy`, `menu.paste` |
| Ctrl++ and Ctrl+- | `menu.insertCells` and `menu.deleteCells`; on whole rows `menu.insertRowsAbove` and `menu.deleteRows`, on whole columns `menu.insertColumnsLeft` and `menu.deleteColumns` |
| Shift+F2 | `menu.newNote`, `menu.editNote` or `menu.showComments`, as the active cell has none, a note or comments |
| Ctrl+Enter (⌘Enter) | `links.menu.open` |
| Alt+Shift+Right | `outline.menu.groupRows` or `outline.menu.groupColumns`: on while either is |
| Alt+Shift+Left | `outline.menu.ungroupRows` or `outline.menu.ungroupColumns`: on while either is |
| Ctrl+Shift+L (⇧⌘F) | `menu.filterOn` |
| Alt+F1 | `core.chartType.column` |
| Alt+F5, Ctrl+Alt+F5 | `pivot.menu.refresh`, `pivot.menu.refreshAll` |
| Ctrl+F (⌘F) | `menu.find` |

Shortcuts with no entry, such as undo, save and moving around, aren't changed. Toolbar actions aren't menu entries either: [`features`](/guides/options-and-features/#features) turns them off, with their shortcuts where they have one (`save` takes Ctrl+S, `shortcuts` takes Ctrl+/), and the sheet tab menu goes with `sheetMenus`. The list Ctrl+/ shows leaves out the shortcuts that are off.

## Which entries edit

`MENU_COMMANDS` has every built-in id, with the menus it's in and whether choosing it can change the workbook (its cells, formats, layout or anything saved with it), for keeping only entries that read:

```ts
import { MENU_COMMANDS } from '@ruled/sheet';

const reading = Object.entries(MENU_COMMANDS).filter(([, c]) => !c.edits).map(([id]) => id);
createWorkbookEditor({ menus: { cell: reading, header: reading } });
```

Edit note and Show comments count as reading: on a note or comment that can't be changed they only show it. Plugin commands aren't in `MENU_COMMANDS`.

To stop the workbook being changed at all, not only from the menus, use [`readOnly`](/guides/options-and-features/#read-only-workbooks): it drops the same entries and their shortcuts, keeps plugins' commands, and refuses edits from everywhere else too. `menus` still applies on top of it.

## Every id

An entry's id is the message key of its label in the [English catalogue](/guides/languages/). Labels below are the English ones; *n* and *cell* are filled in from the selection. **Menus** says whether the entry is in the cell menu, the heading menu, or both.

### Clipboard and filling

Values, Formulas, Formats and Transposed are in the Paste special submenu. Fill by dragging shows on touch screens only.

| Id | Entry | Menus | Edits |
| --- | --- | --- | --- |
| `menu.cut` | Cut | both | yes |
| `menu.copy` | Copy | both | no |
| `menu.paste` | Paste | both | yes |
| `menu.pasteSpecial` | Paste special | both | yes |
| `menu.pasteValues` | Values | both | yes |
| `menu.pasteFormulas` | Formulas | both | yes |
| `menu.pasteFormats` | Formats | both | yes |
| `menu.pasteTransposed` | Transposed | both | yes |
| `menu.fillDrag` | Fill by dragging | both | yes |
| `menu.fillDown` | Fill down | both | yes |
| `menu.fillRight` | Fill right | both | yes |
| `fill.menu.series` | Fill series… | both | yes |

### Notes and comments

In the cell menu only, for the active cell: New note and New comment where it has neither, Edit note and Delete note where it has a note, Show comments and Delete comment thread where it has comments.

| Id | Entry | Menus | Edits |
| --- | --- | --- | --- |
| `menu.newNote` | New note | cell | yes |
| `menu.newComment` | New comment | cell | yes |
| `menu.editNote` | Edit note, or Show note when it can't be changed | cell | no |
| `menu.deleteNote` | Delete note | cell | yes |
| `menu.showComments` | Show comments | cell | no |
| `menu.deleteThread` | Delete comment thread | cell | yes |

### Cells, rows and columns

On a cell, the row and column entries are in the Rows and Columns submenus; on a heading, or a selection of whole rows or columns, they're the menu itself. Hide detail, Show detail and Show level *n* are in both submenus, and in either heading's menu, while the sheet has groups of that kind. Unhide shows when hidden rows or columns are next to the selection, Merge cells when more than one cell is selected, Unmerge cells when merged ones are.

| Id | Entry | Menus | Edits |
| --- | --- | --- | --- |
| `menu.insertCells` | Insert cells… | cell | yes |
| `menu.deleteCells` | Delete cells… | cell | yes |
| `menu.rows` | Rows | cell | yes |
| `menu.insertRowsAbove` | Insert *n* rows above | both | yes |
| `menu.insertRowsBelow` | Insert *n* rows below | both | yes |
| `menu.deleteRows` | Delete *n* rows | both | yes |
| `menu.rowHeight` | Row height… | both | yes |
| `menu.autofitRowHeight` | AutoFit row height | both | yes |
| `menu.hideRows` | Hide *n* rows | both | yes |
| `menu.unhideRows` | Unhide rows | both | yes |
| `outline.menu.groupRows` | Group rows | both | yes |
| `outline.menu.ungroupRows` | Ungroup rows | both | yes |
| `outline.menu.clearRows` | Clear row outline | both | yes |
| `menu.columns` | Columns | cell | yes |
| `menu.insertColumnsLeft` | Insert *n* columns left | both | yes |
| `menu.insertColumnsRight` | Insert *n* columns right | both | yes |
| `menu.deleteColumns` | Delete *n* columns | both | yes |
| `menu.columnWidth` | Column width… | both | yes |
| `menu.autofitColumnWidth` | AutoFit column width | both | yes |
| `menu.hideColumns` | Hide *n* columns | both | yes |
| `menu.unhideColumns` | Unhide columns | both | yes |
| `outline.menu.groupColumns` | Group columns | both | yes |
| `outline.menu.ungroupColumns` | Ungroup columns | both | yes |
| `outline.menu.clearColumns` | Clear column outline | both | yes |
| `outline.menu.hideDetail` | Hide detail | both | yes |
| `outline.menu.showDetail` | Show detail | both | yes |
| `outline.menu.showLevel` | Show level *n* | both | yes |
| `menu.merge` | Merge cells | cell | yes |
| `menu.unmerge` | Unmerge cells | cell | yes |

### Freezing panes

In the cell menu's Freeze panes submenu. Frozen panes are saved with the workbook, so they count as edits.

| Id | Entry | Menus | Edits |
| --- | --- | --- | --- |
| `menu.freeze` | Freeze panes | cell | yes |
| `menu.freezePanes` | Freeze rows and columns before *cell* | cell | yes |
| `menu.freezeTopRow` | Freeze top row | cell | yes |
| `menu.freezeFirstColumn` | Freeze first column | cell | yes |
| `menu.unfreeze` | Unfreeze panes | cell | yes |

### Pivot tables

First in the data group, when the selection is in a pivot table.

| Id | Entry | Menus | Edits |
| --- | --- | --- | --- |
| `pivot.menu.refresh` | Refresh | both | yes |
| `pivot.menu.refreshAll` | Refresh all | both | yes |
| `pivot.menu.fieldList` | Show Field List | both | yes |

### Data

The `cf.` entries are in the Conditional formatting submenu, the chart types in Insert chart (shown while the `charts` feature is on). Show all rows shows while a filter hides some.

| Id | Entry | Menus | Edits |
| --- | --- | --- | --- |
| `menu.sortAscending` | Sort A to Z | both | yes |
| `menu.sortDescending` | Sort Z to A | both | yes |
| `sort.menu.custom` | Custom sort… | both | yes |
| `menu.filter` | Filter by values… | both | yes |
| `validation.menu` | Data validation… | both | yes |
| `cf.menu` | Conditional formatting | both | yes |
| `cf.quick.greaterThan` | Greater than… | both | yes |
| `cf.quick.lessThan` | Less than… | both | yes |
| `cf.quick.between` | Between… | both | yes |
| `cf.quick.equal` | Equal to… | both | yes |
| `cf.quick.textContains` | Text that contains… | both | yes |
| `cf.quick.duplicate` | Duplicate values… | both | yes |
| `cf.quick.top` | Top 10 items… | both | yes |
| `cf.quick.topPercent` | Top 10%… | both | yes |
| `cf.quick.bottom` | Bottom 10 items… | both | yes |
| `cf.quick.bottomPercent` | Bottom 10%… | both | yes |
| `cf.quick.aboveAverage` | Above average… | both | yes |
| `cf.quick.belowAverage` | Below average… | both | yes |
| `cf.menu.gallery` | Data bars, color scales and icon sets… | both | yes |
| `cf.quick.formula` | New rule using a formula… | both | yes |
| `cf.menu.clearSelection` | Clear rules from selected cells | both | yes |
| `cf.menu.clearSheet` | Clear rules from entire sheet | both | yes |
| `cf.menu.manage` | Manage rules… | both | yes |
| `charts.menu.insert` | Insert chart | both | yes |
| `core.chartType.column` | Column | both | yes |
| `core.chartType.bar` | Bar | both | yes |
| `core.chartType.line` | Line | both | yes |
| `core.chartType.area` | Area | both | yes |
| `core.chartType.pie` | Pie | both | yes |
| `core.chartType.doughnut` | Doughnut | both | yes |
| `core.chartType.scatter` | Scatter | both | yes |
| `pivot.menu.insert` | Insert PivotTable… | both | yes |
| `menu.filterOn` | Add filter buttons, or Remove filter buttons | both | yes |
| `menu.clearFilter` | Show all rows (clear filter) | both | yes |

### Links

In the cell menu only. Open link shows on a cell with a link while links are followed; Edit link and Remove links are in the Link submenu, shown when the selection has links, and Insert link otherwise.

| Id | Entry | Menus | Edits |
| --- | --- | --- | --- |
| `links.menu.open` | Open link | cell | no |
| `links.menu.link` | Link | cell | yes |
| `links.menu.edit` | Edit link… | cell | yes |
| `links.menu.remove` | Remove links | cell | yes |
| `links.menu.insert` | Insert link… | cell | yes |

### Names and finding

| Id | Entry | Menus | Edits |
| --- | --- | --- | --- |
| `names.menu.define` | Define name… | both | yes |
| `names.menu.manager` | Name manager… | both | yes |
| `menu.find` | Find and replace… | both | no |
