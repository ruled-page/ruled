---
title: Options and features
description: Every option the editor takes, and the parts of its interface you can turn off.
sidebar:
  order: 2
---

The same options go to `useWorkbookEditor`, `createWorkbookEditor`, `mountWorkbookEditor` and the element's `options` property. All are optional. The [API reference](/reference/api/ruled/sheet/interfaces/editoroptions/) has each one's type.

## Workbooks and saving

| Option | What it does |
| --- | --- |
| `schemas` | [Schemas](/schemas/overview/) to try, in order, against each workbook opened. |
| `onSave` | Receives the workbook when people save. Without it, saving downloads the file. |
| `plugins` | [Plugins](/plugins/using/): places to open and save files, such as OneDrive, and buttons, panels and dialogs of their own. Read when the editor is made. |
| `saveLabel` | The Save button's label, such as "Submit price list", in one language or several. |
| `drafts` | Keep unsaved edits in the browser and offer them back; never for a password-protected workbook. Default `true`. |
| `warnOnUnload` | Ask before leaving the page with unsaved changes. Default `true`. |
| `numbers` | How numbers are shown, copied and pasted, for every workbook; a schema can override it. |
| `csvExport` | The delimiter for CSV downloads, and whether formulas are escaped. |
| `limits` | Size and complexity limits for files people open. See [Large workbooks](/guides/opening-and-saving/#large-workbooks). |
| `macros` | Files with macros: `'keep'` them (the default), `'remove'` them when saving, or `'block'` such files. |
| `activeContent` | ActiveX controls and embedded objects: `'keep'` (the default) or `'block'`. |
| `readOnly` | Shows the workbook without letting it be changed. See [Read-only workbooks](#read-only-workbooks). Can change while the editor is open. |
| `recalc` | Live recalculation: `'auto'` (the default) turns it off for workbooks with more than 200,000 formulas or 2 million cells, where people can turn it on; `'on'` and `'off'` force it. |

## People and language

| Option | What it does |
| --- | --- |
| `user` | The name recorded against each change in the history, and the author of new notes and comments. |
| `language` | The language of the editor's own text, such as `'de'`. Defaults to the locale. |
| `locale` | How dates and numbers are read and shown. Defaults to the browser's. |
| `translations` | More languages, or messages worded your way. See [Languages](/guides/languages/). |
| `customLists` | Lists that dragging the fill handle continues, such as `[['North', 'East', 'South', 'West']]`. |

## Look

| Option | What it does |
| --- | --- |
| `theme` | Theme tokens: colours, fonts, sizes, corners and shadows. See [Theming](/guides/theming/). |
| `icons` | Replacement icons by name: SVG path data on a 24×24 grid, or, in React, an element. |
| `features` | Parts of the interface to show (below). |
| `menus` | Which entries of the cell and heading menus to keep, by id, and with them their shortcuts. Can change while the editor is open. See [Menus and shortcuts](/guides/menus-and-shortcuts/). |

## Links

| Option | What it does |
| --- | --- |
| `links` | `'follow'` (the default): Ctrl+click (⌘-click on a Mac) follows a hyperlink, asking first the first time for each address. `'off'`: links show and can be edited, but are never followed. |
| `onLinkClick` | Called instead when people follow a link to a web page or an email. Return `true` to have the editor open it after all. |

## Running it

| Option | What it does |
| --- | --- |
| `createWorker` | Creates the worker that holds the workbook, for bundlers that don't bundle it. |
| `workerTimeout` | How long the worker may spend on one task before it's taken to have stopped. Default 120,000 ms. |
| `onError` | Called when the editor stops working: its worker couldn't load, crashed or stopped answering. The editor shows what happened and offers to reopen the file. |
| `onEvent` | Hears what the editor does and how long it took, for your logging. See [Events and the handle](/guides/events-and-handle/). |
| `diagnostics` | Shows in the status bar how long opening, calculating and the last edit took. |

## Features

`features` turns parts of the interface off; each is on unless set to `false`.

| Feature | What it is |
| --- | --- |
| `open` | The Open button (or menu, with [plugins](/plugins/using/)), the file picker and dropping files. Turn it off when your app chooses the file. |
| `save` | The Save button and Ctrl+S (⌘S). |
| `download` | With `onSave`, a separate button to download a copy. |
| `schemas` | The schema designer's button. |
| `report` | "What's in this file": what the workbook holds that the editor keeps but doesn't show, and how its formulas are calculated. Useful to people who know Excel; turn it off for people who only fill in a template. |
| `charts` | Charts drawn from the workbook's values. Off, they show as placeholders with their titles; the chart library loads only when a workbook has a chart. |
| `format` | The format bar: fonts, colours, borders, alignment and number formats. Turn it off where people only fill in values. |
| `csv` | Opening CSV files, and a button to download the sheet as CSV. |
| `autoComplete` | Finishing text as it's typed from the entries above and below, as Excel's AutoComplete does. |
| `issuesPanel` | The schema issues panel (**Checks**) and the schema chip's button that opens it. Off, the panel never opens, by itself or from anything in the editor, and the chip only shows the status. The checks still run: cells are marked, saves are blocked or confirmed as the schema says, and each save's `schema` tells your app the counts of errors and warnings. The panel's list itself isn't exposed to the host; the counts and the cell markers are. |
| `sourceSave` | Writing to [plugins' file sources](/plugins/using/#saving-with-a-source): **Save to *source*** and the sources in **Save a copy**. Turn it off where files must only go to your app. |
| `titleBar` | The title bar at the start of the toolbar: the editor's name and mark, the file's name, the count of changes, the schema's chip and the working spinner. |
| `statusBar` | The status bar: live recalculation, circular references, filtering, and the count, sum and average of the selection. |
| `sheetTabButtons` | The buttons beside the sheet tabs: **New sheet**, and the shown sheet's menu. |
| `sheetMenus` | Changing sheets from their tabs: the tab menu (insert, rename, delete, move, colour, hide and unhide), renaming by double-click or F2, moving by dragging or Alt+Shift+arrows, and unhiding from the hidden-sheets note. Off, clicking a tab still shows its sheet, and the arrow keys, Home and End still move between the tabs. |
| `lockNotes` | The note under the formula bar that the selected cell is locked or its sheet protected. Off, the cells stay read-only; the note just isn't shown. |
| `undoButtons` | The toolbar's Undo and Redo buttons. Ctrl+Z and Ctrl+Y (⌘Z and ⇧⌘Z) work either way. |
| `recordButtons` | The toolbar's **Insert row** and **Delete row** buttons (or column, for records across) for a [schema's](/schemas/overview/) tables that can grow. |
| `exactNumbers` | The toolbar's **Exact numbers** switch. |
| `highlightEdits` | The toolbar's **Highlight changes** switch. |
| `schemaLabels` | The toolbar's **Mark schema labels** switch, shown while a schema matches. |
| `shortcuts` | The list of keyboard shortcuts: its toolbar button and Ctrl+/ (⌘/). |

Every toolbar action has a flag: **Open** is `open`; **Save** and **Save a copy** are `save`; downloading a copy is `download`, and as CSV `csv`; the phone's **Format bar** switch is `format`; **Schemas** is `schemas`; and **What's in this file** is `report`. Buttons that plugins add are theirs; leave out the plugin, or its command, to drop them. The flags apply in the phone layout too, where the less used actions are behind **More**; with all of those off, **More** goes as well.

`features` and [`menus`](/guides/menus-and-shortcuts/) don't overlap. `menus` keeps entries of the cell and heading menus, and their shortcuts; the toolbar's actions and their shortcuts, such as Save with Ctrl+S, undo with Ctrl+Z and the shortcuts list with Ctrl+/, aren't menu entries, so only `features` turns them off. The sheet tab menu isn't one of the menus `menus` trims; `sheetMenus` turns it off whole.

`features` is read when the editor is made; to change it, make the editor again.

## Read-only workbooks

`readOnly: true` shows a workbook as it is, for people who should look at a file but not change it, such as a customer's original. Every edit is refused: typing, pasting, cutting and moving cut cells, filling, dragging cells, inserting and deleting rows, columns and cells, sizes, hiding, grouping, freezing panes, merging, formats, notes and comments, links, sorting, filters, data validation, conditional formats, names, sheets (adding, renaming, deleting, moving, hiding and unhiding), charts, pivot tables, and undoing or redoing earlier edits.

What would make them goes too: the cell and heading menus keep only the entries that don't edit (those whose [`edits`](/guides/menus-and-shortcuts/) is `false`, and plugins' commands), and the shortcuts of the others are off whatever `menus` keeps. The format bar, the schema designer, the buttons beside the sheet tabs, the tab menu, Undo and Redo, and the record buttons are hidden; Find has no Replace; the fill handle and the selection's border can't be dragged, column and row edges can't be resized, and charts, pictures and shapes can't be moved, deleted or edited. These come back when `readOnly` is turned off; `features` itself keeps what you passed.

Selecting, copying, finding, following links, moving between sheets, reading notes and comments, and the switches that only change what is shown (**Exact numbers**, **Highlight changes**) still work. Opening and saving are as `features` has them. A refused edit isn't shown on screen, which would get in the way at every key pressed; screen readers are told the workbook is read-only, at most every few seconds.

Nothing is marked unsaved and nothing is written to [drafts](/guides/opening-and-saving/#drafts), so saving gives back the workbook as it was opened (or, if it was edited before `readOnly` was turned on, as it was then). A draft left from editing the same file before isn't offered back while it is read-only, and is kept for when it isn't.

`readOnly` can change while the editor is open: from `options` with React's hook, `setReadOnly(readOnly)` on the [handle](/guides/events-and-handle/), or the `read-only` attribute or `readOnly` property of [`<workbook-editor>`](/start/web-component/). Turning it on closes any menu or dialog that edits, and an edit being typed is refused when it is entered.

A supplier filling in a template might get:

```ts
createWorkbookEditor({
  schemas: [priceList],
  features: { open: false, schemas: false, report: false, format: false },
  saveLabel: 'Submit price list',
  user: supplier.name,
});
```
