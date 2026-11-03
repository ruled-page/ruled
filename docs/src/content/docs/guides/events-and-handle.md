---
title: Events and the handle
description: Control the editor from your app, follow what it does, and hear about errors.
sidebar:
  order: 5
---

Every way of showing the editor gives you the same handle: `createWorkbookEditor()` returns it, `mountWorkbookEditor()` returns it, and the `<workbook-editor>` element is one.

## Methods

| Method | What it does |
| --- | --- |
| `open(bytes, fileName)` | Opens a workbook or CSV file. Resolves `true` once open, `false` if the user kept the one already open. |
| `save()` | Runs the Save button's steps, resolving `'saved'`, `'downloaded'`, `'cancelled'`, `'blocked'` or `'failed'`. |
| `getWorkbook()` | The workbook as a save would hand it over, without the save's checks. |
| `downloadCopy()` | Downloads the workbook as `.xlsx` without saving it to your app. |
| `getCsv(options?)`, `downloadCsv(options?)` | The sheet shown, as CSV. |
| `goTo(ref, { focus }?)` | Selects and shows a cell, a range or a name: `'B4'`, `'Prices!A1:C3'`. The grid takes the keyboard focus, as after the name box, unless `focus` is `false`. |
| `ready` | A promise (a property, not a method) that resolves once every plugin's `setup` has run and the editor can draw what they added; at once when there are no plugins. It never rejects: a plugin that throws is reported as a `pluginError` event and the rest still count. If the editor is destroyed first it resolves with the plugins not started. See [when plugins start](/plugins/writing/#when-setup-runs). |
| `scrollTo(ref)` | Scrolls a cell to the top left of the part of the sheet that scrolls, below and right of any frozen panes, keeping the selection and the focus. `false` when `ref` names no cell on a sheet shown. |
| `freeze({ rows, cols, sheet? })` | Freezes the top `rows` and left `cols` of the active sheet, or of the sheet named by `sheet`, for viewing only. It is not a change to the file: no undo step, the workbook isn't marked unsaved, and saving writes the file's own panes. `{ rows: 0, cols: 0 }` shows the file's own panes again. It lasts until the file is opened again; freezing or unfreezing from the menu changes the file as usual and replaces it on that sheet. `viewport` follows it. `false` when no workbook is open or `sheet` names none. |
| `discardDraft()` | Deletes the [draft](/guides/opening-and-saving/#drafts) kept in the browser for the open file, and changes nothing else: edits made afterwards are drafted as usual. |
| `registerSchema(schema)`, `removeSchema(id)` | Add or remove a schema while the editor is open. |
| `setTheme(theme)`, `setLanguage(language)`, `setUser(user)`, `setCustomLists(lists)`, `setMenus(menus)` | Change these without closing the workbook. |
| `subscribe(listener)` | Calls `listener` when the properties below change, `selection` and `viewport` included. Returns a function that stops it. |
| `destroy()` | Stops the editor and its worker. |

## Properties

| Property | What it is |
| --- | --- |
| `status` | `'empty'` before a workbook is open, `'loading'` while one opens, `'ready'`, or `'error'` when it stopped working. |
| `fileName` | The open workbook's name. |
| `editedCount` | How many cells differ from the file as opened. |
| `unsaved` | Whether there are edits your app hasn't received through a save. |
| `selection` | `{ sheet, active, ranges }`: the sheet shown, its active cell and its selected ranges in A1 style, such as `{ sheet: 'Prices', active: 'C4', ranges: ['B2:C4'] }`. |
| `viewport` | `{ sheet, row, col, cell }`: the first row and column in view below and right of any frozen panes, wholly or in part, numbered as in A1 (row 1 is 1, column A is 1), and that cell, such as `'C10'`. |

`selection` and `viewport` stay the same object until they change, so a listener can tell what changed by comparing them with the last it saw. While the sheet scrolls, `viewport` changes at most once a frame. Rows of their own height, hidden rows and collapsed groups are all taken into account: a viewport row is always one shown.

## Following another editor

Two editors side by side can follow each other with `selection`, `viewport`, `goTo(ref, { focus: false })` and `scrollTo(ref)`. Leave alone what already matches, so the two don't keep answering each other:

```ts
const quote = (sheet: string) => `'${sheet.replace(/'/g, "''")}'`;

function follow(from: EditorHandle, to: EditorHandle) {
  let selection = from.selection;
  let viewport = from.viewport;
  return from.subscribe(() => {
    if (from.selection && from.selection !== selection) {
      selection = from.selection;
      const now = to.selection;
      if (now?.sheet !== selection.sheet || now.ranges.join() !== selection.ranges.join()) to.goTo(`${quote(selection.sheet)}!${selection.ranges[0]}`, { focus: false });
    }
    if (from.viewport && from.viewport !== viewport) {
      viewport = from.viewport;
      const now = to.viewport;
      if (now?.sheet !== viewport.sheet || now.cell !== viewport.cell) to.scrollTo(`${quote(viewport.sheet)}!${viewport.cell}`);
    }
  });
}

follow(left, right);
follow(right, left);
```

The keyboard focus stays in the editor being used, so typing carries on there.

## Following the editor

`onEvent` hears what the editor does, with timings, for your logs or analytics. Its messages are in English whatever the editor's language, so logs read the same:

```ts
createWorkbookEditor({
  onEvent: (event) => {
    if (event.type === 'opened') metrics.timing('editor.open', event.ms, { sheets: event.sheets });
    if (event.type === 'saved') metrics.count('editor.save', { result: event.result, cells: event.editedCount });
    if (event.type === 'openFailed') log.warn('Could not open', event.fileName, event.reason);
  },
});
```

| Event | When |
| --- | --- |
| `opened` | A workbook opened: its name, how long it took and how many sheets it has. |
| `openFailed` | A workbook couldn't be opened, and why. |
| `calculated` | The formulas are ready: how long that took and how many are kept up to date. |
| `saved` | A save finished: what happened, how long it took, and how many cells had changed. |
| `error` | An error the editor showed the user. |

## When the editor stops

The workbook lives in a worker. If the worker can't load, crashes, runs out of memory or stops answering, the editor says so and offers to reopen the file (bringing back unsaved edits when drafts are on), and calls `onError` so your app can report it.

## The element's events

The `<workbook-editor>` element also fires DOM events: `ready`, `opened`, `change`, `save`, `editorerror`, `selectionchange` and `viewportchange`. See [Web component and plain HTML](/start/web-component/#events).
