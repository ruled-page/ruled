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
| `goTo(ref)` | Selects and shows a cell, a range or a name: `'B4'`, `'Prices!A1:C3'`. |
| `registerSchema(schema)`, `removeSchema(id)` | Add or remove a schema while the editor is open. |
| `setTheme(theme)`, `setLanguage(language)`, `setUser(user)`, `setCustomLists(lists)` | Change these without closing the workbook. |
| `subscribe(listener)` | Calls `listener` when the properties below change. Returns a function that stops it. |
| `destroy()` | Stops the editor and its worker. |

## Properties

| Property | What it is |
| --- | --- |
| `status` | `'empty'` before a workbook is open, `'loading'` while one opens, `'ready'`, or `'error'` when it stopped working. |
| `fileName` | The open workbook's name. |
| `editedCount` | How many cells differ from the file as opened. |
| `unsaved` | Whether there are edits your app hasn't received through a save. |

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

The `<workbook-editor>` element also fires DOM events: `ready`, `opened`, `change`, `save` and `editorerror`. See [Web component and plain HTML](/start/web-component/#events).
