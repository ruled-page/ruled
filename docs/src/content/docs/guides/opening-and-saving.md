---
title: Opening and saving
description: Give the editor a workbook, receive it back with its values and changes, and handle conflicts, drafts and CSV.
sidebar:
  order: 1
---

## Opening a workbook

Hand the editor the file's bytes and its name:

```ts
const response = await fetch('/api/files/42');
const opened = await editor.open(await response.arrayBuffer(), 'price-list.xlsx');
```

`open` resolves `true` once the workbook is open, and `false` if the user chose to keep the workbook already open, with its unsaved changes. It rejects, saying why, when the file can't be opened, and the editor shows the reason too. The editor copies the bytes, so yours stay usable.

The element can fetch the file itself: `<workbook-editor src="/api/files/42">`.

`.xlsx` and `.xlsm` open as workbooks, and `.csv`, `.tsv` and `.txt` as a workbook of one sheet. People can also open files themselves, with the Open button or by dropping one on the editor, unless you set `features: { open: false }`.

## Large workbooks

The workbook is read, and its formulas worked out, in a worker. Before anything is unzipped, and then as the cells are read, the file is checked against `limits`, and refused, saying why, when it goes over one:

| Limit | Default | What it checks |
| --- | --- | --- |
| `maxFileBytes` | 50 MB | The file's size. |
| `maxUnzippedBytes` | 500 MB | The total size of its parts once unzipped. |
| `maxRatio` | 200 | How many times its zipped size a part that unzips to more than 10 MB may grow, as a zip bomb's would. |
| `maxParts` | 20,000 | How many parts the file has. |
| `maxCells` | 2,000,000 | The cells on every sheet, together. |

With `recalc: 'auto'`, the default, live recalculation starts off for a workbook with more than 200,000 formulas or 2 million cells. Its formulas show the values Excel saved, and a banner says so and lets people turn it on. Excel works them out again when it opens the saved file.

Two generated workbooks of 12 columns, three of them formulas, took this on an Apple M3 Max, running the worker's code in Node:

| Workbook | Live recalculation | Opening | Calculating | Saving an edit | Memory held | Most memory used |
| --- | --- | --- | --- | --- | --- | --- |
| 1.1 million cells, 300,000 formulas, 10 MB | Off | 1.0 s | | 1.4 s | 215 MB | 530 MB |
| | On | 1.0 s | 2.4 s | 1.5 s | 375 MB | 640 MB |
| 1.9 million cells, 525,000 formulas, 18 MB | Off | 1.8 s | | 2.5 s | 385 MB | 800 MB |
| | On | 1.8 s | 4.5 s | 3.0 s | 670 MB | 1.2 GB |

Memory held is what the worker keeps while the workbook is open, and the most used is while it opens, as the file is unzipped and read. The page keeps a copy of the cells to draw them, about 55 and 95 MB more here. Memory grows with the number of cells: about 200 MB a million with live recalculation off, and 340 MB with it on.

Phones give a page less memory than computers do. If people will open large workbooks on phones, lower `maxCells` so that a workbook too large for the phone is refused with a message, rather than its worker running out of memory. When that happens, the editor says so and offers to open the file again (see [Events and the handle](/guides/events-and-handle/)).

## Receiving the saved workbook

Without `onSave`, saving downloads the file. With it, the editor hands the workbook to you:

```ts
createWorkbookEditor({
  saveLabel: 'Submit',
  onSave: async (saved) => {
    await api.upload(saved.fileName, saved.bytes);
    await api.storeLines(saved.values?.Prices.tables.prices ?? []);
  },
});
```

Throwing from `onSave` shows the error and keeps the changes, so people can try again. So does answering `{ error: 'message' }` instead, which a [Server Action](/start/nextjs/#saving-with-a-server-action) has to do, as frameworks hide the errors it throws. While the promise is pending, the editor shows the save as under way.

What you receive:

| Property | What it is |
| --- | --- |
| `bytes` | The `.xlsx` file. |
| `fileName` | Its name. |
| `values` | With a schema, the values in its shape: for each sheet, each table's records keyed by field, and each single cell. Dates are ISO text. |
| `schema` | With a schema, its `id`, `name` and `version`, and how many `errors` and `warnings` remained. |
| `editedCount`, `changes` | How many cells changed, and each one, with its value when opened and now. |
| `history` | What was done, by whom and when. |
| `outOfDate` | Formulas whose saved values may be stale: ones the editor can't calculate whose inputs changed. Excel recalculates them on opening. |
| `cell(sheet, ref)`, `range(sheet, ref?)` | Any cell's or range's values, as JSON. |

## What saving does first

1. **Waits for checks still running**, including async ones.
2. **Refuses to save while errors remain**, or asks first under `saveWithErrors: 'confirm'`.
3. **Runs `beforeSave`**, which can stop it with a message.
4. **Shows every changed cell** for confirmation, if the schema's `numbers.reviewBeforeSave` is on.
5. **Writes the file** and hands it to `onSave`, or downloads it.

`save()` on the handle runs the same steps and resolves with what happened: `'saved'`, `'downloaded'`, `'cancelled'`, `'blocked'` or `'failed'`.

## Someone else saved first

When the stored file changed since it was opened, throw `SaveConflictError` from `onSave`. The editor keeps the user's changes and offers to download them as a copy, explaining why in its language:

```ts
import { SaveConflictError } from '@ruled/sheet';

onSave: async (saved) => {
  const response = await api.put(id, saved.bytes, { ifMatch: etag });
  if (response.status === 412) throw new SaveConflictError();
},
```

Answering `{ conflict: true }` does the same, and `{ conflict: 'Ana saved it at 10:02.' }` says why in your words.

## What's kept in the file

The editor writes back only what changed. Everything it doesn't show (shapes, slicers, macros, styles it doesn't use, and parts it doesn't understand) is kept as it was, and a workbook nobody changed is saved byte for byte. See [Excel fidelity](/guides/excel-fidelity/).

## Without saving

- `getWorkbook()` gives you the workbook as a save would, without its checks: for autosave, or to read values as people work.
- `downloadCopy()` downloads it, as an export, without saving it to your app. With `onSave`, the toolbar has a download button for this too, unless `features.download` is off.
- `getCsv()` and `downloadCsv()` give the sheet shown as CSV. Values that would run as formulas in another spreadsheet are escaped unless `escapeFormulas` is `false`.

## Drafts

Unsaved edits are kept in the browser, and offered back when the same file is opened again, so a closed tab or a crash doesn't lose work. Turn it off with `drafts: false`. `warnOnUnload` (on by default) asks before people leave the page with unsaved changes.
