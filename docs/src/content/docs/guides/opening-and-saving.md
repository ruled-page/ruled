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

Throwing from `onSave` shows the error and keeps the changes, so people can try again. While the promise is pending, the editor shows the save as under way.

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
import { SaveConflictError } from '@sitepen/spreadsheet';

onSave: async (saved) => {
  const response = await api.put(id, saved.bytes, { ifMatch: etag });
  if (response.status === 412) throw new SaveConflictError();
},
```

## What's kept in the file

The editor writes back only what changed. Everything it doesn't show (pictures, slicers, macros, styles it doesn't use, and parts it doesn't understand) is kept as it was, and a workbook nobody changed is saved byte for byte. See [Excel fidelity](/guides/excel-fidelity/).

## Without saving

- `getWorkbook()` gives you the workbook as a save would, without its checks: for autosave, or to read values as people work.
- `downloadCopy()` downloads it, as an export, without saving it to your app. With `onSave`, the toolbar has a download button for this too, unless `features.download` is off.
- `getCsv()` and `downloadCsv()` give the sheet shown as CSV. Values that would run as formulas in another spreadsheet are escaped unless `escapeFormulas` is `false`.

## Drafts

Unsaved edits are kept in the browser, and offered back when the same file is opened again, so a closed tab or a crash doesn't lose work. Turn it off with `drafts: false`. `warnOnUnload` (on by default) asks before people leave the page with unsaved changes.
