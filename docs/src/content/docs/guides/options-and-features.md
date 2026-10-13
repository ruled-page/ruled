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
| `saveLabel` | The Save button's label, such as "Submit price list", in one language or several. |
| `drafts` | Keep unsaved edits in the browser and offer them back. Default `true`. |
| `warnOnUnload` | Ask before leaving the page with unsaved changes. Default `true`. |
| `numbers` | How numbers are shown, copied and pasted, for every workbook; a schema can override it. |
| `csvExport` | The delimiter for CSV downloads, and whether formulas are escaped. |
| `limits` | Size and complexity limits for files people open. |
| `macros` | Files with macros: `'keep'` them (the default), `'remove'` them when saving, or `'block'` such files. |
| `activeContent` | ActiveX controls and embedded objects: `'keep'` (the default) or `'block'`. |
| `recalc` | Live recalculation: `'auto'` (the default) turns it off for workbooks too large to recalculate quickly, where people can turn it on; `'on'` and `'off'` force it. |

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
| `open` | The Open button, the file picker and dropping files. Turn it off when your app chooses the file. |
| `save` | The Save button and Ctrl+S (⌘S). |
| `download` | With `onSave`, a separate button to download a copy. |
| `schemas` | The schema designer's button. |
| `report` | "What's in this file": what the workbook holds that the editor keeps but doesn't show, and how its formulas are calculated. Useful to people who know Excel; turn it off for people who only fill in a template. |
| `charts` | Charts drawn from the workbook's values. Off, they show as placeholders with their titles; the chart library loads only when a workbook has a chart. |
| `format` | The format bar: fonts, colours, borders, alignment and number formats. Turn it off where people only fill in values. |
| `csv` | Opening CSV files, and a button to download the sheet as CSV. |
| `autoComplete` | Finishing text as it's typed from the entries above and below, as Excel's AutoComplete does. |

A supplier filling in a template might get:

```ts
createWorkbookEditor({
  schemas: [priceList],
  features: { open: false, schemas: false, report: false, format: false },
  saveLabel: 'Submit price list',
  user: supplier.name,
});
```
