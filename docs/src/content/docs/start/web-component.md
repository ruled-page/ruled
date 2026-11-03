---
title: Web component and plain HTML
description: Use the <workbook-editor> element on any page, or mount the editor into an element you choose.
sidebar:
  order: 6
---

The standalone build defines a `<workbook-editor>` element. Set its `options`, give it a workbook, and listen to its events.

```html
<workbook-editor src="/files/forecast.xlsx" save-handler="event"></workbook-editor>

<script type="module">
  import { defineSchema } from '@ruled/sheet/standalone';

  const editor = document.querySelector('workbook-editor');
  editor.options = {
    features: { open: false },
    saveLabel: 'Save to the app',
    schemas: [defineSchema({ id: 'forecast', name: 'Forecast', sheets: [/* … */] })],
  };

  editor.addEventListener('save', (event) => {
    const saved = event.detail;
    // Keep the save pending until the upload finishes; a rejection is shown to the user.
    saved.waitUntil(fetch('/api/forecasts', { method: 'POST', body: saved.bytes }));
  });
</script>

<style>
  workbook-editor { display: block; height: 80vh; }
</style>
```

## Attributes

| Attribute | What it does |
| --- | --- |
| `src` | A workbook's URL. The element fetches and opens it, and opens it again when it changes. |
| `save-handler="event"` | Saving fires a `save` event with the workbook, instead of downloading the file. |

## Properties and methods

The element has the same methods and properties as the [handle](/guides/events-and-handle/): `open(bytes, fileName)`, `save()`, `downloadCopy()`, `getWorkbook()`, `getCsv()`, `downloadCsv()`, `goTo(ref, { focus }?)`, `discardDraft()`, `registerSchema()`, `removeSchema()`, `setTheme()`, `setLanguage()`, `setUser()`, `setCustomLists()`, `setMenus()`, and `status`, `fileName`, `unsaved`, `editedCount`, `selection` and `viewport`. `whenReady()` returns the handle's `ready` promise, resolving once the plugins are set up and the editor can draw what they add (waiting for the element to be connected first); an element that re-creates its editor gives the new one's. The handle's `scrollTo(ref)` is `scrollToCell(ref)` on the element, as `scrollTo` already scrolls the element itself.

Setting `options` on a connected element applies a new theme, language, user, schemas or custom lists as they are; changing any other option re-creates the editor, closing the workbook.

## Events

| Event | `detail` |
| --- | --- |
| `ready` | The editor is made and can take a workbook. Its plugins may not be set up yet: `await el.whenReady()` for that. |
| `opened` | `{ fileName }`, once a workbook is open. |
| `change` | `{ editedCount, unsaved }`, as the edits change. |
| `save` | The saved workbook, with `waitUntil(promise)`, when `save-handler="event"`. |
| `editorerror` | `{ message }`, when the workbook from `src` couldn't be opened or the editor stopped working. |
| `selectionchange` | The new `selection`, when the selection changes or another sheet is shown. |
| `viewportchange` | The new `viewport`, when the sheet scrolls or another is shown; at most once a frame. |

Moving the element elsewhere in the page keeps the editor as it is.

## Mounting into your own element

Without the element, mount the editor into any element and keep the handle:

```ts
import { mountWorkbookEditor } from '@ruled/sheet/standalone';

const editor = mountWorkbookEditor(document.getElementById('sheet')!, { schemas: [forecast] });
await editor.open(await (await fetch('/files/forecast.xlsx')).arrayBuffer(), 'forecast.xlsx');
```
