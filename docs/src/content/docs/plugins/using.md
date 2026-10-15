---
title: Using plugins
description: Add plugins to the editor, how their saves sit beside your app's own, the Content Security Policy they need, and why they're trusted code.
sidebar:
  order: 1
---

A plugin adds to the editor: a place to open workbooks from and save them to (OneDrive, Google Drive), buttons on the toolbar or in More, entries in the cell menu, a side panel, or dialogs. Plugins are plain objects, so the same plugin works with the React component and the `<workbook-editor>` element.

## Passing them in

Give the editor a list of plugins in `plugins`:

```tsx
import { useWorkbookEditor, WorkbookEditor } from '@ruled/sheet';
import { oneDrive } from '@ruled/plugin/onedrive';

function Prices() {
  const editor = useWorkbookEditor({
    plugins: [oneDrive({ clientId: '11111111-2222-3333-4444-555555555555' })],
  });
  return <WorkbookEditor editor={editor} />;
}
```

Plugins are objects, not text, so the element can't take them as attributes. Set them in its `options` property instead:

```js
import { oneDrive } from '@ruled/plugin/onedrive';

document.querySelector('workbook-editor').options = {
  plugins: [oneDrive({ clientId: '11111111-2222-3333-4444-555555555555' })],
};
```

`plugins` is read when the editor is made, like most options. To change it, make a new editor: a new `key` in React, or a new element. Setting the element's `options` again with different plugins re-creates the editor.

The code that runs plugins is loaded only for an editor given some, so an editor without plugins is no larger. Each plugin starts a moment after the editor is made, once that code has loaded.

## What a plugin can add

| Where | What |
| --- | --- |
| **Open** | With a plugin's file source, the Open button becomes a menu: **This device**, then each source. A source that names files, such as recent ones, has them in a submenu. `features.open: false` hides the menu as it hides the button. |
| **Save to *source*** | Writes a workbook opened from a source back over its file. |
| **Save a copy** | A menu with **This device** (a download) and each source that can save new files. |
| **Toolbar and More** | Buttons. On phones, the ones meant for More sit behind it. |
| **Cell menu** | Entries after the editor's own, for the cells selected. |
| **Panels** | A side panel beside Issues and Changes, with a button that shows and hides it. |

## Saving, with a source

A source adds two ways to save beside your app's own. Each keeps its own meaning:

| Action | What it does | Shown when |
| --- | --- | --- |
| **Save**, your app's | Runs the schema's checks, then your `onSave`, as without plugins. Labelled by `saveLabel`. | You give `onSave`. |
| **Save to *source***, such as "Save to OneDrive" | Writes the workbook back over the file it was opened from, unless the file changed since it was opened. | The workbook came from a source that can save, and the file can be written. |
| **Save a copy** | Downloads a copy, or writes a new file to a source. | Downloads are on, or a source can save new files. |

The main Save button, and Ctrl+S (⌘S), is your app's Save when there is one. Without `onSave`, it is **Save to *source*** for a workbook from a source, and otherwise a download, as before. The others sit beside it.

Saving to a source doesn't wait for the schema's errors to be fixed, as a download doesn't: the file is the person's own. The editor tells them how many errors there are and lets them carry on. A file changed by someone else since it was opened is handled as a [save conflict](/guides/opening-and-saving/#someone-else-saved-first) from `onSave` is.

The editor keeps where the workbook came from, and tells your app:

- `editor.origin` is the open workbook's [`FileOrigin`](/reference/api/ruled/sheet/plugin/interfaces/fileorigin/): the source's id, the file's name, the source's reference to it, and its version. It is `undefined` for files from your app or from this device.
- `onSave` receives it as `origin` on the saved workbook, in case your app records where a workbook came from.
- `onEvent` hears `{ type: 'savedToSource', source, name, ms }` for each write to a source.

To keep files going only to your app, set `features.sourceSave: false`. That removes **Save to *source*** and the sources in **Save a copy**. With `features.download: false` as well, only your app's Save is left.

## Their text

A plugin's messages are kept under `plugin.<id>.`, so they can't collide with the editor's or another plugin's. Your [`translations`](/guides/languages/) reword them, or add a language, as they do the editor's own:

```ts
createWorkbookEditor({
  plugins: [oneDrive({ clientId })],
  translations: { en: { 'plugin.page.ruled.onedrive.signIn': 'Sign in with Microsoft' } },
});
```

Each plugin's page says which languages it has and what its id is.

## Content Security Policy

The editor's own plugin code loads from the same place as the editor, so it needs nothing beyond what [Install › Content Security Policy](/start/install/#content-security-policy) lists. Each plugin needs what it connects to: a connector's page lists the hosts it talks to, frames it opens and scripts it loads.

Sign-in popups, as OneDrive's and Google Drive's, are blocked by `Cross-Origin-Opener-Policy: same-origin`. A page that sends it needs `same-origin-allow-popups`.

## Plugins are trusted code

A plugin runs on your page with everything your own code can reach: the workbook, its cells and saves, the page's DOM and its network. The editor doesn't sandbox it. Treat a plugin as you would any other dependency: use plugins you trust, from packages you can read, at versions you pin.

What the editor does do is keep a plugin's mistakes from becoming yours:

- Every call into a plugin is caught. A plugin that throws is named in a message, and your `onEvent` hears `{ type: 'pluginError', plugin, message }`.
- A plugin that fails to start is removed, with everything it added, and the editor carries on without it.
- A plugin written for a version of the contract this editor doesn't know isn't started, and is reported the same way.
- A failed save to a source keeps the changes, as a failed `onSave` does.

Plugins may keep small values in the browser's `localStorage`, under `ruled-plugin:<id>`: OneDrive's sign-in is kept in `sessionStorage` by Microsoft's library; the [Recent files](/plugins/writing/) example keeps a list of files.

## Plugins

- [OneDrive and SharePoint](/plugins/onedrive/): open workbooks from OneDrive and SharePoint and save them back.
- [Google Drive](/plugins/google-drive/): open workbooks and Google Sheets from Google Drive, save workbooks back, and save copies there.
- [Writing a plugin](/plugins/writing/): the contract, built around a complete "Recent files" plugin.
