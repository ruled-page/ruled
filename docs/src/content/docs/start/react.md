---
title: React
description: Show the editor in a React app with useWorkbookEditor, open a workbook and receive it when it's saved.
sidebar:
  order: 4
---

`useWorkbookEditor` makes an editor that lives as long as your component: it's made when the component mounts and destroyed, with its worker, when it unmounts. Render it with `<WorkbookEditor>`.

```tsx
import { useWorkbookEditor, WorkbookEditor, type WorkbookFile } from '@ruled/sheet';
import '@ruled/sheet/style.css';
import { priceList } from './schemas';

export function PriceListReview({ file }: { file: WorkbookFile }) {
  const editor = useWorkbookEditor({
    schemas: [priceList],
    file,
    // People open files from your app, not from their disk.
    features: { open: false },
    saveLabel: 'Submit price list',
    onSave: async (saved) => {
      await fetch('/api/price-lists', { method: 'POST', body: saved.bytes });
    },
  });

  return <div style={{ height: '80vh' }}>{editor && <WorkbookEditor editor={editor} />}</div>;
}
```

The editor fills its container, so give the container a height. It's `null` on the first render, before it's made.

## Opening a workbook

`file` (its `name` and `bytes`) is opened once the editor is ready, and again whenever it's a different file: another name or other bytes. A new object holding the same name and bytes isn't a new file, so it can be written inline. If the user has unsaved changes they're asked first, and if they keep their workbook it stays open. A file that can't be opened says why in the editor.

To open files at moments of your own choosing, leave `file` out and call `editor.open(bytes, name)`, which resolves once the workbook is open.

## Options as they change

- **Theme, language, user, schemas and custom lists** follow the options as they change, without closing the workbook. Pass `language: i18n.language` and the editor follows your app's language. A new array or object with the same entries isn't a change, so they can be written inline.
- **`onSave`, `onEvent`, `onError` and `onLinkClick`** can be new functions on every render, closing over your component's props and state: the editor always calls the latest.
- **Every other option** (`features`, `numbers`, `locale`, `limits` and the rest) is read when the editor is made. To change one, give the component a new `key`, which makes a new editor:

```tsx
<PriceListReview key={submission.id} file={submission.file} />
```

## The editor's state as React state

`useEditorState` re-renders your component as the editor's status, file name and edits change, for your own unsaved badge, page title or button:

```tsx
const { unsaved, editedCount, fileName } = useEditorState(editor);

useEffect(() => {
  document.title = unsaved ? `• ${fileName}` : fileName;
}, [unsaved, fileName]);
```

## The component's props

| Prop | What it does |
| --- | --- |
| `editor` | The editor to show. |
| `emptyState` | What to show before a workbook is open. It receives a function that opens the file picker. |
| `toolbarEnd` | Your own controls, at the end of the toolbar before the Save button. |
| `className` | A class for the editor's root element. |

Everything else the editor does (saving from your own button, going to a cell, reading the workbook without saving) is on the handle: see [Events and the handle](/guides/events-and-handle/).

## Without the hook

`createWorkbookEditor(options)` makes an editor you manage yourself: for one shared across components, or made outside React. Destroy it when you're done with it, which stops its worker:

```tsx
const editor = createWorkbookEditor({ schemas: [priceList] });
// …
editor.destroy();
```

Make it in an effect, not with `useMemo` or during rendering: React may throw a memoized value away and make it again, which would leave an editor's worker running with nothing to stop it.
