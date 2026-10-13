---
title: React
description: Show the editor in a React app, open a workbook and receive it when it's saved.
sidebar:
  order: 3
---

Create an editor with `createWorkbookEditor`, show it with `<WorkbookEditor>`, and destroy it when you're done with it. The editor is created outside rendering so that re-rendering never re-creates it.

```tsx
import { useEffect, useState } from 'react';
import { createWorkbookEditor, WorkbookEditor, type EditorHandle } from '@sitepen/spreadsheet';
import '@sitepen/spreadsheet/style.css';
import { priceList } from './schemas';

export function PriceListReview({ file }: { file: { name: string; bytes: ArrayBuffer } }) {
  const [editor, setEditor] = useState<EditorHandle | null>(null);

  useEffect(() => {
    const editor = createWorkbookEditor({
      schemas: [priceList],
      // People open files from your app, not from their disk.
      features: { open: false },
      saveLabel: 'Submit price list',
      onSave: async (saved) => {
        await fetch('/api/price-lists', { method: 'POST', body: saved.bytes });
      },
    });
    setEditor(editor);
    void editor.open(file.bytes, file.name);
    return () => editor.destroy();
  }, [file]);

  return <div style={{ height: '80vh' }}>{editor && <WorkbookEditor editor={editor} />}</div>;
}
```

The editor fills its container, so give the container a height.

## The component's props

| Prop | What it does |
| --- | --- |
| `editor` | The editor to show, from `createWorkbookEditor()`. |
| `emptyState` | What to show before a workbook is open. It receives a function that opens the file picker. |
| `toolbarEnd` | Your own controls, at the end of the toolbar before the Save button. |
| `className` | A class for the editor's root element. |

## Following the editor's state

The handle tells you when its status, file name, edit count or unsaved state change:

```tsx
useEffect(() => {
  if (!editor) return;
  return editor.subscribe(() => setUnsaved(editor.unsaved));
}, [editor]);
```

Everything else the handle does (opening, saving, going to a cell, changing the theme or language) is in [Events and the handle](/guides/events-and-handle/).

## Options that change while it's open

Theme, language, user, schemas and custom lists change without closing the workbook:

```ts
editor.setTheme(dark ? darkTheme : lightTheme);
editor.setLanguage(i18n.language);
```

Any other option is set when the editor is created; create a new editor to change it.
