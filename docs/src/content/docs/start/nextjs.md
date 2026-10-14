---
title: Next.js and Server Components
description: Render the editor from a React Server Component, with a workbook read on the server and a Server Action that receives what's saved.
sidebar:
  order: 5
---

The editor runs in the browser: it draws the sheet on a canvas, keeps the workbook in a Web Worker and answers keys and pointers. In an app with React Server Components, such as Next.js's App Router, it's a client component. The package says so (`'use client'`), so a Server Component can import it and render it without a wrapper of your own.

`<Workbook>` is the editor in one component for a Server Component to render. Its props are plain data: the file's name and bytes, the theme, features, language and the other options that aren't functions. Its `onSave` can be a Server Action.

```tsx
// app/price-list/page.tsx
import { Workbook } from '@ruled/sheet';
import { readPriceList } from '@/lib/price-lists';
import { savePriceList } from './actions';

export default async function PriceListPage() {
  const bytes = await readPriceList(); // a Uint8Array
  return (
    <main style={{ height: '100vh' }}>
      <Workbook
        file={{ name: 'price-list.xlsx', bytes }}
        features={{ open: false }}
        saveLabel="Submit price list"
        onSave={savePriceList}
      />
    </main>
  );
}
```

```ts
// app/price-list/actions.ts
'use server';
import type { SaveAnswer, SavedWorkbookData } from '@ruled/sheet';
import { storePriceList } from '@/lib/price-lists';

export async function savePriceList(saved: SavedWorkbookData): Promise<SaveAnswer> {
  const stored = await storePriceList(saved.bytes, saved.values);
  if (!stored) return { conflict: true };
}
```

Import the stylesheet once, in your root layout:

```tsx
// app/layout.tsx
import '@ruled/sheet/style.css';
```

The editor fills its container, so give the container a height. Next.js bundles the editor's worker with your app, as Vite and webpack do.

## Saving with a Server Action

`onSave` receives the saved workbook as plain data (`SavedWorkbookData`): the `.xlsx` bytes, the values in your schema's shape, the changes and the history. That's everything `onSave` receives elsewhere except `cell()` and `range()`, which are functions and can't be sent to a server. See [Opening and saving](/guides/opening-and-saving/#receiving-the-saved-workbook) for each property.

An error thrown in a Server Action doesn't reach the page as written, as frameworks hide server errors in production. Answer instead:

- `{ error: 'The price list is closed.' }` shows the message and keeps the user's changes, so they can try again.
- `{ conflict: true }` says someone else saved first and offers the user's changes as a download, as throwing `SaveConflictError` does. Give your own message with `{ conflict: 'Ana saved it at 10:02.' }`.

Answering nothing means the save worked.

## Options that are functions

Functions can't be passed from a Server Component to the browser, Server Actions apart. For `onEvent`, `onError`, `onLinkClick`, `createWorker`, or schemas with code (`validate`, `onChange`, `beforeSave` and the other hooks), write a client component of your own and render that from your Server Component:

```tsx
// app/price-list/price-list-editor.tsx
'use client';
import { Workbook, type WorkbookProps } from '@ruled/sheet';
import { priceList } from './schemas';

export function PriceListEditor(props: WorkbookProps) {
  return <Workbook {...props} schemas={[priceList]} onEvent={(event) => console.log(event)} />;
}
```

The Server Component still passes the file and the Server Action: `<PriceListEditor file={{ name, bytes }} onSave={savePriceList} />`.

Everything else in `@ruled/sheet` is client code too. Call `defineSchema`, `schemaFromJSON` and the rest from client components, not from a Server Component.

## Rendered on the server

On the server, and until the browser has made the editor, `<Workbook>` renders the editor's empty box, so the page's layout doesn't move when the editor appears.

Each render of the Server Component sends the file's bytes again. A copy holding the same bytes isn't opened again, so refreshing the page's data, with `router.refresh()` or by revalidating after a Server Action, leaves the user's edits alone. A different file is opened, asking the user first if they have unsaved changes.

To recalculate a workbook on the server before sending it, or after it's saved, use `@ruled/sheet/node`: see [Recalculating saved workbooks](/formula/saved-workbooks/).
