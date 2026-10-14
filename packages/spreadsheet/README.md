# @ruled/sheet

**Spreadsheets that know the rules.** Ruled is a spreadsheet editor for web
apps. It opens the Excel files people already have (.xlsx, .xlsm and CSV),
lets them work as they would in Excel, checks every cell against a schema
with your own rules as they type, and hands the workbook back: as an .xlsx
with its charts, pivot tables and formulas intact, and as values in the shape
you asked for.

- **Documentation:** [ruled.page](https://ruled.page)
- **Live examples:** [the demo](https://demo.ruled.page), [a React supplier portal](https://react.ruled.page), [an Angular app](https://angular.ruled.page)
- **Source:** [github.com/ruled-page/ruled](https://github.com/ruled-page/ruled)

## Install

```sh
npm install @ruled/sheet
```

The package has two builds:

- **`@ruled/sheet`**: the React component and hooks. React 18 or 19 is a
  peer dependency.
- **`@ruled/sheet/standalone`**: everything bundled, React, the worker and
  the styles included, defining a `<workbook-editor>` element for any other
  page: plain HTML, Angular, Vue, Svelte.

## React

```tsx
import { useWorkbookEditor, WorkbookEditor } from '@ruled/sheet';
import '@ruled/sheet/style.css';

export function PriceListReview({ file }) {
  const editor = useWorkbookEditor({
    file, // { name, bytes }
    onSave: async (saved) => {
      await fetch('/api/price-lists', {
        method: 'POST',
        body: saved.bytes,
      });
    },
  });
  return (
    <div style={{ height: '80vh' }}>
      {editor && <WorkbookEditor editor={editor} />}
    </div>
  );
}
```

The editor fills its container, so give the container a height. The
workbook lives in a web worker, which Vite, webpack 5, Parcel and Next.js
bundle with your app by themselves. See
[React](https://ruled.page/start/react/).

## Next.js and Server Components

`<Workbook>` takes props a Server Component can give it: the file's bytes,
the options that are data, and a Server Action as `onSave`.

```tsx
// app/price-list/page.tsx
import { Workbook } from '@ruled/sheet';
import { savePriceList } from './actions'; // 'use server'

export default async function Page() {
  const bytes = await readPriceList(); // a Uint8Array
  return (
    <main style={{ height: '100vh' }}>
      <Workbook
        file={{ name: 'price-list.xlsx', bytes }}
        onSave={savePriceList}
      />
    </main>
  );
}
```

See [Next.js and Server Components](https://ruled.page/start/nextjs/).

## Any other page

```html
<workbook-editor src="/files/forecast.xlsx" save-handler="event">
</workbook-editor>

<script type="module">
  import '@ruled/sheet/standalone';

  const editor = document.querySelector('workbook-editor');
  editor.addEventListener('save', (event) => {
    const saved = event.detail;
    saved.waitUntil(fetch('/api/forecasts', {
      method: 'POST',
      body: saved.bytes,
    }));
  });
</script>
```

See [the web component](https://ruled.page/start/web-component/) and
[Angular](https://ruled.page/start/angular/).

## Schemas

A schema says what a workbook should hold. A workbook that matches is
checked as people type, saving waits until it passes, and your app gets its
values by field.

```ts
import { defineSchema } from '@ruled/sheet';

const timesheet = defineSchema({
  id: 'acme.timesheet',
  name: 'Timesheet',
  sheets: [{
    name: 'Hours',
    tables: [{
      key: 'entries',
      at: 'A1',
      fields: [
        { key: 'date', label: 'Date', type: 'date', required: true },
        { key: 'hours', label: 'Hours', type: 'number', min: 0, max: 24 },
      ],
    }],
  }],
});
```

Pass it as `schemas: [timesheet]`. Fields can have checks in code that call
your services, and business logic that fills cells in. See
[schemas](https://ruled.page/schemas/overview/).

## In Node

`recalculate` from `@ruled/sheet/node` opens a saved .xlsx, sets the cells
you give it, works out its formulas and saves it again, with no browser:

```ts
import { readFileSync, writeFileSync } from 'node:fs';
import { recalculate } from '@ruled/sheet/node';

const result = await recalculate(readFileSync('quote.xlsx'), {
  set: { 'Inputs!B2': 120 },
});
writeFileSync('quote-priced.xlsx', result.bytes);
```

See [recalculating saved workbooks](https://ruled.page/formula/saved-workbooks/).

## What's in it

- **Formulas,** calculated live by its own engine,
  [`@ruled/formula`](https://www.npmjs.com/package/@ruled/formula): more than
  470 of Excel's functions, dynamic arrays and `LAMBDA`.
- **Charts and pivot tables,** drawn from the current cells and editable as
  in Excel.
- **Formatting, sorting, filters, data validation, names, hyperlinks, notes
  and comments,** with Excel's keys.
- **Saving only what changed:** everything the editor doesn't show stays in
  the file as it was.
- **Theming** through design tokens, **languages** (English and German built
  in), keyboard use and screen reader support, phones and tablets.

What's left out, and what Excel files keep, is in
[Excel fidelity](https://ruled.page/guides/excel-fidelity/).

## Requirements

Current Chrome, Edge, Firefox and Safari. A site with a Content Security
Policy needs to allow a few things, listed under
[Install](https://ruled.page/start/install/#content-security-policy).

## Licence

MIT. The builds include other open-source packages, listed with their
licences in `THIRD-PARTY-NOTICES.md`.
