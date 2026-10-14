---
title: Install
description: Add the package to your app, and choose the React build or the self-contained one.
sidebar:
  order: 3
---

The editor is one package, `@ruled/sheet`, with two builds:

| Import | What it is | Use it from |
| --- | --- | --- |
| `@ruled/sheet` | The React components, the `useWorkbookEditor` hook, `createWorkbookEditor` and the types. React 18 or 19 is a peer dependency. It's client code, which a React Server Component can render. | React apps, Next.js |
| `@ruled/sheet/standalone` | Everything bundled, React, the worker and the styles included, defining the `<workbook-editor>` element. | Plain HTML, Angular, Vue, Svelte, server-rendered pages |

```sh
npm install @ruled/sheet
```

## The React build

Import the stylesheet once, next to the component:

```ts
import '@ruled/sheet/style.css';
```

The component starts its worker with `new Worker(new URL('./assets/workbook.worker.js', import.meta.url))`. Vite, webpack 5, Parcel and Next.js recognise that and bundle the worker with your app, so there is nothing more to set up. With a bundler that doesn't, serve `@ruled/sheet/worker` yourself and pass where it is:

```ts
createWorkbookEditor({
  createWorker: () => new Worker('/assets/workbook.worker.js', { type: 'module' }),
});
```

## The standalone build

Importing it defines `<workbook-editor>`. It injects its own styles, and its worker is inlined, so it needs nothing else. It's larger than the React build, so load it when the editor is needed:

```ts
await import('@ruled/sheet/standalone');
```

## Browsers

Current Chrome, Edge, Firefox and Safari. The editor needs module workers, which all of them have.

## Fonts

Workbooks are mostly set in Microsoft's fonts, which many devices lack: a Mac, a phone or Linux usually has no Calibri or Cambria. Where the device lacks one, the grid draws its text in a free font drawn to the same widths, so text fits its columns as it does in Excel. Each is fetched only then, from a chunk of the package, Calibri's as soon as the editor starts:

| Excel's font | Drawn with |
| --- | --- |
| Calibri | Carlito |
| Cambria | Caladea |
| Arial | Arimo |
| Times New Roman | Tinos |
| Courier New | Cousine |
| Georgia | Gelasio |

They cover Latin letters; others are drawn in the device's own fonts. Other fonts a device lacks, Aptos among them, are drawn in its own font, a little wider or narrower than in Excel.

## Content Security Policy

On a site with a Content Security Policy, allow what the editor loads:

- `font-src data:`, for the fonts above, which the builds carry as `data:` URLs.
- For the React build, its worker, a file beside the script: `worker-src 'self'` when you serve it from your own site.
- For the standalone build, its worker, carried inside the script and started from a `blob:` URL (or a `data:` URL where that's refused): `worker-src blob: data:`. It also adds its styles to the page in a `<style>` element, which needs `style-src 'unsafe-inline'`.

## Styles stay inside

Every class the editor uses starts with `wbe-`, and all its styles are scoped under `.wbe`, so your app's `.btn` or `.modal` rules don't reach into it, and its rules don't reach out. To restyle it, use [theme tokens](/guides/theming/).
