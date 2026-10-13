---
title: Install
description: Add the package to your app, and choose the React build or the self-contained one.
sidebar:
  order: 3
---

The editor is one package, `@sitepen/spreadsheet`, with two builds:

| Import | What it is | Use it from |
| --- | --- | --- |
| `@sitepen/spreadsheet` | The React component, the `useWorkbookEditor` hook, `createWorkbookEditor` and the types. React 18 or 19 is a peer dependency. | React apps |
| `@sitepen/spreadsheet/standalone` | Everything bundled, React, the worker and the styles included, defining the `<workbook-editor>` element. | Plain HTML, Angular, Vue, Svelte, server-rendered pages |

```sh
npm install @sitepen/spreadsheet
```

:::note[Not on the npm registry yet]
Until the package is published, install the tarball a release makes (`npm run release` in the editor's repository packs it into `releases/`):

```sh
npm install ./sitepen-spreadsheet-0.1.0.tgz
```
:::

## The React build

Import the stylesheet once, next to the component:

```ts
import '@sitepen/spreadsheet/style.css';
```

The component starts its worker with `new Worker(new URL('./assets/workbook.worker.js', import.meta.url))`. Vite, webpack 5 and Parcel recognise that and bundle the worker with your app, so there is nothing more to set up. With a bundler that doesn't, serve `@sitepen/spreadsheet/worker` yourself and pass where it is:

```ts
createWorkbookEditor({
  createWorker: () => new Worker('/assets/workbook.worker.js', { type: 'module' }),
});
```

## The standalone build

Importing it defines `<workbook-editor>`. It injects its own styles, and its worker is inlined, so it needs nothing else. It's larger than the React build, so load it when the editor is needed:

```ts
await import('@sitepen/spreadsheet/standalone');
```

## Browsers

Current Chrome, Edge, Firefox and Safari. The editor needs module workers, which all of them have.

## Styles stay inside

Every class the editor uses starts with `wbe-`, and all its styles are scoped under `.wbe`, so your app's `.btn` or `.modal` rules don't reach into it, and its rules don't reach out. To restyle it, use [theme tokens](/guides/theming/).
