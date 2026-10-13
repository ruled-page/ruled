<img src="brand/ruled-lockup.svg" alt="Ruled" height="48">

# Ruled

**Spreadsheets that know the rules.** Ruled opens real Excel workbooks in a
web app, checks every cell against a schema with the app's own sync and async
rules as people type, and saves the .xlsx back with its charts, pivot tables
and formulas intact.

| | |
| --- | --- |
| Documentation | [ruled.page](https://ruled.page): installing, embedding, schemas, guides, the API reference, and `llms.txt` for coding agents |
| Live examples | [The demo](https://demo.ruled.page) · [a React supplier portal](https://react.ruled.page) · [an Angular app](https://angular.ruled.page) |
| Packages | `@ruled/sheet`, the editor · `@ruled/formula`, its calculation engine on its own |

This README is for working on Ruled itself. To use it in an app, start with
the documentation.

## Using Ruled

```bash
npm install @ruled/sheet
```

```tsx
import { useWorkbookEditor, WorkbookEditor } from '@ruled/sheet';
import '@ruled/sheet/style.css';

function Review() {
  const editor = useWorkbookEditor({ schemas, onSave: (file) => upload(file.bytes) });
  return editor && <WorkbookEditor editor={editor} />;
}
```

Where to go next:

- [Install](https://ruled.page/start/install/), then [React](https://ruled.page/start/react/), [plain HTML and other frameworks](https://ruled.page/start/web-component/) or [Angular](https://ruled.page/start/angular/)
- [Schemas](https://ruled.page/schemas/overview/): fields, tables, async checks and business logic
- [Capabilities](https://ruled.page/start/capabilities/), and [what Excel files keep](https://ruled.page/guides/excel-fidelity/) when the editor saves them
- [The formula engine on its own](https://ruled.page/formula/introduction/)
- [The API reference](https://ruled.page/reference/api/readme/)

## The repository

An npm workspace. Packages use each other's source, through the `source`
condition of their exports that TypeScript and Vite are set to prefer, so a
change to the engine shows in the editor and the examples straight away.

| Folder | What it is |
| --- | --- |
| `packages/spreadsheet` | `@ruled/sheet`: the React component, the `<workbook-editor>` element and the worker |
| `packages/formula-engine` | `@ruled/formula`: the calculation engine, and Excel's references and formula text |
| `packages/e2e` | The browser tests, and a React app that installs the packed library as a host would |
| `examples/demo` | The demo, with the sample workbooks |
| `examples/react-portal` | The editor in a React supplier portal |
| `examples/angular-portal` | The editor in an Angular app (its own install) |
| `examples/html` | The `<workbook-editor>` element in a plain page |
| `docs` | The documentation site: Astro and Starlight (its own install, as TypeDoc needs TypeScript 6) |
| `brand` | Ruled's mark, wordmark, colours and type, and [how to use them](brand/README.md) |
| `scripts` | Building and deploying the sites, releasing, and the bundle checks |

[`BACKLOG.md`](BACKLOG.md) lists what is deliberately not done yet and the
known issues.

## Getting started

Node 22.23, which `.nvmrc` asks for (`nvm use` or `fnm use`). Some checks
need more: Python with openpyxl for the round-trip check, and the .NET SDK
for `validate:saves`.

```bash
npm install
npm run dev
```

The demo is then at http://localhost:5180: pick a sample or drop in a
workbook.

| Command | Serves |
| --- | --- |
| `npm run dev` | The demo, at http://localhost:5180 |
| `npm run dev:portal` | The React portal, at http://localhost:5181 |
| `npm run build:lib`, then `npm run dev:html` | The element in a plain page, at http://localhost:5182/examples/html/ |
| `npm run build:lib`, then `npm install` and `npm start` in `examples/angular-portal` | The Angular app, at http://localhost:4200 |
| `npm run docs:dev` | The documentation site, at http://localhost:4321 |

`npm run samples` makes the demo's sample workbooks again, including the
1.1 million cell `big.xlsx`, which is generated rather than committed.

## Checks

| Command | What it checks |
| --- | --- |
| `npm run typecheck` | TypeScript, across the workspace |
| `npm run lint` | Oxlint, with React's hooks rules and accessibility checks (`.oxlintrc.json`) |
| `npm test` | The unit tests, with Vitest |
| `npm run test:coverage` | The same, failing below 98% of lines, 97% of statements and functions and 90% of branches; the screens are covered by the browser tests |
| `npm run e2e` | The browser tests, with Playwright, in Chromium (see below) |
| `npm run check:bundle` | After the builds: each bundle within its size budget, the engine only in the worker, and the type declarations importing nothing but React |
| `npm run test:perf` | Opening, calculating, editing and saving `big.xlsx`, against time budgets |
| `npm run engine:bench`, `npm run lookup:bench` | The engine's speed and memory on the samples and the test corpus, and on lookup-heavy workbooks as they grow |
| `npm run excel:check` | Workbooks of formulas for Excel to grade (see below) |
| `npm run validate:saves` | After `npm run build:verify`: saved copies of the samples and the test corpus are as valid as their originals, by Microsoft's Open XML SDK |

The browser tests run against production builds, on servers they start on
ports 4173 to 4177 (`E2E_PORT` moves them). They cover editing, saving,
schemas, the hosts, keyboard-only use, phones (an emulated Pixel 7), and an
axe scan of every screen, menu and dialog in each theme.
`E2E_BROWSERS=all` adds Firefox and WebKit, after
`npx playwright install firefox webkit`.

[CI](.github/workflows/ci.yml) runs all of these on Linux, the browser tests
in Chromium, Firefox and WebKit, and also builds the Angular app and a React
app from the packed library. To reproduce a Linux failure, run the tests in
Playwright's image (`mcr.microsoft.com/playwright:v1.63.0-noble`) with two
workers, as CI does.

### Calculations against Excel

`npm run excel:check` writes a workbook per family of functions to
`packages/spreadsheet/tests/excel-checks/`: each formula, the engine's answer
beside it, and a cell Excel fills on opening with whether the two agree. Open
them in Excel (Excel for the web has the newest functions). Once Excel agrees,
rename `<family>.pending.json` to `<family>.json`, and
`excel-checks.test.ts` holds the engine to those answers from then on.

### Round trips on real files

`npm run build:verify` builds a checker that round-trips workbooks through
the editor:

```bash
node packages/spreadsheet/dist-verify/verify.js path/to/workbooks --out verify-out --structure
```

For each workbook it checks that saving without edits gives back the same
bytes; that edits in a copy read back while every other cell, formula and
part stays as it was; how many formulas the engine calculates as Excel saved
them; and, with `--structure`, that references survive an inserted row.
`verify-out/` gets the edited copies and a report.

## How it works

```
Page                                      Worker
────────────────────────────              ─────────────────────────────────────
React and Glide Data Grid (canvas)        WorkbookSession
SheetView: the file's cells, edits  ◄──     OoxmlPackage   unzipped, every part kept
store: undo, redo, dialogs          ──►     sheetScanner   cells and their places in the XML
                                            engine         @ruled/formula, checked against Excel
                                            writer         changed cells patched in place
```

- **Reading** (`core/workbook.ts`, `core/sheetScanner.ts`): the package is
  unzipped, and each worksheet's cells are scanned into compact typed arrays
  that record where each sits in the original XML. Every other part is left
  as it is.
- **Calculating** (`packages/formula-engine`, `core/session.ts`): the engine
  works out the formulas, and on opening each result is compared with the
  value Excel saved; a formula that disagrees keeps Excel's value. Table
  references are handed to the engine as the cells they stand for, and LET
  is written out before a formula reaches it. A formula the engine can't
  calculate, such as one linking to another workbook, keeps Excel's value
  and is marked out of date when its inputs change. A spilling formula is
  given exactly the room its result needs, and given it again when that
  changes.
- **Editing** (`core/session.ts`, `core/input.ts`, `core/fill.ts`): typed and
  pasted text is read as Excel reads it, in the user's locale and the
  workbook's date system, and references move as Excel moves them.
- **Charts and pivot tables** (`core/charts.ts`, `core/pivotTable.ts`):
  charts are drawn with Chart.js from the current cells, loaded only for
  workbooks that have them; pivot tables are worked out again from their
  source as it changes.
- **Saving** (`core/writer.ts`): only changed cells are rewritten, and Excel is
  told to recalculate and refresh pivot tables on opening. Formulas are
  written as Excel writes them, dynamic arrays included. A workbook with no
  edits comes back byte for byte.

Paths are under `packages/spreadsheet/src`.

## Building and deploying

`npm run build:lib` builds both packages. The editor goes into
`packages/spreadsheet/dist`: `react` for React hosts (React 18 or 19 is a
peer dependency), `standalone` with React, the worker and the styles
bundled, and `types`, one declaration file for each. Everything else is
bundled, so it has no dependencies of its own. The engine goes into
`packages/formula-engine/dist`, as one module and its types, with its
formula parser and number formatter as its two dependencies; within the
workspace only its root is public, the rest is read from source by the
editor. `npm pack -w @ruled/formula` (or `-w @ruled/sheet`) packs either one.

The documentation site and the examples are static, each its own
Cloudflare Pages project:

| Site | Address | Project | Deploy |
| --- | --- | --- | --- |
| Documentation | https://ruled.page | `ruled-docs` | `npm run deploy:docs` |
| Demo | https://demo.ruled.page | `workbook-editor-poc` | `npm run deploy:demo` |
| React portal | https://react.ruled.page | `workbook-editor-portal` | `npm run deploy:portal` |
| Angular app | https://angular.ruled.page | `workbook-editor-angular` | `npm run deploy:angular` |

`npm run deploy` deploys the three examples. A project is made once with
`npx wrangler pages project create <project> --production-branch main`, and
its address is set under Custom domains in the Cloudflare dashboard. The
sites link to each other at these addresses; `DOCS_URL`, `DEMO_URL`,
`PORTAL_URL` and `ANGULAR_URL` build them for others. A `_headers` file keeps
each site out of search results: the docs' own, and the demo's, which the
other examples copy.

`npm run build:sites` builds the examples into `dist-sites/` without
deploying, and `npm run pages:dev` serves the demo as Pages would.

## Releases

Versions follow semantic versioning; before 1.0.0, a minor version may change
the API. `packages/spreadsheet/CHANGELOG.md` says what changed: add notes
under "Unreleased" as you go. Then, from a clean working tree:

```bash
npm run release -- minor
```

(or `patch`, `major`, or a version). It runs the checks CI runs short of the
browser tests, sets the version, moves the notes under it, writes
`THIRD-PARTY-NOTICES.md` again, commits, tags and packs the package into
`releases/`. Nothing is pushed or published.

## Licence

Proprietary; see `LICENSE`. The builds include open-source packages, listed
with their licence texts in `packages/spreadsheet/THIRD-PARTY-NOTICES.md`,
which `npm run notices` writes and the package ships.
