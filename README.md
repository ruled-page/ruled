# Workbook editor (proof of concept)

Open an Excel workbook in the browser, edit it like a spreadsheet, and save it
back as .xlsx without losing what the editor can't display.

- View any .xlsx or .xlsm: every sheet, number formats, fonts, fills, column
  widths, merged cells, hidden rows and columns, frozen columns.
- Edit cells, paste from Excel or Google Sheets, copy and paste within the
  workbook (formulas move their references like Excel), drag to fill (series,
  dates, months, "Item 1" → "Item 2"), clear, undo and redo.
- Formulas recalculate live as you edit (HyperFormula, ~420 Excel functions).
  Formulas it can't calculate keep Excel's saved value and are marked; when
  their inputs change they are marked stale, and Excel recalculates them on
  open.
- Save writes only the cells that changed back into the original file.
  Charts, pivot tables, macros, conditional formatting, validation, comments,
  pictures, sparklines and everything else stay byte-for-byte as they were.
- Notes and threaded comments are marked in the corner of their cell and
  shown when you point at the cell or select it. They are read-only.
- "What's in this file" lists everything the editor keeps but can't show,
  and the editor asks before any edit would replace a formula with a value.
- Excel's own data validation is enforced. List rules get a dropdown,
  whether the list is typed into the rule, a range on any sheet, a named
  range, or a formula such as `OFFSET(…,COUNTA(…))` or a dependent
  `INDIRECT(C2)`; lists are read live, so editing the source updates the
  dropdown. Whole number, decimal, date, time, text-length and custom-formula
  rules check typed values, with bounds that may themselves be formulas or
  cells. Failures raise the rule's stop, warning or information alert, and
  the formula bar shows the rule or its input message. Pasted values and
  formulas aren't checked, as in Excel.
- Right-click a cell or heading to cut, copy and paste, insert or delete rows
  and columns, sort, filter by values, or find and replace (⌘F / Ctrl+F).
  Cutting and pasting moves cells, and references to them follow.
- Optional **schemas** describe the workbooks a consumer expects. A workbook
  that matches one is validated and constrained by it; one that matches none
  opens as a plain spreadsheet with a warning. Schemas can be written in
  TypeScript (with async checks and business logic), generated from a
  workbook, edited in the schema designer, and exported or imported as JSON.
- Number handling a host or schema can tighten: show and copy every stored
  digit instead of the rounded format, ask before pasting values that would
  be stored differently from how they read (day/month-ambiguous dates, more
  than 15 significant digits, leading zeros), and list every changed cell
  for review before saving.
- Dates and decimals are read in the user's locale (`3/4/2026` is 3 April in
  the UK, `1.234,5` is 1234.5 in Germany).
- A Changes panel lists every cell that differs from the opened file, old and
  new value, and a timestamped history of actions. Unsaved edits are kept in
  the browser and offered back if the same file is reopened after a crash or
  a closed tab.

Everything runs in the browser; the file never leaves the machine.

## Running it

```bash
npm install
npm run samples   # rebuilds the samples in public/samples, including the 100k-row big.xlsx
npm run dev
```

`npm run build:lib` builds the embeddable library; with the dev server
running, http://localhost:5173/examples/html/ shows it in a plain page.

Then open http://localhost:5173 and pick a sample or drop in a workbook.

http://localhost:5173/embedded.html is the editor embedded in a host app
(`src/embedded`): a procurement portal where suppliers submit price lists.
A supplier uploads a workbook (or tries the sample, which has one problem of
each kind) and corrects it in the editor under the portal's brand and
language. Three kinds of check run: each column's type and limits, the Excel
validation rules the template carries, and the portal's business checks,
which call its (simulated, local) services. Products must be in the
catalogue, prices more than 15% from the last approved price are flagged,
and a list uses one currency. "Download a copy" exports at any point;
"Submit price list" hands the portal the file and its values, which it reads
into its own records (`parse.ts`) and shows as a summary, a breakdown by
category and the change from approved prices. Approving a list makes its
prices the ones the next list is checked against. Submissions are kept in
the browser's IndexedDB; nothing is sent anywhere.

`npm test` runs the core tests (reading, recalculation, editing, saving and
reopening the samples). `npm run e2e` runs the browser tests with Playwright
against a production build: editing, saving, schemas, the HTML host page,
keyboard-only use, and an axe accessibility scan of every screen, menu and
dialog in the light and dark themes (run `npx playwright install chromium`
once first). `.github/workflows/ci.yml` runs both, the builds, and the
round-trip check on the samples. `public/samples/kitchen-sink.xlsx` and
`public/samples/formatting.xlsx` are committed; they are generated by
`scripts/gen_kitchen_sink.py` and `scripts/gen_formatting.py` with XlsxWriter
if you need to rebuild them.

## Embedding the editor

The editor is a component; the demo app in `src/demo` is one consumer of it.
`npm run build:lib` builds two library flavours into `dist-lib/`:

- `dist-lib/react`: the React component, with React as a peer dependency.
- `dist-lib/standalone`: React, the worker and the styles bundled, plus a
  `<workbook-editor>` custom element, for hosts without React.

React:

```tsx
import { WorkbookEditor, WorkbookStore } from 'fake-spreadsheet';
import 'fake-spreadsheet/style.css';

const store = new WorkbookStore({ schemas, onSave: (file) => upload(file.bytes) });
<WorkbookEditor store={store} />;
```

Any other host:

```ts
import { mountWorkbookEditor } from 'fake-spreadsheet/standalone';

const editor = mountWorkbookEditor(element, { schemas, features: { open: false } });
await editor.open(bytes, 'report.xlsx');
```

Or the element (see `examples/html/index.html`):

```html
<workbook-editor src="/files/report.xlsx" save-handler="event"></workbook-editor>
```

In Angular (see `examples/angular`, a working Angular 22 app), the element
is used like any custom element: allow it with `CUSTOM_ELEMENTS_SCHEMA`,
bind its options, and listen to its events. Loading the editor with a
dynamic `import()` keeps it out of Angular's initial bundle; the element
keeps options Angular set before it loaded. The standalone build has type
declarations, including `WorkbookEditorElement` for the element itself.

```ts
const editorLoaded = import('fake-spreadsheet/standalone');

@Component({
  schemas: [CUSTOM_ELEMENTS_SCHEMA],
  template: `<workbook-editor save-handler="event" [options]="options" (save)="onSave($event)"></workbook-editor>`,
})
export class Review {
  options: EditorOptions = { schemas: [forecastSchema], saveLabel: 'Save to the app' };
  onSave(event: Event) {
    const saved = (event as CustomEvent<SaveEventDetail>).detail;
    saved.waitUntil(this.api.upload(saved.bytes, saved.values));
  }
}
```

To run it: `npm run build:lib`, then in `examples/angular`, `npm install` and
`npm start` (Angular's CLI needs Node 22.22.3 or later; the folder's
`.node-version` asks for it). `E2E_ANGULAR=1 npx playwright test
e2e/angular.spec.ts` tests it.

Options include `schemas`, `numbers`, `onSave`, `saveLabel`, `features` (hide
the open, save, download or schema buttons when the host provides them), `user` (recorded in the
history), `drafts`, `warnOnUnload`, `locale`, `limits`, `macros`, `recalc`
and `createWorker`. All styles are scoped under `.wbe`.

Opened files are treated as untrusted. Before anything is decompressed the
editor checks the file size (default limit 200 MB), the total unzipped size
(1.5 GB), the compression ratio of large parts (zip bombs) and the part
count; after reading, the cell count (5 million). `macros: 'remove'` strips
the VBA project on save and saves as .xlsx; `macros: 'block'` refuses
macro-enabled files. Links to other workbooks are kept but never fetched.

Large workbooks: above 200,000 formulas or 2 million cells, live
recalculation starts off (`recalc: 'auto'`, the default). Formulas then show
the values Excel saved and Excel recalculates on open; the status bar offers
to turn it on. On the 1.1-million-cell sample that is 1 second and about
200 MB to open, against 5.5 seconds and 750 MB with the engine. `recalc:
'on'` or `'off'` forces it either way. The engine only takes the sheets that
formulas use, adding others when a formula first names them, and sheet XML
is decoded again when saving rather than kept in memory.

**Saving to the host.** Edits are kept locally as a draft until saved; with
`onSave` (or the element's `save-handler="event"`), the Save button (labelled
by `saveLabel`) runs the schema checks and hands the host the saved workbook:

- `bytes` and `fileName`: the .xlsx, ready to upload or store;
- `values`: when a schema applies, its tables' records keyed by field and its
  single cells, as plain JSON (dates as ISO text), with `schema` naming the
  schema, its version and the errors and warnings it found;
- `cell(sheet, ref)` and `range(sheet, ref?)`: any value, for hosts without a schema;
- `changes` (each change, before and after) and `history`.

The host can keep the file, the values or both. A separate Download button
exports a copy without saving it to the host, and leaving the page warns only
about changes the host hasn't received. Without `onSave`, Save downloads the
file. If the stored copy changed since it was opened, throw
`SaveConflictError`: the editor keeps the user's work and offers it as a
download. `examples/html/index.html` shows what arrives.

### Theming

Every colour, font, type size, corner radius, shadow and icon can be
replaced; layout can't. The look comes from `--wbe-*` custom properties on
the editor's root, declared with their defaults at the top of
`src/editor/editor.css`. That includes the grid, which is drawn on a canvas
and reads its `--wbe-grid-*` properties when it renders. Set them in CSS:

```css
.wbe {
  --wbe-accent: #0b5cad;
  --wbe-font: Georgia, serif;
  --wbe-grid-selection: #0b5cad;
}
```

or through the `theme` option, where each token is the property's name in
camelCase (`gridHeaderBg` sets `--wbe-grid-header-bg`), and change it later
with `setTheme(theme)` on the store, the mount handle or the element. The
`EditorTheme` type lists every token. A dark theme sets `colorScheme: 'dark'`
along with the surfaces and the grid's colours. Text and borders the workbook
leaves in Excel's automatic colour follow `gridText` and `gridBorderAuto`, so
they stay readable; on cells the file fills, automatic text uses
`gridTextOnFill`, which stays dark because the fills are light. Colours the
workbook sets itself are kept.

Icons are replaced by name with `icons: { download: 'M12 3v12…' }` (SVG path
data on a 24×24 grid, drawn in the current text colour), or with any element
from React; `--wbe-icon-stroke-width` sets the line weight. React hosts can put
their own controls at the end of the toolbar with `toolbarEnd`. The demo's
theme picker (`src/demo/themes.ts`) shows three themes built this way.

### Languages

Every word the editor shows comes from a message catalogue, with i18next
and react-i18next. English and German are built in; `language` picks one
(it defaults to `locale`, and "de-CH" falls back to German, anything unknown
to English), and `setLanguage()` on the store, mount handle or element
changes it while the editor is open. Each editor has its own i18next
instance, so a host that uses i18next for its own pages keeps its setup.

Other languages, or different wording, come in as `translations`, keyed by
language, in i18next's JSON format: `{{name}}` placeholders,
`{{count, number}}`, and plural keys with the forms the language needs
(`rows_one`, `rows_few`, `rows_many`, `rows_other` for Polish). Start from the
exported `en` catalogue:

```ts
new WorkbookStore({
  language: 'fr',
  translations: { fr: { 'toolbar.open': 'Ouvrir', /* … */ }, en: { 'toolbar.saveAsExcel': 'Export' } },
});
```

Numbers and dates in the editor's own text follow the language; reading
typed numbers and dates follows `locale`. Right-to-left languages mirror the
editor around the grid, which stays left to right as Excel's sheets do.

What the workbook says is shown as written: sheet names, values, and the
messages its author gave Excel's data validation. Text the host writes in a
schema (its name and description, field descriptions, rule and hook
messages) can be given in several languages, `{ en: 'Price', de: 'Preis' }`,
and hooks are told the editor's language. Messages from the worker reach
the editor as keys rather than sentences, so changing language changes
everything on screen. The Node API and the command-line tools answer in
English.

### Keyboard and screen readers

Everything works without a mouse. In the grid, the keys are Excel's: arrows
and Tab move, typing replaces a cell, Enter edits it, Ctrl+D and Ctrl+R fill
down and right, Shift+F10 (or the menu key) opens the cell menu with
insert, delete, column width, sort and filter, Ctrl+G goes to a cell, and
F6 (Ctrl+F6 where the browser keeps F6) moves between the toolbar, formula
bar, sheet, side panel and sheet tabs, since Tab stays in the grid as it
does in Excel. Ctrl+/ (⌘/ on a Mac) lists every shortcut. Dialogs keep the
focus inside, close with Escape and hand the focus back.

On a phone the editor keeps the grid in view: the toolbar's actions scroll
sideways in one row, the side panel opens over the sheet when asked for
(going to a problem closes it and shows the cell), a cell's note or problem
appears under the formula bar since nothing hovers on touch, and the ⋯ button
there opens the cell menu, as there is no right-click. A tap selects a cell
and a second tap edits it; the formula bar edits it too.

Screen readers hear each cell as it is shown (the value, not the formula,
which the formula bar has), followed by any schema problem the grid marks in
colour, e.g. "twelve, error: Expected a number."; loading, messages and
errors are announced. The default and demo themes meet WCAG AA contrast.

## On a server (Node)

The same core runs without a browser, so a server can re-check what users
save, or a batch job can fill and check workbooks. `npm run build:node` builds
`dist-node/` (the package's `./node` export and the `workbook-check` command).

```ts
import { ServerWorkbook, validateWorkbook } from 'fake-spreadsheet/node';

const result = await validateWorkbook(bytes, { schemas }); // { ok, schema, errors, issues, values }

const wb = ServerWorkbook.open(bytes, { fileName: 'prices.xlsx', schemas });
await wb.set('Prices', 'A16', 'SKU-2001');   // typed as a user would; business logic runs
const issues = await wb.issues();            // async checks finished
const out = wb.save();                       // stamped with the schema's version
```

```
workbook-check prices.xlsx --schema price-list.schema.json --set 'Prices!D4=12.50' --out fixed.xlsx --json
```

It prints the issues (or JSON with `--json`, values with `--values`) and exits
0 when the workbook matches a schema with no errors, 1 when it has errors or
matches none, and 2 when it can't be read. `examples/schemas/` has the demo
schemas as JSON. Schemas loaded from JSON have no code hooks; code that
needs them imports the schemas as modules instead.

## Checking it on real files

`workbook-verify` (from `npm run build:node`) round-trips workbooks through
the editor and reports what, if anything, was lost:

```
node dist-node/verify.js path/to/workbooks --out verify-out --structure
```

For each .xlsx/.xlsm it checks that saving without edits gives back the same
bytes; edits a few cells in a copy and checks they read back, that every other
cell and formula is unchanged, and that charts, pivots, pictures and every
other part are byte for byte the same; has openpyxl read the edits when a
Python with it is available (an independent reader); reports how many of the
file's formulas the engine calculates exactly as Excel saved them, listing
the ones it doesn't; and with `--structure` inserts a row and checks every
reference still works. `verify-out/` gets the edited copies, `report.md` and
`report.json`. The report ends with the one step that needs desktop Excel:
opening the copies there.

Calculation accuracy is also tested with a catalogue of 117 formulas and the
results Excel gives (`tests/calc-accuracy.test.ts`). It found that
HyperFormula's INT, MOD, ROUND, ROUNDUP, ROUNDDOWN, TEXT and DATEDIF differ
from Excel and that RANK and RANK.EQ were missing; the editor replaces them
with Excel-exact versions (`src/core/excelFunctions.ts`).

## Sharing it on Cloudflare Pages

Everything is static (it all runs in the browser), so each example is its
own Cloudflare Pages site:

| Site | What it is | Pages project | Config |
| --- | --- | --- | --- |
| Demo | The editor with the sample workbooks | `workbook-editor-poc` | `wrangler.jsonc` |
| Portal | The editor embedded in a React supplier portal | `workbook-editor-portal` | `sites/portal/wrangler.jsonc` |
| Angular | The editor embedded in an Angular app | `workbook-editor-angular` | `sites/angular/wrangler.jsonc` |

Once, create each project (each prints its `*.pages.dev` address):

```bash
npx wrangler pages project create workbook-editor-poc --production-branch main
```

```bash
npx wrangler pages project create workbook-editor-portal --production-branch main
```

```bash
npx wrangler pages project create workbook-editor-angular --production-branch main
```

Then, each time you want to publish the current code, build and upload one
site, or all three with `npm run deploy`:

```bash
npm run deploy:demo
```

```bash
npm run deploy:portal
```

```bash
npm run deploy:angular
```

`npm run build:sites` builds them into `dist-sites/` without uploading
anything (`node scripts/build-sites.mjs portal` builds one). The sites link to
each other at `https://<project>.pages.dev/`; if Cloudflare gives a project
another address, or you use custom domains, set `DEMO_URL`, `PORTAL_URL` or
`ANGULAR_URL` when building. Angular's CLI needs Node 22.22.3 or later, so use
it for the Angular site (with fnm, `fnm use 22.23`).

To check a site locally the way Pages serves it, build it, then:

```bash
npx wrangler pages dev --port 8788
```

(`--cwd sites/portal` or `--cwd sites/angular` for the others;
`npm run pages:dev` builds and serves the demo.) `public/_headers` keeps the
sites out of search results; anyone with an address can open them, so put
Cloudflare Access in front to restrict that. The demo uploads `big.xlsx` only
if you have generated it locally.

## Schemas

Without a schema the editor is a general spreadsheet. A consumer passes
schemas when it creates the store (the example app registers the ones in
`src/demo/schemas.ts`):

```ts
const store = new WorkbookStore({ schemas: [priceListSchema, budgetSchema] });
```

Each opened workbook is checked against them in order and the first that fits
applies. If none fits, a banner names the closest schema and what didn't line
up (e.g. `Prices!C1 should be "Category" but is "Group"`), and the workbook
opens as a plain spreadsheet.

A schema lists sheets, and on each sheet tables and single cells. A table is
fields × records, and its orientation says which way the records run:

| Table | Fields | Records | Adding and removing |
|---|---|---|---|
| `orientation: 'rows'` | columns, labelled in the first row | rows | rows, up to `records.max` |
| `orientation: 'columns'` | rows, labelled in the first column | columns | columns |
| `records.labels` given | as above | fixed, labelled in the first column (or row) | none: a fixed grid |

```ts
export const priceListSchema = defineSchema({
  id: 'demo.price-list',
  name: 'Price list',
  numbers: { display: 'full', copy: 'raw' },
  sheets: [{
    name: 'Prices',
    tables: [{
      key: 'prices', at: 'A1', orientation: 'rows', records: { min: 1, max: 500 },
      fields: [
        { key: 'sku', label: 'SKU', type: 'text', required: true, unique: true, pattern: '^SKU-\\d{4}$',
          validate: async (value, { signal }) => /* call a service */ undefined },
        { key: 'category', label: 'Category', type: 'enum', values: ['Hardware', 'Software', 'Services'] },
        { key: 'price', label: 'Unit price', type: 'number', min: 0, decimals: 2, formula: 'forbid' },
        { key: 'priceIncTax', label: 'Price inc. tax', type: 'number', formula: 'require' },
      ],
    }],
  }],
  onChange: async (ctx) => { /* look things up, then ctx.set(...) other cells */ },
});
```

- **Fields** have a type (`text`, `number`, `integer`, `percent`, `date`,
  `boolean`, `enum`) and optional `required`, `min`/`max`, `decimals`,
  `pattern`, `values`, `unique`, `readOnly`, and `formula: allow | forbid |
  require`. Formula cells are checked on their calculated value.
- **Constraints**: labels fixed by the schema, read-only fields and cells
  outside the schema's tables can't be edited (configurable per sheet).
  Typed and pasted input is stored as the field's type, so `007` in a text
  field stays text and `12` in a percent field means 12%. Enum fields edit with
  a dropdown: Enter, typing or clicking the selected cell opens it, typing
  filters it, and arrow keys and Enter pick a value.
- **Records** can be inserted or deleted from the toolbar, and a new one can be
  started in the empty row (or column) after a table. New records take the
  calculated fields and number formats of the record before them, as in an
  Excel table.
- **Checks** run as you edit: field rules (type, required, bounds, decimals,
  text length, pattern, allowed values, unique), declarative table checks
  (`monotonic`, `firstEquals`, `sum`, and per record `compare`, `requiredIf`
  and `uniqueTogether`), and code hooks at field, table, sheet and workbook
  level. Hooks may be async; results from superseded runs are dropped, and
  pending checks show in the grid and the issues panel. Each async hook has
  `hookTimeoutMs` (default 10 seconds): a slow check is reported and ignored,
  slow business logic changes nothing, and a slow `beforeSave` stops the save.
- **Versions**: a schema with a `version` stamps files it saves (a custom
  document property, `WorkbookSchema = id@version`). A stamped file tries
  that schema first when opened, and says so if it was saved under another
  version.
- **Business logic**: `onChange` runs after each edit and can write other cells;
  its writes are part of the same undo step. `beforeSave` can stop a save.
  Saving with errors is blocked, or asks first with `saveWithErrors: 'confirm'`.

Everything except the hook functions is JSON, so the **schema designer**
(toolbar → Schemas) can generate a schema from the open workbook (every table
on each sheet, and "Label | value" cells as single cells), edit it in a form or
as JSON, add tables and single cells by selecting them on the sheet, test it
against the workbook, apply it, and export or import `.schema.json` files. Exporting a schema written in code leaves its hooks out.
Columns with an Excel list validation are generated as enum fields with the
list's values. Where a schema enum field and an Excel list cover the same
cell, the schema wins.

## How it works

```
UI thread                                   Web Worker
──────────────────────────────              ─────────────────────────────────────
React + Glide Data Grid (canvas)            WorkbookSession
SheetView: file cells + overlay  ◄─updates─   OoxmlPackage   unzip, keep every part
store: undo/redo, confirm dialogs ─edits──►   sheetScanner   cells + their XML offsets
                                              CalcEngine     HyperFormula, checked
                                                             against Excel's values
                                              writer         patch changed cells only
```

- **Reading** (`src/core/workbook.ts`, `sheetScanner.ts`): the package is
  unzipped and each worksheet's `<sheetData>` is scanned into compact typed
  arrays, recording where every row and cell sits in the original XML.
  Everything else in the package is left untouched.
- **Calculating** (`src/core/calc.ts`, `session.ts`): all sheets are loaded
  into HyperFormula, then every formula's result is compared with the value
  Excel saved. Formulas the engine can't parse or calculate (LET, LAMBDA,
  table references, spills…) are given Excel's value instead, so formulas
  that depend on them still calculate correctly.
- **Editing** (`session.ts`, `fill.ts`, `input.ts`): typed and pasted text is
  parsed the way Excel parses it (numbers, percentages, dates, formulas).
  Formula references are shifted with `@borgar/fx`. New functions get the
  `_xlfn.` prefix Excel needs in the file.
- **Saving** (`src/core/writer.ts`): only the `<c>` elements of changed cells
  are rewritten, plus cached results of formulas the engine recalculated.
  `calcChain.xml` is removed and the workbook is flagged to recalculate on
  open; pivot caches are flagged to refresh on open. A workbook with no edits
  is returned byte for byte.

## Measured (MacBook, Chrome)

| Workbook | Open (grid visible) | Recalculating live | Edit + recalc | Save |
|---|---|---|---|---|
| 100,000 rows × 12 columns, 300,000 formulas | 1.2 s | 4.4 s | 0.2–0.3 s | 1.5 s |

The worker holds about 0.8 GB for that file; the UI thread stays under 100 MB.

## Known limits

- Saved files have been checked with openpyxl, ExcelJS and this editor, not
  with Excel itself. Open a few saved files in Excel before relying on them,
  especially `pivot.xlsx`, whose pivot table was written by hand because
  ExcelJS can't create one.
- Not drawn: charts (placeholders only), mixed fonts inside a cell. Merges
  spanning several rows show only the top row.
- Conditional formatting is evaluated in the worker and redrawn after each
  edit: cell-value, formula, text, blank, error, date, top/bottom,
  above-average and duplicate rules, colour scales, data bars and icon sets,
  in priority order with "stop if true". Data bars don't draw a separate
  axis for negative values, icons are simplified drawings, and a sheet's
  rules stop being evaluated past 200,000 formula evaluations.
- Excel files name their built-in table styles without defining them, so
  those are drawn from each style family's pattern in the theme colours:
  close to Excel, not exact. Custom table styles are drawn as defined.
- Glide can freeze columns but not rows, so frozen rows are a second grid
  above the scrolling one. Arrow keys and sideways scrolling carry across;
  a drag or shift-selection can't span the two.
- Inserting or deleting rows and columns, and moving cells with cut and
  paste, rebuild the workbook in the worker: every reference to cells that
  move follows them (formulas on all sheets, names, merges, validation,
  conditional formatting, links, tables, pivot sources, charts, pictures and
  notes), references to deleted cells become #REF!, and changes that would
  split an array formula, a pivot table or a table's header are refused, as in
  Excel. Each one rebuilds the calculation engine, which takes a moment on
  large workbooks, and up to 20 can be undone. Cells moved to another sheet
  leave merges and tables behind, so such moves are refused.
- Sorting reorders rows as Excel does (formulas move as if copied; references
  from elsewhere don't follow). Filtering hides rows on screen only; the file
  keeps its own filter, as it was.
- Sheets can't be added, renamed or deleted.
- Pasted dates are read US-style (`3/4/2026` is 4 March) unless ISO.
- Validation formulas run in HyperFormula, so a rule using a function it
  lacks is treated as passing. OFFSET and INDIRECT are resolved by the
  editor for list sources only.
- HyperFormula is GPLv3 or commercially licensed.
- Cells are parsed into typed arrays and held in memory, so files much past
  a few million cells would need streaming parsing; the engine for live
  recalculation needs several times more memory than the cells themselves.
- Schema hooks run on the UI thread in the browser, and in Node through
  `ServerWorkbook`; the browser doesn't call a server itself. Inserting or deleting schema records
  moves cells within the table, so formulas outside the table that point
  into it are not adjusted (inserting rows from the grid's menu does adjust
  them).
- The generator's guesses are a starting point: tables need a row of typed
  headings, and a "Label | value" pair with a text value only counts when the
  label ends with a colon or the value is calculated.
