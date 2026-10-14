# Backlog

Work that is deliberately not being done yet, with what it would take. Ruled
needs no server, and nothing here should add one.

## Needs a person

- **Screen-reader testing.** The editor works from the keyboard and passes
  automated accessibility checks in Chromium, Firefox and WebKit, but nobody
  has used it with NVDA, JAWS or VoiceOver. A tester should go through
  opening a file, moving around the grid, editing, the checks panel and the
  dialogs with each.
- **Checks in a licensed desktop Excel.** Excel for the web has opened the
  editor's saved pivot tables without repairing them; desktop Excel should
  open the pivot test workbooks and the demo's kitchen-sink sample, saved
  after edits, too.
- **Details checked only against Excel's documentation**, to try in Excel:
  that a saved checkbox shows as one; how Top 10 by percent rounds (down
  here, keeping at least one); the width of an indent level (9 pixels
  here); and that a range moved with Shift+drag keeps its size when it only
  changes the order of its cells.
- **Typing with a real input method** (Japanese, Chinese, Korean) into a
  selected cell. The end-to-end tests compose text through Chromium's
  DevTools protocol; nobody has tried a real input method, or one in
  Firefox or Safari.
- **Text set to align left or right on a right-to-left sheet.** The grid
  swaps the two, so text set to the left sits at the cell's right, where
  it starts; to check in Excel.

## Excel features left out for now

- **Slicers and timelines.** Kept in the file; not shown, and pivot tables
  and tables aren't filtered by them here.
- **Pictures and shapes.** Kept in the file and shown as placeholder boxes
  where they sit, which can be selected, moved, resized and deleted.
- **Surface charts and Excel 2016's newer kinds** (waterfall, histogram,
  treemap, funnel and the rest). Kept in the file and shown as placeholders;
  the other kinds are drawn, and a host can turn drawing off
  (`features.charts`).
- **Sparklines.** Kept in the file and moved with their cells; not drawn.
- **.xls, .xlsb and .ods files.** Refused with a message. A converter such
  as SheetJS could open them as .xlsx, losing charts and pivot tables.
- **Printing and PDF.** Print settings are kept in the file; nothing prints.
  A print view could use the browser's own printing.
- **What-if data tables** (`{=TABLE(…)}`). Their cells keep Excel's saved
  values.
- **Filtering by colour or icon, Excel's date filters by period** (This
  Month, Q1 and the rest), and **sorting by colour or icon.** A file's
  filters of these kinds are kept as they are; the editor doesn't offer
  them.
- **Outline bars on right-to-left sheets.** A sheet that reads right to
  left is drawn from right to left, but its group bars stay at the left,
  away from its row numbers at the right.
- **Several people editing at once.** Live editing needs a server. A host
  can still turn down a save that conflicts (`SaveConflictError`).
- **Password-protected and sensitivity-labelled files.** Refused with a
  message. Password-protected files could be decrypted in the browser;
  sensitivity-labelled ones can only be decrypted through Microsoft's
  service.

## Could be smaller

- **The standalone build carries its calculation worker inside it**: 184 kB
  of its 441 kB gzipped (`packages/spreadsheet/src/editor/standalone.ts`
  imports it with `?worker&inline`), so the one script works when served
  from another domain, as browsers start workers only from the page's own.
  Loading the worker from a file beside the script, as the React build does,
  would bring the script to about 260 kB, with the worker downloading
  alongside it; the total stays the same. Hosts would then serve both from
  their own domain, or the editor would start the worker from a script
  fetched across domains.

## Known issues

Small gaps found along the way.

- **Arrays are held to a million values**, so `=ROWS(WRAPROWS({1,2},2000000))`
  is #NUM! where Excel gives 1. Excel only refuses an array too big to
  hold; the engine's limit (`MAX_ARRAY_CELLS` in
  `packages/formula-engine/src/sizeLimits.ts`) keeps a stray formula from
  taking the page's memory.
- **LET formulas written by other tools without Excel's `_xlpm.` prefix**
  on their names show #NAME? in Excel, and Excel for the web refuses the
  file; the editor works them out as intended. Saving keeps formulas no one
  edited as the file has them, so such a file stays as it was; formulas
  typed or edited here get the prefix. Adding it on saving would mend such
  files, at the cost of rewriting cells no one changed.
