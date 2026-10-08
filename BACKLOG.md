# Backlog

Work that is deliberately not being done yet, with what it would take. The
editor is a client-side proof of concept; nothing here should add a server.

## Needs a decision first

- **The calculation engine's licence.** HyperFormula is used under its GPLv3
  licence key, which a closed-source product can't do. Before production:
  buy its commercial licence and pass the key in through an option, or
  replace the engine. The key is set in `src/core/calc.ts` and
  `src/core/engineLimits.ts`.

## Needs a person

- **Screen-reader testing.** The editor works from the keyboard and passes
  automated accessibility checks in Chromium, Firefox and WebKit, but nobody
  has used it with NVDA, JAWS or VoiceOver. A tester should go through
  opening a file, moving around the grid, editing, the checks panel and the
  dialogs with each.

## Excel features left out for now

- **Refreshing pivot tables, and slicers.** Pivot tables are kept and shown,
  and Excel refreshes them on opening when their source data changed.
  Refreshing them in the editor would need a pivot engine.
- **Pictures and shapes.** Kept in the file and moved with their cells, and
  shown as placeholder boxes where they sit. Charts are drawn, and a host can
  turn that off (`features.charts`).
- **Sparklines.** Kept in the file and moved with their cells; not drawn.
- **.xls, .xlsb and .ods files.** Refused with a message. A converter such
  as SheetJS could open them as .xlsx, losing charts and pivot tables.
- **Printing and PDF.** Print settings are kept in the file; nothing prints.
  A print view could use the browser's own printing.
- **What-if data tables and iterative calculation.** Their cells keep
  Excel's saved values. The engine has no iterative calculation.
- **Several people editing at once.** Live editing needs a server. A host
  can still turn down a save that conflicts (`SaveConflictError`).
- **Password-protected and sensitivity-labelled files.** Refused with a
  message. Password-protected files could be decrypted in the browser;
  sensitivity-labelled ones can only be decrypted through Microsoft's
  service.

## Known issues

Small gaps found along the way, none of which loses data or shows a wrong
value without saying so.

- **A name qualified with its sheet** (`Prices!Rate`, for a name that belongs
  to one sheet) isn't calculated: the engine can't read the form. Such
  formulas keep Excel's saved value and are marked out of date when their
  inputs change. The engine would need each sheet's names under names of
  their own, kept in step with renames.
- **Dragging the fill handle doesn't copy links**, as copying and pasting
  does. `copyLinks` in `src/core/links.ts` tiles them; filling upwards or
  leftwards needs the tiles lined up from the far end.
- **The function list doesn't offer the workbook's names or tables** as it
  types, as Excel's does. They work in formulas; only the suggestions miss
  them (`src/core/formulaAssist.ts`).
- **The report's "Notes and comments" count** is what the file had when it
  was opened, not what it has now.
- **Deleting the rows that hold a sheet's last notes** leaves its comments
  part with no comments in it. It is valid, but Excel removes the part.
- **The package lists the code it bundles as dependencies.** Everything but
  React is bundled into the library builds, yet `dependencies` names
  HyperFormula, Glide and the rest, so a host app installs them for nothing.
  They can become devDependencies once the type declarations stop importing
  from them, for example by bundling the declarations into one file.
- **German short month names** come out as numfmt writes them ("März")
  where German Excel writes "Mrz".
- **A long list in a cell's dropdown** scrolls in a region that axe, in
  Firefox and WebKit, reports as not reachable from the keyboard; the arrow
  keys do reach every value.
- **HSTACK and VSTACK** could, in a case seen once and not reproduced, be
  given the wrong size by the engine, which then stops calculating the
  workbook.
- **Some of the saved XML is rewritten with regular expressions**
  (`src/core/writer.ts`, `src/core/restructure.ts`). They hold for what Excel
  and other common tools write; unusual but valid XML, such as a `>` inside
  an attribute value, could defeat them.
