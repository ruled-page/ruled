# Changelog

Notable changes to workbook-editor, newest first. Version numbers follow
semantic versioning; before 1.0.0, a minor version may change the API.

## Unreleased

- Cells move by dragging the selection's border, and copy with Ctrl (Option
  on a Mac) held, as in Excel; dropping them on cells that hold something
  asks first.
- A drag that ends off the grid no longer swallows the next click.
- Shift+dragging the selection's border moves the cells in between others,
  as in Excel, with a line marking where they'll go.
- Insert and delete cells, moving the others down or right (or up or left),
  or whole rows or columns, from Excel's keys (Ctrl++ and Ctrl+-) or the
  cell menu, through Excel's Insert and Delete dialogs.
- AutoComplete: text typed into a cell, or the formula bar, is finished from
  the entries above and below it in its column, as in Excel. Hosts can turn
  it off with `features.autoComplete`.
- Excel's Sort dialog, from "Custom sort…" in the cell menu: sort by
  several levels, by a custom list's order, with case mattering, or left to
  right.
- Dynamic array formulas, typed or from the file, spill their values into
  the cells around them and recalculate live. A spill is outlined while one
  of its cells is selected, a blocked one shows #SPILL!, and spills are
  saved as Excel saves them.
- Array formulas entered with Ctrl+Shift+Enter recalculate live, filling the
  range they were entered in, and show `{=formula}` in the formula bar.
- LET formulas recalculate live.
- Whole rows and columns are selected by their numbers and letters, with
  dragging, Shift and Ctrl, the whole sheet from the corner, and Shift+Space,
  Ctrl+Space and Ctrl+A.
- Double-clicking a column's or row's edge fits it to its contents, as does
  AutoFit in the cell menu. Pressing an edge without dragging no longer
  changes its size.
- A banner says when a large workbook opens with live recalculation off,
  with a button to turn it on.
- A sheet runs to column XFD and row 1,048,576, as in Excel, more rows and
  columns appearing as it is scrolled; going to a cell beyond them brings
  them with it.
- The name box lists what each name refers to, and says a name or reference
  on a hidden sheet can't be gone to instead of showing the hidden sheet.
- Drop-down lists draw the theme's arrow, the toolbar folds its labels only
  as far as it must to fit (and never cuts off a button), and the formula
  bar uses the interface's font.
- The demo's themes show more of what theming can do: Indigo, Paper, Dark
  and High contrast.
- The editor calculates with its own engine in place of HyperFormula: every
  function it offers, 3-D references (Jan:Mar!B2), unions, FORMULATEXT, CELL,
  LAMBDA and its helpers, dynamic arrays and Ctrl+Shift+Enter arrays, its
  answers checked against Excel's. The library's worker is 156 kB gzipped
  where it was 297 kB, the largest sample calculates in about a third of the
  time with about half the memory, and the editor no longer contains
  GPL-licensed code. `diagnostics` shows in the status bar how long opening,
  calculating and the last edit take, and `npm run engine:bench` measures
  them.
- Copying numbers as HTML gives them to 15 significant digits, as Excel does.
- Circular references calculate as 0, as in Excel; making one warns, and
  the status bar names where one closes until it is broken.
- Workbooks with iterative calculation on work their circular references
  round, up to the file's number of passes or until nothing moves by more
  than its maximum change, without warning of them.
- #CALC!, #SPILL! and the other errors Excel saves as #VALUE! for older
  readers are read as the errors they are, so checks on open compare the
  real error.

## 0.1.0 (2026-10-08)

The first release: an Excel workbook editor for the browser that saves
workbooks back without losing what it doesn't change.

### Opening and saving

- Opens .xlsx and .xlsm workbooks, and .csv, .tsv and .txt files as
  one-sheet workbooks read in the user's locale.
- Saves only what was edited. Charts, pivot tables, macros, pictures and
  every other part of the file stay byte-for-byte as they were.
- Hands the host the saved workbook, its values, the changes made and the
  cells whose saved values are out of date. Downloads a copy, or the sheet
  shown as CSV.
- Refuses damaged and hostile files and files too large for a browser tab,
  and keeps or removes macros as the host chooses.

### Editing

- Typing, copying and pasting (with Excel's and Google Sheets' formats, and
  paste special), drag fill, Fill Series and custom lists, find and replace
  with wildcards, sorting, and filtering by values.
- Inserting and deleting rows and columns, and cutting to move cells, with
  references following them.
- Adding, renaming, deleting, moving, hiding and unhiding sheets.
- Formatting from a format bar: fonts, colours, borders, alignment, wrapping
  and number formats.
- Column widths, row heights, hiding, frozen panes, merged cells and outline
  groups.
- Notes and threaded comments, hyperlinks and defined names.
- Every change can be undone, and unsaved changes are kept as a draft on the
  device.

### Formulas

- Live recalculation with about 465 of Excel's functions. Formulas the
  engine can't calculate keep Excel's value and are marked out of date when
  their inputs change.
- Formula help: function suggestions, argument hints, coloured references
  and F4.
- Data validation and conditional formatting, as Excel applies them.

### Checking with schemas

- Schemas describe what a workbook should hold, with field types, checks
  and code hooks. A designer generates and edits them.

### Embedding

- A React component, a standalone `<workbook-editor>` element and a mount
  function, each with an editor handle and events.
- Themes through CSS tokens, English and German and any language the host
  adds, right-to-left layouts, and phone-sized screens.
- Charts drawn with Chart.js, loaded only for workbooks with charts. Hosts
  can turn them off.

### Accessibility

- Everything works from the keyboard. Screen readers hear each cell by its
  reference and value, and what actions did, and the focus stays visible in
  forced colours.
