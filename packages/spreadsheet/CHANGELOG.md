# Changelog

Notable changes to Ruled (`@ruled/sheet`, and `@ruled/plugin`, which is
released with it), newest first. Version numbers follow semantic versioning;
before 1.0.0, a minor version may change the API.

## Unreleased

- A new package, `@ruled/plugin`, holds Ruled's own plugins. The first,
  `@ruled/plugin/onedrive`, opens workbooks from OneDrive and SharePoint,
  saves them back unless they changed there, and saves copies to them.
- Plugins: `plugins` in the options (and the element's `options`) takes
  objects written against `@ruled/sheet/plugin`, which add commands to the
  toolbar, More and the cell menu, side panels, dialogs, and file sources.
  With a source, Open is a menu, a workbook opened from it can be saved back
  with "Save to <source>" (kept apart from the host's Save, with its own
  unsaved state and conflicts), and "Save a copy" can write to it. A plugin
  that throws is reported, by a message and a `pluginError` event, and never
  stops the editor. The code that runs plugins loads only for an editor that
  has some. `EditorHandle.origin` and `SavedWorkbook.origin` say where a
  workbook came from, `features.sourceSave` turns writing to sources off,
  and `savedToSource` events report writes. A plugin's context gives the
  editor's `limits.maxFileBytes`, so a source can refuse a file before
  downloading it, and `text()`, which gives a plugin's text in the editor's
  language for the DOM it puts in dialogs and panels.
- Both packages have a README on npm, and keywords to find them by.
- Typing straight after ⌘F (Ctrl+F) goes into the Find box. In Safari, soon after a
  workbook opened, it could go into the selected cell, as the box was still loading.
- Charts' data labels are drawn: values in their number format, category
  and series names and a pie's percentages, at the end, middle or base of a
  bar, beside a point or inside a slice, as the file places them.
- Excel 2016's waterfall and funnel charts are drawn from the cells, in place
  of placeholders: a waterfall's steps and totals in Excel's increase,
  decrease and total colours, a funnel's centred bars. They can be moved,
  resized and deleted, but not changed. Excel's newer charts could
  previously be opened in the chart dialog, which didn't apply to them.
- Password-protected workbooks open: the editor asks for the password, as
  Excel does, and saves the workbook encrypted with it again. They're
  decrypted in the browser, and no drafts of them are kept there. Node's
  `recalculate` takes a `password` too. Workbooks protected by a sensitivity
  label are still refused, now saying why.
- Shapes and text boxes are drawn, in place of placeholders: rectangles,
  ellipses, arrows, stars, lines and the other common kinds, with their
  fill, outline and text, as the file or its theme sets them. Groups and the
  rarer kinds stay placeholders.
- Charts' theme colours made lighter or darker (tint and shade) come out as
  Excel draws them.
- Sorting by cell colour, font colour or conditional formatting's icon, on
  top or at the bottom, from the Sort dialog or a column's filter menu.
- Filters by colour and by icon, and Excel's date filters by period (Today,
  This Month, Quarter 1, February and the rest), from a column's filter
  menu. A file's filters of these kinds open and change like any other.
- Pictures are drawn where they sit, cropped, turned and flipped as in Excel,
  in place of placeholders. A site with a Content Security Policy needs
  `img-src blob:` for them; without it, they stay placeholders.
- Sparklines are drawn: line, column and win/loss, with their markers, high,
  low, first, last and negative points and axis, from the cells as they are now.
- On a sheet that reads right to left, the outline bars are mirrored with the
  grid: the row bar is at the right, beside the row numbers.
- LET and LAMBDA formulas that other tools wrote without Excel's `_xlpm.` prefix
  on their names get it when the workbook is saved, as Excel showed them as
  #NAME? and Excel for the web refused the file. A workbook nobody changed is
  still saved byte for byte.

## 0.1.0 (2026-10-14)

The first release: a spreadsheet editor for web apps that opens Excel
workbooks, checks them against a schema, and saves them back without losing
what it doesn't change. Released with its calculation engine,
`@ruled/formula`, which also works on its own, both under the MIT licence.

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
  paste special), drag fill, Fill Series and custom lists, and find and
  replace with wildcards. AutoComplete finishes text from the entries above
  and below it in its column; hosts can turn it off with
  `features.autoComplete`.
- Numbers keep Excel's 15 significant digits: one typed or pasted with more
  keeps the first 15 and the rest become zeros, as Excel stores it, and
  strict pasting shows the value that will be stored.
- Cells move by dragging the selection's border, copy with Ctrl (Option on
  a Mac) held, and go in between others with Shift held, as in Excel.
- Inserting and deleting cells, rows and columns, from Excel's keys (Ctrl++
  and Ctrl+-) or the cell menu through Excel's Insert and Delete dialogs,
  with references following them.
- Whole rows and columns are selected by their numbers and letters, with
  dragging, Shift and Ctrl, the whole sheet from the corner, and Shift+Space,
  Ctrl+Space and Ctrl+A. A sheet runs to column XFD and row 1,048,576, as in
  Excel.
- Sorting, with Excel's Sort dialog for several levels, a custom list's
  order, case mattering, or left to right.
- Filter buttons on the headings of filters and tables (Alt+↓, and
  Ctrl+Shift+L or ⇧⌘F on a Mac): filtering by values, and by condition as
  Excel's Text, Number and Date Filters do (Custom AutoFilter, Top 10, above
  or below the average). A file's filter by a period of dates, colour or
  icon is kept as it is.
- Adding, renaming, deleting, moving, hiding and unhiding sheets.
- Notes and threaded comments, hyperlinks, defined names, and Excel 365
  checkboxes.
- Data validation as Excel applies it, with Excel's Data Validation dialog
  to set its rules, input messages and error alerts.
- Every change can be undone, and unsaved changes are kept as a draft on the
  device.
- The grid draws only the cells in view, on a canvas, however large the
  sheet. It shows frozen rows and columns, merged
  cells and sheets that read right to left as Excel does, and takes
  Excel's keys, typing in any language or through an input method, a pen,
  and on a phone a finger: it scrolls, the selection's handles select, and
  a long press opens the cell menu.

### Formatting

- A format bar: fonts, colours, borders in Excel's thirteen line styles and
  any colour, horizontal and vertical alignment, indent, wrapping, and
  number formats, with Excel's Custom list for format codes.
- Wrapped text is drawn as Excel draws it, and rows fit their contents as
  Excel's do. Double-clicking a column's or row's edge, or AutoFit, fits it.
- On a device without Calibri, Cambria, Arial, Times New Roman, Courier New
  or Georgia, text is drawn in a free font of the same widths (Carlito,
  Caladea, Arimo, Tinos, Cousine or Gelasio), fetched only then, so it fits
  its columns as in Excel.
- Column widths, row heights, hiding, frozen panes, merged cells and outline
  groups.
- Conditional formatting as Excel applies it, with its rules added, changed
  and removed: Highlight Cells and Top/Bottom rules, data bars, colour
  scales, icon sets, formula rules and the Rules Manager.

### Formulas

- Live recalculation by Ruled's own engine (`@ruled/formula`), with 477 of
  Excel's functions, 3-D references (Jan:Mar!B2), unions, LET, LAMBDA and
  its helpers, dynamic arrays that spill as Excel's do, and Ctrl+Shift+Enter
  arrays, its answers checked against Excel's. Formulas it can't calculate
  keep Excel's value and are marked out of date when their inputs change.
- Circular references calculate as 0 and warn, as in Excel; workbooks with
  iterative calculation on work them round.
- A banner says when a large workbook opens with live recalculation off,
  with a button to turn it on. `diagnostics` shows how long opening,
  calculating and the last edit take.
- Formula help: function suggestions, argument hints, coloured references
  and F4.

### Pivot tables and charts

- Pivot tables are worked out live from their source as it is edited, and
  on Refresh: rows, columns, values, report filters, subtotals and grand
  totals, in compact, outline or tabular form, with Show Values As. Insert
  PivotTable makes one, and the PivotTable Fields pane changes it; a changed
  pivot table is saved for Excel to refresh.
- Charts are drawn with Chart.js, loaded only for workbooks with charts:
  column, bar, line, area, pie, doughnut, scatter, radar, bubble and stock,
  and 3-D charts drawn flat. Insert Chart (Alt+F1) charts the selection, and
  Edit chart changes a chart's type, title, legend and data. Hosts can turn
  charts off.
- Charts, pictures and shapes can be selected, moved, resized and deleted,
  with the mouse or the keyboard.

### Checking with schemas

- Schemas describe what a workbook should hold, with field types, checks
  and code hooks. A designer generates and edits them.
- The cells a schema reads its headings and labels from can't be changed,
  and are hatched while "Mark schema labels" is ticked, over the file's own
  formatting. The theme token `gridLabelMark` colours the hatch.

### Embedding

- A React component and hooks (`useWorkbookEditor`, `useEditorState`), a
  standalone `<workbook-editor>` element and a mount function, each with an
  editor handle and events.
- Themes through CSS tokens, English and German and any language the host
  adds, right-to-left layouts, and phone-sized screens.
- React Server Components, as in Next.js's App Router: the React build is
  client code (`'use client'`), and `<Workbook>`, the editor in one
  component, takes props a Server Component can give it: the file's bytes,
  the options that are data, and a Server Action as `onSave`, which gets
  the saved workbook as plain data. `onSave` can answer `{ error }` or
  `{ conflict }` rather than throwing, as a server's errors don't reach the
  page as written.

### In Node

- `recalculate` from `@ruled/sheet/node` opens a saved workbook, sets the
  cells given, works out its formulas and saves it with its values up to
  date, with no browser.

### Accessibility

- Everything works from the keyboard. Screen readers hear each cell by its
  reference and value, and what actions did, and the focus stays visible in
  forced colours.
