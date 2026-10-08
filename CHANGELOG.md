# Changelog

Notable changes to workbook-editor, newest first. Version numbers follow
semantic versioning; before 1.0.0, a minor version may change the API.

## Unreleased

- Dynamic array formulas, typed or from the file, spill their values into
  the cells around them and recalculate live. A spill is outlined while one
  of its cells is selected, a blocked one shows #SPILL!, and spills are
  saved as Excel saves them.

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
