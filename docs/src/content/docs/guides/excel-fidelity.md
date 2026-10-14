---
title: Excel fidelity
description: What the editor shows and edits, what it keeps without showing, and how a saved file compares with the one opened.
sidebar:
  order: 6
---

Ruled opens the workbooks people already have and gives them back as Excel files. What it can show, it shows and lets people change; what it can't, it keeps, and writes back as it was.

## Worked out live

- **Formulas**, by the editor's own calculation engine, as cells change: Excel's function library, dynamic arrays and their spills, `LET`, `LAMBDA` with `MAP`, `REDUCE`, `SCAN` and the rest, names, tables' structured references, 3-D references, and iterative calculation where a workbook asks for it. On opening, its results are compared with the values Excel saved, and a formula that disagrees is reported.
- **Charts**, drawn from the current cells: column, bar, line, area, pie, doughnut, scatter, radar, bubble and stock charts, and 3-D charts drawn flat. People insert, move, resize, restyle and delete them as in Excel.
- **Pivot tables**, worked out again as their source changes, with Excel's PivotTable Fields pane, filters, subtotals, grand totals, compact, outline and tabular forms, and Show Values As. People insert their own.
- **Conditional formatting**: cell rules, top and bottom, averages, duplicates, colour scales, data bars and icon sets.
- **Pictures**, drawn where they sit; people move, resize and delete them as in Excel.
- **Sparklines**: line, column and win/loss, with their markers, high, low, first, last and negative points, and axis.

## Shown and edited

- Number formats, fonts, fills, borders, alignment, wrapping and indents; merged cells; column widths and row heights.
- Hidden, grouped and frozen rows and columns, with outlines.
- Sorting, filters (by value, by condition, by colour or icon, and by dates in a period), data validation, names, hyperlinks, notes and threaded comments.
- Excel 365's checkboxes.
- Tables, with their styles and structured references.
- Sheets: added, renamed, moved, hidden and deleted, with references following them.

## Kept as it was

Macros (or removed, if you set `macros: 'remove'`), shapes and text boxes (shown as placeholders), slicers and timelines, external links, embedded objects, custom XML and anything else in the file the editor doesn't draw. The editor's "What's in this file" report lists what a workbook holds that it keeps without showing.

## The saved file

The editor writes back only what changed: the cells and parts edits touched. Everything else is carried over from the original file, and a workbook nobody changed is saved byte for byte. Saved files are checked against Microsoft's Open XML SDK validator in the editor's own test suite.

Excel recalculates a saved workbook when it opens it. Formulas the editor can't calculate keep Excel's last result until then, and a save lists them as `outOfDate` so your app doesn't rely on them.

## Known limits

What the file holds is kept in each case; these are limits on how the editor shows and changes it.

### Drawing

- Text in one cell with several fonts is drawn in one font.
- A font the device lacks is drawn with a free font of the same widths for Calibri, Cambria, Arial, Times New Roman, Courier New and Georgia (see [Fonts](/start/install/#fonts)), and otherwise in the device's own font, a little wider or narrower than in Excel. Aptos, Excel's newest default, is one of those.
- Charts don't draw data labels, trendlines, error bars, or gradient and picture fills, and take their fonts and backgrounds from the editor's theme rather than the file. Excel's newer chart kinds (waterfall, treemap, histogram, funnel) and surface charts show as placeholders.
- Excel's built-in table styles are drawn from each style's pattern in the theme's colours: close to Excel, not exact. A file's own table styles are drawn as it defines them.
- Pictures are drawn cropped, turned and flipped as the file sets them, but without their borders, shadows and other effects. Pictures in formats browsers don't show, such as Windows metafiles (.emf, .wmf) and TIFFs, show as placeholders.
- Sparklines are drawn with their points evenly spaced, even on a date axis, and can't be changed here; they move with their cells.
- Conditional formatting's data bars start from zero, without a separate axis for negative values, and its icons are simpler drawings of Excel's. On a sheet with very many formula rules, rules past 200,000 cells' worth of formulas aren't applied.

### Calculating

- The Bessel functions match Excel to about eight significant digits.
- Iterative calculation starts from 0 rather than from the values saved in the file, so a formula that never settles, such as one counting itself up, comes to a different value from Excel's once it is worked out again.
- An array formula entered in Excel with Ctrl+Shift+Enter across several cells recalculates here, but can't be changed or cleared, and new ones can't be entered. Data tables (`{=TABLE(…)}`) keep Excel's values. A dynamic array of more than a million cells isn't spilled.
- A data validation rule whose formula uses a function the editor can't calculate lets every value through.
- Pivot tables with grouped or calculated fields are shown as Excel saved them, and refreshed by Excel when it opens the file.

### Editing

- Excel's Point mode isn't offered: while a formula is typed in a cell, the arrow keys move the caret, and clicking another cell enters the formula. Type the references instead.
- A changed note keeps its text but not its formatting (bold, colours), and a changed threaded comment loses its @mentions.
- Only a formula that is a `HYPERLINK` call is a link; one with `HYPERLINK` inside a larger formula isn't.
- Cells holding merged cells or tables can't be moved to another sheet. Pictures and charts move with inserted and deleted rows and columns, but not with inserted or deleted cells. Up to 20 inserts, deletes and moves can be undone, fewer in very large workbooks.
- Only one filter is active on a sheet at a time: filtering another table or range replaces it. A file's filters by groups of dates (years, months or days ticked in Excel's list) are kept, but can't be changed. Filtering by colour looks at a cell's own fill or text colour and conditional formatting's, not a table style's.
- Excel's Auto Outline, outline styles and outline settings aren't offered; a file's settings are kept.
- In a workbook with slicers or timelines, sheets holding pivot tables or tables can't be deleted. A pivot table whose source sheet is deleted keeps its values and isn't refreshed.
