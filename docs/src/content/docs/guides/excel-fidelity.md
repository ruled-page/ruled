---
title: Excel fidelity
description: What the editor shows and edits, what it keeps without showing, and how a saved file compares with the one opened.
sidebar:
  order: 6
---

Workbook Editor opens the workbooks people already have and gives them back as Excel files. What it can show, it shows and lets people change; what it can't, it keeps, and writes back as it was.

## Worked out live

- **Formulas**, by the editor's own calculation engine, as cells change: Excel's function library, dynamic arrays and their spills, `LET`, `LAMBDA` with `MAP`, `REDUCE`, `SCAN` and the rest, names, tables' structured references, 3-D references, and iterative calculation where a workbook asks for it. On opening, its results are compared with the values Excel saved, and a formula that disagrees is reported.
- **Charts**, drawn from the current cells: column, bar, line, area, pie, doughnut, scatter, radar, bubble and stock charts, and 3-D charts drawn flat. People insert, move, resize, restyle and delete them as in Excel.
- **Pivot tables**, worked out again as their source changes, with Excel's PivotTable Fields pane, filters, subtotals, grand totals, compact, outline and tabular forms, and Show Values As. People insert their own.
- **Conditional formatting**: cell rules, top and bottom, averages, duplicates, colour scales, data bars and icon sets.

## Shown and edited

- Number formats, fonts, fills, borders, alignment, wrapping and indents; merged cells; column widths and row heights.
- Hidden, grouped and frozen rows and columns, with outlines.
- Sorting, filters (by value and by condition), data validation, names, hyperlinks, notes and threaded comments.
- Excel 365's checkboxes.
- Tables, with their styles and structured references.
- Sheets: added, renamed, moved, hidden and deleted, with references following them.

## Kept as it was

Macros (or removed, if you set `macros: 'remove'`), pictures and shapes, slicers and timelines, external links, embedded objects, custom XML and anything else in the file the editor doesn't draw. The editor's "What's in this file" report lists what a workbook holds that it keeps without showing.

## The saved file

The editor writes back only what changed: the cells and parts edits touched. Everything else is carried over from the original file, and a workbook nobody changed is saved byte for byte. Saved files are checked against Microsoft's Open XML SDK validator in the editor's own test suite.

Excel recalculates a saved workbook when it opens it. Formulas the editor can't calculate keep Excel's last result until then, and a save lists them as `outOfDate` so your app doesn't rely on them.

## Known limits

- Text in one cell with several fonts is drawn in one font; the file keeps them all.
- Charts don't draw data labels, trendlines, error bars, or gradient and picture fills. Excel's newer chart kinds (waterfall, treemap, histogram, funnel) and surface charts show as placeholders, and are kept.
- The Bessel functions match Excel to about eight significant digits.
- Pivot tables with grouped or calculated fields are shown as Excel saved them, and refreshed by Excel when it opens the file.
