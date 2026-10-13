---
title: Recalculating saved workbooks
description: Open a saved .xlsx in Node, set its input cells, work out its formulas, and save it again with its values up to date.
sidebar:
  order: 3
---

`recalculate` works on a whole workbook as Excel saved it: it opens the file, sets the cells you give it, works out every formula with the [formula engine](/formula/introduction/), and saves the file again. Whatever reads the file next, Excel or a program that calculates nothing, finds the values up to date. It runs in Node, with no browser, and comes with the editor's package, which reads and writes the files:

```ts
import { readFileSync, writeFileSync } from 'node:fs';
import { recalculate } from '@ruled/sheet/node';

const result = await recalculate(readFileSync('quote-template.xlsx'), {
  set: { 'Inputs!B2': 120, 'Inputs!B3': 'Express', 'Inputs!B4': new Date('2026-11-01') },
});

result.get('Summary!B10'); // the total, worked out
writeFileSync('quote.xlsx', result.bytes);
```

## Setting cells

`set` takes cells by reference, with their sheet (`"Inputs!B2"`, `"'Q1 Sales'!A1"`) or on the first sheet (`"B2"`). Each is stored as given:

| Given | Stored as |
| --- | --- |
| A number | A number |
| Text | Text, even text that looks like a number, such as `"007"` |
| Text starting with `=` | A formula, typed as in Excel in English |
| `true` or `false` | TRUE or FALSE |
| A `Date` | A date, read in UTC, with a date format |
| `null` or `undefined` | Nothing: the cell is emptied |

A reference to a sheet that isn't there, or to no cell, throws an `Error`, and so does a cell the workbook won't let be changed, such as a locked cell on a protected sheet.

## The result

| Member | What it is |
| --- | --- |
| `bytes` | The workbook saved again, with each formula's value as worked out. Everything else in the file is kept as it was, and Excel works the formulas out once more when it opens it. |
| `sheets` | The sheets' names, in order. |
| `get(cell)`, `getRange(range)` | Values, as [a workbook on its own](/formula/cells/) gives them. |
| `outOfDate` | Formulas that keep the value the file saved though a cell they read was set: ones the engine can't calculate, and ones in `differs`. |
| `differs` | Formulas whose value worked out here differs from the one the file saved. |

## Excel's saved values

When the engine's value for a formula differs from the one the file saved, the formula keeps the saved value, and is listed in `differs`: a workbook saved by Excel holds Excel's own answers. A file saved by a program that doesn't calculate may hold values that are out of date instead; `ignoreSavedValues: true` gives every formula the engine's value.

A formula the engine can't calculate, such as one linking to another workbook, keeps its saved value either way. The editor's [known limits](/guides/excel-fidelity/#known-limits) say what the engine leaves out.
