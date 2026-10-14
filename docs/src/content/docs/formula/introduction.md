---
title: The formula engine
description: Excel's calculation with no editor, in the browser or Node, from sheets of values and formulas.
sidebar:
  order: 1
---

The editor calculates with its own engine, and the engine works on its own too, as `@ruled/formula`. Give it sheets of values and formulas, and it works them out as Excel does and keeps the results up to date as cells change. It needs no editor, no DOM and no server, so it runs in a browser, a worker or Node.

```ts
import { createWorkbook } from '@ruled/formula';

const book = createWorkbook({
  sheets: {
    Prices: [
      ['Tea', 4.5],
      ['Cake', 3],
      ['Total', '=SUM(B1:B2)'],
    ],
  },
});

book.get('Prices!B3'); // 7.5
book.set('Prices!B1', 10); // [{ cell: 'Prices!B1', value: 10 }, { cell: 'Prices!B3', value: 13 }]
book.evaluate('=XLOOKUP("Cake",Prices!A1:A2,Prices!B1:B2)'); // 3
```

## What it calculates

The same as the editor: more than 470 of Excel's functions (the [functions reference](/reference/functions/) lists them), with Excel's rules for numbers, text, dates, errors and empty cells. That includes:

- dynamic arrays that spill into the cells beside them, as in Excel today;
- LET, LAMBDA and its helpers (MAP, REDUCE, SCAN, BYROW and the rest);
- defined names, references to other sheets, and 3-D references (`Jan:Mar!B2`);
- circular references, which calculate as 0, as Excel shows them.

Its answers are checked against Excel's, function by function, with workbooks Excel itself graded.

## When to use it

- Working out a workbook's figures where there is no one to look at a grid: in a service, a script or a test.
- Checking what a user typed against a formula before it reaches a sheet.
- Running a pricing or quoting model written as Excel formulas, with inputs from a form.

To work out a saved .xlsx in Node, setting its inputs and saving it with its values up to date, see [recalculating saved workbooks](/formula/saved-workbooks/). When people edit the workbook themselves, use the editor, `@ruled/sheet`, which calculates with this engine in a worker.

## Install

```sh
npm install @ruled/formula
```

Next: [Cells, changes and formulas](/formula/cells/) shows what a workbook can do.
