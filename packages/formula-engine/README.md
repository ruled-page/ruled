# @ruled/formula

**Excel's calculation on its own.** Give it sheets of values and formulas,
and it works them out as Excel does and keeps the results up to date as cells
change. It needs no editor, no DOM and no server, so it runs in a browser, a
worker or Node. It is the engine [Ruled](https://ruled.page)'s spreadsheet
editor, [`@ruled/sheet`](https://www.npmjs.com/package/@ruled/sheet),
calculates with.

- **Documentation:** [ruled.page/formula](https://ruled.page/formula/introduction/)
- **Functions:** [the reference](https://ruled.page/reference/functions/), what each does and what it doesn't
- **Source:** [github.com/ruled-page/ruled](https://github.com/ruled-page/ruled)

## Install

```sh
npm install @ruled/formula
```

## Use

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

// Every value a change moves, each cell once.
book.set('Prices!B1', 10);
// [
//   { cell: 'Prices!B1', value: 10 },
//   { cell: 'Prices!B3', value: 13 },
// ]

// A formula worked out without a cell.
book.evaluate('=XLOOKUP("Cake",Prices!A1:A2,Prices!B1:B2)'); // 3
```

A string starting with `=` is a formula, written as in Excel in English.
Values are numbers, text, `true` or `false`, `null` for an empty cell, or an
error such as `{ error: '#DIV/0!' }`; dates are numbers, as Excel stores
them. See [cells, changes and formulas](https://ruled.page/formula/cells/)
for ranges, spills, names, sheets, rows and columns, and dates.

## What it calculates

More than 470 of Excel's functions, with Excel's rules for numbers, text,
dates, errors and empty cells, including:

- dynamic arrays that spill into the cells beside them;
- `LET`, `LAMBDA` and its helpers (`MAP`, `REDUCE`, `SCAN`, `BYROW` and the
  rest);
- defined names, references to other sheets, and 3-D references
  (`Jan:Mar!B2`);
- circular references, which calculate as 0, as Excel shows them.

Its answers are checked against Excel's, function by function, with
workbooks Excel itself graded.

## When to use it

- Working out a workbook's figures where no one looks at a grid: in a
  service, a script or a test.
- Checking what someone typed against a formula before it reaches a sheet.
- Running a pricing or quoting model written as Excel formulas, with inputs
  from a form.

To work out a saved .xlsx in Node, set its inputs and save it with its
values up to date, use `recalculate` from `@ruled/sheet/node`: see
[recalculating saved workbooks](https://ruled.page/formula/saved-workbooks/).
For people editing workbooks themselves, use the editor,
[`@ruled/sheet`](https://www.npmjs.com/package/@ruled/sheet).

## Licence

MIT.
