---
title: Cells, changes and formulas
description: Reading and writing cells by A1 reference, the changes a write makes, spills, names, sheets and dates.
sidebar:
  order: 2
---

A workbook is made from its sheets, in order, each as rows of values and formulas from A1. A string starting with `=` is a formula, typed as it would be in Excel in English: commas between arguments and a dot for decimals.

```ts
import { createWorkbook } from '@ruled/formula';

const book = createWorkbook({
  sheets: {
    Prices: [
      ['Tea', 4.5],
      ['Cake', 3],
      ['Total', '=SUM(B1:B2)'],
    ],
    'Q1 Sales': [['=SEQUENCE(3)', null, '=Prices!B3*2']],
  },
  names: { Rate: '=Prices!$B$1' },
});
```

With no sheets, a workbook has one, `Sheet1`. The examples below each start from this workbook as it is made.

## References

Cells are named as Excel names them: `"Prices!B2"`, or `"'Q1 Sales'!A1"` for a sheet whose name needs quotes. A reference without a sheet is to the first. Sheet names are matched whatever their case, as in Excel.

| Method | What it gives |
| --- | --- |
| `get('Prices!B3')` | One cell's value: `7.5` |
| `getRange("'Q1 Sales'!A1:A4")` | A range's values as rows: `[[1], [2], [3], [null]]` |
| `formula('Prices!B3')` | The formula a cell holds as it was given, `'=SUM(B1:B2)'`, or `undefined` |
| `sheets` | The sheets' names, in order |

A value is a number, text, `true` or `false`, `null` for an empty cell, or an error as an object, such as `{ error: '#DIV/0!' }`. Dates are numbers, as Excel stores them (see [Dates](#dates)).

## Changing cells

`set` puts a value or formula in a cell and returns every value that changed as a result, each cell once, with its value now:

```ts
book.set('Prices!B1', 10);
// [
//   { cell: 'Prices!B1', value: 10 },
//   { cell: 'Prices!B3', value: 13 },
//   { cell: "'Q1 Sales'!C1", value: 26 },
// ]
```

It also takes several cells at once, or rows of them from a cell:

```ts
book.set({ 'Prices!B1': 1, 'Prices!B2': 2 });
book.set('Prices!D1', [
  [1, 2],
  [3, '=D1+E1+D2'],
]);
```

`null` or `undefined` empties a cell.

## Spills

A formula whose result is an array spills it into the cells beside it, as Excel does today. Above, `=SEQUENCE(3)` in `'Q1 Sales'!A1` fills A1:A3. A spill:

- shows `#SPILL!` in its formula's cell while something is in its way, and spills again when that is cleared;
- grows and shrinks as its result does, `set` reporting the cells it reaches and the cells it leaves (as `null`).

```ts
book.set("'Q1 Sales'!A2", 'x');
// [
//   { cell: "'Q1 Sales'!A2", value: 'x' },
//   { cell: "'Q1 Sales'!A3", value: null },
//   { cell: "'Q1 Sales'!A1", value: { error: '#SPILL!' } },
// ]
```

## Formulas without a cell

`evaluate` works out a formula without putting it anywhere. References without a sheet are to the sheet named, or the first; an array result comes back as rows.

```ts
book.evaluate('=Rate*2'); // 9
book.evaluate('=B1+1', 'Prices'); // 5.5
book.evaluate('=SEQUENCE(2,2)'); // [[1, 2], [3, 4]]
book.evaluate('=MAP(SEQUENCE(3),LAMBDA(v,v*2))'); // [[2], [4], [6]]
```

## Names and sheets

Names are given when the workbook is made, as a reference or formula, and changed with `setName`, which returns the values that changed. `null` removes one, and formulas using it show `#NAME?`.

```ts
book.set('Prices!C1', '=Rate*10');
book.setName('Rate', 'Prices!$B$2'); // [{ cell: 'Prices!C1', value: 30 }]
```

`addSheet` adds a sheet after the others. A formula can name a sheet before it exists: it shows `#REF!` until the sheet is added, and then reads it.

```ts
book.set('Prices!A5', '=Later!A1+1'); // [{ cell: 'Prices!A5', value: { error: '#REF!' } }]
book.addSheet('Later', [[41]]); // [{ cell: 'Prices!A5', value: 42 }]
```

## Dates

Excel stores a date as a number: the days since 1900 began, with the time of day as a fraction. A JavaScript `Date` given to a cell is turned into one, read in UTC, and dates come back as numbers. `dateToSerial` and `serialToDate` convert between the two.

```ts
import { createWorkbook, serialToDate } from '@ruled/formula';

const book = createWorkbook({ sheets: { S: [[new Date('2026-10-13T12:00:00Z'), '=TEXT(A1,"yyyy-mm-dd hh:mm")']] } });
book.get('A1'); // 46308.5
book.get('B1'); // '2026-10-13 12:00'
serialToDate(46308.5); // 2026-10-13T12:00:00.000Z
```

A workbook made with `date1904: true` counts from 1904, as workbooks from older Macs do; pass the same flag to the helpers.

## Mistakes

A reference to a sheet that isn't there, a reference that isn't one (`"1A"`), a range given to `get`, a sheet name Excel doesn't allow or already has, and a name that can't be defined each throw an `Error` saying what is wrong. Mistakes in formulas don't throw: they give Excel's errors, such as `#NAME?` for a function Excel doesn't have.

## What it leaves out

- Sheets can be added but not renamed, moved or deleted, and rows and columns can't be inserted or deleted; references are as written.
- Tables and their structured references (`Sales[Amount]`), links to other workbooks, and CELL's information about formats and files.
- Text starting with `=`, which is always read as a formula.
- Opening and saving `.xlsx` files: the engine works on values and formulas. The editor, `@ruled/sheet`, reads and writes files.
