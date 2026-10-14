---
title: Tables, cells and sheets
description: Where a schema's data sits, how tables grow, rules across records, and settings for the whole workbook.
sidebar:
  order: 3
---

## Tables

A table is a block of records with a heading for each field. `at` is its top-left cell, headings included.

People can't change the headings, or a single cell's label, as the schema finds the table by them. Ticking **Labels** in the toolbar hatches those cells, over whatever formatting the file gives them; the hatch is the editor's own, never saved.

```ts
{
  key: 'prices',
  at: 'A1',
  orientation: 'rows',
  records: { min: 1, max: 500 },
  fields: [/* … */],
}
```

| Property | What it does |
| --- | --- |
| `orientation` | `'rows'` (the default): one record per row, fields across. `'columns'`: one record per column, fields down, as in a forecast with a column per quarter. |
| `header` | Whether the first row (or column) holds the field labels. Default `true`. |
| `records.min`, `records.max` | How many records there may be. People add one by typing just past the last, or with the toolbar's insert and delete buttons. The toolbar's buttons move cells within the table, so formulas outside it that point into it aren't adjusted; inserting rows from the grid's menu adjusts them. |
| `records.labels` | Fixed records, named in the first column (or row): the number of records is fixed and the labels can't change, as in a budget with a row per category. |
| `corner` | The text expected in the top-left cell when the table has labels both ways. |
| `checks` | Rules across the records, as data. |
| `rules` | Rules across the records, in code. They can be async. |

### Rules as data

| Rule | Checks that |
| --- | --- |
| `{ rule: 'monotonic', field, direction }` | the field's values only go `'increasing'`, `'decreasing'`, `'nonIncreasing'` or `'nonDecreasing'`. |
| `{ rule: 'firstEquals', field, value }` | the first record's value is `value`. |
| `{ rule: 'sum', field, equals?, min?, max? }` | the field's total is, or is within, the given amounts. |
| `{ rule: 'compare', field, op, other }` | in every record, `field` compares with `other` this way: `<`, `<=`, `>`, `>=`, `=` or `<>`. |
| `{ rule: 'requiredIf', field, when, equals? }` | `field` has a value wherever `when` has one (or has the value `equals`). |
| `{ rule: 'uniqueTogether', field, fields }` | no two records share the same combination of values. |

Each takes `severity: 'warning'` too.

### Rules in code

A rule gets the table, with its records, and returns the issues it finds, each pinned to a record and field:

```ts
rules: [
  (table) =>
    table.records
      .filter((r) => r.values.discount !== null && Number(r.values.discount) > 0.3)
      .map((r) => ({ record: r.index, field: 'discount', message: 'Discounts over 30% need approval.', severity: 'warning' })),
],
```

Each record has its `index`, its `label` (for fixed records), its `values` by field key, and `ref(field)`, the address of one of its cells.

## Single cells

`cells` describes values that sit on their own, such as who approved a budget. Each is a field with a `ref` instead of a heading, and optionally a label that must be in another cell:

```ts
cells: [
  { key: 'approvedBy', ref: 'B8', label: { ref: 'A8', text: 'Approved by' }, type: 'text', required: true },
  { key: 'total', ref: 'B9', type: 'number', formula: 'require', readOnly: true },
],
```

## Sheets

| Property | What it does |
| --- | --- |
| `name` | The sheet's name, matched like a label. |
| `optional` | The workbook can lack the sheet and still match. |
| `outside` | Whether cells outside its tables and cells can be edited: `'readOnly'` (the default) or `'editable'`. |
| `validate` | A check across the whole sheet, in code. It can be async. |

## The whole workbook

| Property | What it does |
| --- | --- |
| `id`, `name`, `description` | Identify the schema, to your code and to people. |
| `version` | Stamped into saved files, so a file saved under another version says so. |
| `otherSheets` | Whether sheets the schema doesn't describe can be edited: `'editable'` (the default) or `'readOnly'`. |
| `saveWithErrors` | `'block'` (the default) refuses to save while errors remain; `'confirm'` asks first. |
| `hookTimeoutMs` | How long an async hook may take. Default 10,000. |
| `numbers` | How numbers are shown, copied and pasted (below). |
| `validate`, `onChange`, `beforeSave` | Code for the whole workbook: see [Async checks](/schemas/async-hooks/) and [Business logic](/schemas/business-logic/). |

### Numbers

Workbooks of prices or measurements often need numbers handled exactly:

```ts
numbers: {
  display: 'full',          // show every stored digit, not the cell's rounded format
  copy: 'raw',              // copy exact values, not the displayed text
  paste: 'strict',          // ask before pasting values that would be stored differently from how they read
  reviewBeforeSave: true,   // list every changed cell, old and new, before saving
}
```

`paste: 'strict'` asks about dates that could be day/month or month/day, numbers beyond 15 significant digits, and numbers with leading zeros. The same options can be set for every workbook in the editor's `numbers` option.
