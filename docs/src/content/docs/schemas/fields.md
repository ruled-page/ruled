---
title: Fields and checks
description: Each field's type, limits, patterns and checks, and how its values reach your code.
sidebar:
  order: 2
---

A field is one column of a table (or one row, for a table that runs across), or a single cell. It has a `key` your code uses, a `label` that matches its heading in the workbook, a `type`, and the checks that apply.

```ts
{
  key: 'sku',
  label: 'SKU',
  type: 'text',
  required: true,
  unique: true,
  pattern: '^SKU-\\d{4}$',
  patternHint: 'like SKU-1234',
  description: 'The product code from the catalogue.',
}
```

## Types

| Type | Holds | In your code |
| --- | --- | --- |
| `text` | Any text | `string` |
| `number` | A number | `number` |
| `integer` | A whole number | `number` |
| `percent` | A percentage, stored as a fraction: 12.5% is `0.125` | `number` |
| `date` | A date, typed in the user's locale or as ISO | `Date`, at UTC midnight |
| `boolean` | TRUE or FALSE, or a checkbox | `boolean` |
| `enum` | One of `values`, typed in any case and offered as a drop-down | `string` |

A cell that can't be read as its type, such as text in a number field, is an error. An empty cell is `null`, and an Excel error value is `{ error: '#DIV/0!' }`.

## Checks

| Property | Checks that |
| --- | --- |
| `required` | the cell isn't empty. |
| `min`, `max` | a number, percentage (as a fraction) or date (as `'2026-01-31'`) is within the bounds, inclusive. |
| `decimals` | a number has no more decimal places; for percentages, of the percentage (12.5% has 1). |
| `minLength`, `maxLength` | text is within the lengths. |
| `pattern` | text matches the regular expression. `patternHint` describes it to people: "Should be like SKU-1234." |
| `values` | an `enum` field holds one of them. |
| `unique` | no two records hold the same value. |
| `formula` | cells `'allow'` formulas (the default), must hold one (`'require'`), or mustn't (`'forbid'`). |
| `readOnly` | nobody changes the cell. |
| `validate` | your own code agrees. It can be async: see [Async checks](/schemas/async-hooks/). |

## Labels

`label` is matched against the heading in the workbook. A string matches ignoring case and surrounding spaces; for anything looser, give a regular expression, or `{ pattern, flags }` where the schema has to be JSON:

```ts
{ key: 'price', label: /^unit price( \(.*\))?$/i, type: 'number' }
```

## Messages in several languages

Wherever a schema has text for people (`name`, `description`, `patternHint`, a check's message), it can be one string or one for each language. The editor shows the one that fits its language best:

```ts
description: { en: 'Unit price before tax.', de: 'Stückpreis vor Steuern.' }
```

A check can also write its message in the editor's language, which it's given as `ctx.language`.

## A field's own check in code

`validate` receives the value and a context, and returns nothing when the value is fine, or a message when it isn't:

```ts
{
  key: 'end',
  label: 'End date',
  type: 'date',
  validate: (end, { values }) => {
    const start = values?.start;
    if (start instanceof Date && end instanceof Date && end < start) return 'Ends before it starts.';
  },
}
```

The context has the sheet's name, the table's key, the record's index and all its values, the editor's language, and an `AbortSignal`. Return `{ message, severity: 'warning' }` for a warning.
