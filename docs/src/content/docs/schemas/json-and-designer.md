---
title: JSON and the designer
description: Schemas as plain data, generated from a workbook and edited in the editor's own designer, with your hooks kept in code.
sidebar:
  order: 6
---

Everything in a schema except its hooks is data, so a schema can be stored, versioned and edited without touching code.

## The designer

The editor's toolbar has a schema button that opens the schema designer. With a workbook open, it can:

- **Generate a schema from the workbook**: it finds the tables and their headings, and guesses each field's type from the values under them. A table needs a row of at least two typed headings (a formula's result is data, not a heading), and a label beside a value counts as a single cell when the label ends with a colon or the value is calculated.
- **Edit it**: the sheets, tables, fields, their types and checks, and the workbook's settings, with the workbook beside it.
- **Export it as JSON**, or import one.

People who know the workbooks best can describe them, and you keep the result with your app. Turn the button off with `features: { schemas: false }` where people only fill workbooks in.

## To and from JSON

```ts
import { schemaFromJSON, schemaToJSON } from '@ruled/sheet';

const json = schemaToJSON(priceList);          // a string
const schema = schemaFromJSON(json);           // a WorkbookSchema again
```

Regular expressions in labels are written as `{ "pattern": "…", "flags": "i" }`, and date bounds as ISO dates.

## Hooks stay in code

JSON can't hold functions, so a schema's hooks (its `validate`, `rules`, `onChange` and `beforeSave`) live in your app. Add them to a schema loaded from JSON before registering it:

```ts
const priceList = schemaFromJSON(await fetch('/schemas/price-list.json').then((r) => r.text()));
priceList.onChange = fillFromCatalogue;
priceList.sheets[0].tables![0].fields.find((f) => f.key === 'sku')!.validate = checkSku;
editor.registerSchema(priceList);
```

When someone edits a schema you registered in the designer, the hooks stay with it: they're matched to the edited schema by the workbook's sheet names, and the tables', fields' and cells' keys.
