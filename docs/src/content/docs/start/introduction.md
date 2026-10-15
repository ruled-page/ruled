---
title: Introduction
description: What Ruled is, what it's for, and how it fits into your app.
sidebar:
  order: 1
---

Ruled is a spreadsheet editor you put in your own web app. It opens the Excel files people already have (`.xlsx`, `.xlsm` and CSV), lets them work in them as they would in Excel, and hands the workbook back to your app as an `.xlsx` file and as values in the shape you asked for.

What sets it apart is that it can **know what a workbook is meant to hold**. You describe the workbooks you expect in a [schema](/schemas/overview/): the sheets, the tables, each field's type and limits, and the rules across them. A workbook that matches is checked as people type, cell by cell, and your own code runs alongside: [checks that call your services](/schemas/async-hooks/), and [business logic](/schemas/business-logic/) that fills cells in from them.

## What it's for

- **Collecting data in spreadsheets**: suppliers sending price lists, teams returning budgets and forecasts, anyone filling in a template. People keep using a spreadsheet; you get values you can trust, checked before they reach you.
- **Reviewing and correcting files** inside your app instead of emailing them back and forth.
- **Editing workbooks in the browser** that must go back to Excel unchanged in everything nobody touched.

## How it works

The editor is a React component, and a `<workbook-editor>` element for every other kind of page. The workbook itself lives in a **web worker**: reading the file, calculating formulas and writing it back happen off the page's main thread, so the page stays responsive with large workbooks.

```
your app ──open(bytes)──▶ editor ──▶ worker: reads the .xlsx, calculates, checks
            ◀──onSave(file)── editor ◀── worker: writes the .xlsx back
```

When the workbook is saved, your app receives:

- the `.xlsx` file, with everything the editor doesn't show kept as it was;
- the values in your schema's shape: each table's records by field, each named cell;
- each changed cell, with its value before and after, and who did what, when.

## What's in it

- **Formulas**, calculated live by the editor's own engine, including dynamic arrays and `LAMBDA`.
- **Charts and pivot tables**, drawn from the current cells and editable as in Excel; pictures, shapes, text boxes and sparklines drawn where they sit.
- **Formatting, sorting, filters, data validation, names, hyperlinks, notes and comments**, with sorting and filtering by colour, icon and date period.
- **Password-protected workbooks**, opened with their password and saved encrypted with it again.
- **Schemas** with declarative checks, async hooks, business logic and a visual designer.
- **Theming** through design tokens, **languages** (English and German built in), keyboard use and screen reader support.

## Next

- See [what it can do](/start/capabilities/), feature by feature.
- [Install it](/start/install/), then add it to a [React](/start/react/), [plain HTML](/start/web-component/) or [Angular](/start/angular/) app.
- Read about [schemas](/schemas/overview/), the reason to use it.
- Try the [live examples](/examples/price-list/).
