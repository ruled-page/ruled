# Workbooks saved by Excel

Calculation workbooks, nearly all saved by Excel, so the values cached in them
are Excel's own answers.

`tests/excel-saved.test.ts` opens each one saved by Excel and compares every
formula's result with the value Excel saved, as the editor does when it opens
a file. `known.json` lists the cells where the engine still disagrees with
Excel, by workbook. The test fails on a disagreement that isn't listed, and on
a listed one that now agrees, so the list only shrinks. After fixing some,
`UPDATE_EXCEL_SAVED=1 npx vitest run tests/excel-saved.test.ts` writes the list
afresh.

## Source

The workbooks come from the `xlsx/tests/calc_tests` folder of
[IronCalc](https://github.com/ironcalc/IronCalc), at commit
`9d5fa0dc1d596fb9ace1f9f5491df34dc051ae50`, and are used under the Apache
License 2.0 in `LICENSE-Apache-2.0`. Only the workbooks are taken; none of its
code.
