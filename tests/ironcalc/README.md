# IronCalc's calculation workbooks

The workbooks here are the calculation tests of
[IronCalc](https://github.com/ironcalc/IronCalc) (`xlsx/tests/calc_tests`, at
commit `9d5fa0dc1d596fb9ace1f9f5491df34dc051ae50`), used under the Apache
License 2.0 in `LICENSE-Apache-2.0`. Nearly all were saved by Excel, so the
values in them are Excel's own answers.

`tests/ironcalc.test.ts` opens each one saved by Excel and compares every
formula's result with the value Excel saved, as the editor does when it opens
a file. `known.json` lists the cells where the engine still disagrees with
Excel, by workbook. The test fails on a disagreement that isn't listed, and on
a listed one that now agrees, so the list only shrinks. After fixing some,
`UPDATE_IRONCALC=1 npx vitest run tests/ironcalc.test.ts` writes the list
afresh.
