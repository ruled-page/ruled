import { describe, expect, it } from 'vitest';
import type { RawContent } from '../src/api';
import { engineFormula } from '../src/formula';
import { nativeEngine } from '../src/native';

const NUM = { error: '#NUM!' };
const NA = { error: '#N/A' };
const DIV0 = { error: '#DIV/0!' };
const VALUE = { error: '#VALUE!' };

/** Each formula worked out on sheet S (holding `rows`, beside a sheet T), against what it should give. */
function check(rows: RawContent[][], cases: [string, unknown][]) {
  const e = nativeEngine.create([{ name: 'S', rows }, { name: 'T', rows: [[1]] }], [], false);
  const got = cases.map(([f]) => e.calculate(`=${f}`, 0));
  e.destroy();
  // Numbers to six decimal places, as Microsoft's examples give them.
  const near = (v: unknown) => (typeof v === 'number' ? Math.round(v * 1e6) / 1e6 : v);
  expect(cases.map(([f], i) => [f, near(got[i])])).toEqual(cases.map(([f, want]) => [f, near(want)]));
}

describe('exact lookups over long lists', () => {
  // 100 keys: numbers (0.3 among them), text, TRUE and FALSE, an empty cell, an error, and "Item 50" three times more.
  const keys: RawContent[] = Array.from({ length: 100 }, (_, i) =>
    i === 49 ? 0.3 : i < 50 ? i + 1 : i < 90 ? `Item ${i}` : i < 95 ? i % 2 === 0 : i === 95 ? null : i === 96 ? '#DIV/0!' : 'Item 50',
  );
  const rows = keys.map((k, i) => [k, (i + 1) * 10]);

  it('find what a search one by one finds: numbers to 15 digits, text in any case, the first or last of equals', () => {
    check(rows, [
      ['MATCH(25,A1:A100,0)', 25],
      ['MATCH(0.1+0.2,A1:A100,0)', 50],
      ['MATCH("ITEM 60",A1:A100,0)', 61],
      ['VLOOKUP("item 60",A1:B100,2,FALSE)', 610],
      ['XLOOKUP("Item 50",A1:A100,B1:B100)', 510],
      ['XLOOKUP("Item 50",A1:A100,B1:B100,,0,-1)', 1000],
      ['MATCH(TRUE,A1:A100,0)', 91],
      ['MATCH(FALSE,A1:A100,0)', 92],
      ['MATCH("nothing",A1:A100,0)', NA],
      // An empty cell looked for isn't found, and an error looked for is the answer.
      ['MATCH(A96,A1:A100,0)', NA],
      ['MATCH(1/0,A1:A100,0)', DIV0],
      // Wildcards are matched one by one.
      ['MATCH("Item 6?",A1:A100,0)', 61],
    ]);
  });

  it('see a change to the list they search', () => {
    const e = nativeEngine.create(
      [{ name: 'S', rows: rows.map(([k], i) => [k, null, engineFormula(`A${i + 1}`), i ? null : engineFormula('MATCH("Item 60",C1:C100,0)')]) }],
      [],
      false,
    );
    expect(e.value(0, 0, 3)).toBe(61);
    e.write([{ sheet: 0, row: 60, col: 0, raw: 'Other' }]);
    expect(e.value(0, 0, 3)).toEqual(NA);
    e.write([{ sheet: 0, row: 10, col: 0, raw: 'Item 60' }]);
    expect(e.value(0, 0, 3)).toBe(11);
    e.destroy();
  });
});

describe('statistics', () => {
  it('give Microsoft’s examples', () => {
    check([], [
      ['STDEVP(2,4,4,4,5,5,7,9)', 2],
      ['VAR(1,2,3,4)', 5 / 3],
      ['VARP(1,2,3,4)', 1.25],
      ['NORMDIST(42,40,1.5,TRUE)', 0.9087887802741321],
      ['NORMDIST(42,40,1.5,FALSE)', 0.10934004978399577],
      ['NORMSDIST(1.333333)', 0.908788726],
      ['NORMSINV(0.9087888)', 1.33333345],
      ['NORMINV(0.908789,40,1.5)', 42.000002],
      ['PERCENTILE({1,2,3,4},0.3)', 1.9],
      ['QUARTILE({1,2,4,7,8,9,10,12},1)', 3.5],
      ['COVAR({3,2,4,5,6},{9,7,12,15,17})', 5.2],
      ['AGGREGATE(3,4,{1,"a",TRUE})', 3],
      ['COUNTA({1,"",TRUE})', 3],
      ['F.DIST(0,2,3,TRUE)', 0],
      ['F.DIST(0,2,3,FALSE)', 1],
    ]);
  });

  it('refuse what they can’t work out', () => {
    check([], [
      ['NORMDIST(42,40,0,TRUE)', NUM],
      ['NORMINV(0.5,40,0)', NUM],
      ['GEOMEAN(4,-1)', NUM],
      ['HARMEAN(4,0)', NUM],
      ['HYPGEOM.DIST(1,4,8,0,FALSE)', NUM],
      ['HYPGEOM.DIST(0,4,8,10,FALSE)', NUM],
      ['CHISQ.TEST({1,2},{0,1})', DIV0],
      ['CORREL({1,1,1},{1,2,3})', DIV0],
      ['RSQ({1,1,1},{1,2,3})', DIV0],
      ['T.TEST({1},{2,3},2,2)', DIV0],
      ['Z.TEST({1,1,1},1)', DIV0],
      ['STEYX({1,2},{3,4})', DIV0],
    ]);
  });
});

describe('criteria', () => {
  it('read arrays and empty criteria (as 0), and need ranges of one size', () => {
    check([[1, null], [0], [null]], [
      ['COUNTIFS({1,2,3},">1")', 2],
      ['COUNTIF(A1:A3,B1)', 1],
      ['COUNTIFS(A1:A3,">0",A1:A2,">0")', VALUE],
    ]);
  });
});

describe('the rest', () => {
  it('give what Excel gives', () => {
    check([[5]], [
      ['CODE("A")', 65],
      ['CODE("€")', 128],
      ['CODE("")', VALUE],
      ['UNICHAR(66)', 'B'],
      ['RANDBETWEEN(3,3)', 3],
      ['RANDBETWEEN(5,1)', NUM],
      ['SHEET()', 1],
      ['SHEET(T!A1)', 2],
      ['SHEET("T")', 2],
      ['SHEET("Nope")', NA],
      ['SHEETS()', 2],
      ['SHEETS(A1)', 1],
      ['AREAS((A1:A2,B1))', 2],
      ['AREAS(A1)', 1],
      ['ISFORMULA(1)', VALUE],
      ['ADDRESS(1,1,1,1,"")', '!$A$1'],
      // Functions of references refuse plain values.
      ['AREAS(1)', VALUE],
      ['SHEETS(1)', VALUE],
      ['COUNTBLANK(1)', VALUE],
      // MAP gives one value for each: a cell is its value, but a range, a union or a LAMBDA can't be one.
      ['MAP({1},LAMBDA(x,A1))', 5],
      ['MAP({1},LAMBDA(x,A1:A2))', { error: '#CALC!' }],
      ['MAP({1},LAMBDA(x,(A1,B1)))', VALUE],
      ['MAP({1},LAMBDA(x,LAMBDA(y,y)))', { error: '#CALC!' }],
      ['FACTDOUBLE(301)', NUM],
      ['MULTINOMIAL(100,80)', NUM],
      ['BASE(10,1)', NUM],
      ['BASE(10,37)', NUM],
    ]);
    const e = nativeEngine.create([{ name: 'S', rows: [] }], [], false);
    const r = e.calculate('=RAND()', 0) as number;
    expect(r >= 0 && r < 1).toBe(true);
    e.destroy();
  });
});

describe('arguments Excel refuses', () => {
  it('in lookups, references and paired lists', () => {
    check([[1, 2], [3, 4], [null], [null]], [
      ['VLOOKUP(1,A1:B2,0,FALSE)', VALUE],
      ['MATCH(1,A1:B2,0)', NA],
      ['XLOOKUP(1,A1:B2,A1:B2)', VALUE],
      ['OFFSET(1,1,1)', VALUE],
      ['TEXTSPLIT("a,b",",","")', [['a', 'b']]],
      ['GCD((A1:A2,B1))', 1],
      ['CORREL({1},{2})', DIV0],
      ['T.TEST({1},{2},2,1)', DIV0],
      ['T.TEST({1,2},{2,3},2,1)', DIV0],
      ['CHISQ.TEST({1},{1})', NA],
      ['Z.TEST(A3:A4,1)', NA],
      ['SUMX2MY2({1,2},{1})', NA],
      ['SUMX2MY2({1,2},{1,#N/A})', NA],
      ['AVEDEV(A3:A4)', NUM],
      ['DEVSQ(A3:A4)', NUM],
      // A database function needs its field.
      ['DSUM(A1:B2,,A1:A2)', VALUE],
    ]);
  });

  it('in maths, text, dates, complex numbers and bases', () => {
    check([], [
      ['SIN(2^27)', NUM],
      ['MID("abc",0,1)', VALUE],
      ['REPLACE("abc",0,1,"x")', VALUE],
      ['NPER(0.1,0,100)', NUM],
      ['WORKDAY.INTL(1,1,"1111111")', VALUE],
      ['DECIMAL("FF",1)', NUM],
      ['COMPLEX(TRUE,1)', VALUE],
      ['IMABS(TRUE)', VALUE],
      ['IMPOWER("2+3i",TRUE)', VALUE],
      ['IMPOWER("0",0)', NUM],
      ['BIN2DEC(TRUE)', VALUE],
      ['AGGREGATE(20,4,1)', VALUE],
      ['FILTER({1,2,3},{1,0})', VALUE],
      ['RRI(2,-1,4)', NUM],
      ['FVSCHEDULE(1,{0.1,#N/A})', NA],
      ['TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),TRUE)', VALUE],
    ]);
  });
});

describe('the engine', () => {
  it('works out a chain of formulas deeper than it reads at once', () => {
    const rows: RawContent[][] = [[1], ...Array.from({ length: 999 }, (_, i) => [engineFormula(`A${i + 1}+1`)])];
    const e = nativeEngine.create([{ name: 'S', rows }], [], false);
    expect(e.value(0, 999, 0)).toBe(1000);
    e.write([{ sheet: 0, row: 0, col: 0, raw: 2 }]);
    expect(e.value(0, 999, 0)).toBe(1001);
    e.destroy();
  });

  it('keeps the ranges a change reads within bounds, letting the oldest go', () => {
    // Forty formulas, each summing its own twenty cells: more ranges than are kept at once.
    const rows: RawContent[][] = Array.from({ length: 60 }, (_, i) => [i + 1, i < 40 ? engineFormula(`SUM(A${i + 1}:A${i + 20})`) : null]);
    const e = nativeEngine.create([{ name: 'S', rows }], [], false);
    expect(e.value(0, 39, 1)).toBe((40 + 59) * 10);
    e.write([{ sheet: 0, row: 39, col: 0, raw: 0 }]);
    expect([e.value(0, 0, 1), e.value(0, 39, 1)]).toEqual([210, (40 + 59) * 10 - 40]);
    e.destroy();
  });
});
