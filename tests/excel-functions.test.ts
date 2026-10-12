import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import type { FormulaEngine } from '../src/core/engine';
import type { EngineValue } from '../src/core/engineValue';
import { toFileFormula } from '../src/core/formula';
import { NativeEngine } from '../src/engine/workbook';

/**
 * The functions added to the engine, case by case: Excel's documented
 * examples where it gives them, and the edges of each function's rules.
 * S!A1:A5 = 4.14, 4.19, 5.17, 5.77, 6.39; B1:B5 = red, orange, yellow,
 * green, blue; C1 = "Little Red Riding Hood's red hood", C2 = "a,b;c,d";
 * D1:D5 = 1, "two", TRUE, empty, #DIV/0!; E1:E2 = dates 1 Jan and 1 Feb 2026.
 */
const ROWS = [
  [4.14, 'red', "Little Red Riding Hood's red hood", 1, '=DATE(2026,1,1)'],
  [4.19, 'orange', 'a,b;c,d', 'two', '=DATE(2026,2,1)'],
  [5.17, 'yellow', null, true],
  [5.77, 'green', 'x', null],
  [6.39, 'blue', null, '#DIV/0!'],
];

type Expected = number | string | boolean | { error: string } | Expected[][];

/** Microsoft's example tables: quiz grades, houses for sale, office buildings, and data with errors. */
const GRADES = [
  ['Emilio', 75, 85, 87],
  ['Julie', 94, 80, 88],
  ['Hans', 86, 93, 'Incomplete'],
  ['Frederique', 'Incomplete', 75, 75],
];
const HOUSES = [
  ['Cozy Rambler', 230000, 'Issaquah', 3, 'No'],
  ['Snug Bungalow', 197000, 'Bellevue', 2, 'Yes'],
  ['Cool Cape Codder', 345678, 'Bellevue', 4, 'Yes'],
  ['Splendid Split Level', 321900, 'Issaquah', 2, 'Yes'],
  ['Exclusive Tudor', 450000, 'Bellevue', 5, 'Yes'],
  ['Classy Colonial', 395000, 'Bellevue', 4, 'No'],
];
const OFFICES = [
  [2310, 2, 2, 20, 142000],
  [2333, 2, 2, 12, 144000],
  [2356, 3, 1.5, 33, 151000],
  [2379, 3, 2, 43, 150000],
  [2402, 2, 3, 53, 139000],
  [2425, 4, 2, 23, 169000],
  [2448, 2, 1.5, 99, 126000],
  [2471, 2, 2, 34, 142900],
  [2494, 3, 3, 23, 163000],
  [2517, 4, 4, 55, 169000],
  [2540, 2, 3, 22, 149000],
];
const AGGREGATE_DATA = [['#DIV/0!', 82], [72, 65], [30, 95], ['#NUM!', 63], [31, 53], [96, 71], [32, 55], [81, 83], [33, 100], [53, 91], [34, 89]];
/** Subtotals inside a range, which AGGREGATE can leave out. */
const SUBTOTALS = [[1], [2], ['=SUBTOTAL(9,A1:A2)'], [4], ['=1+AGGREGATE(9,4,A1:A2)'], ['=SUM({1,2})']];

let engine: FormulaEngine;
beforeAll(() => {
  engine = new NativeEngine(
    [
      { name: 'S', rows: ROWS },
      { name: 'Grades', rows: GRADES },
      { name: 'Houses', rows: HOUSES },
      { name: 'Offices', rows: OFFICES },
      { name: 'Agg', rows: AGGREGATE_DATA },
      { name: 'Sub', rows: SUBTOTALS },
    ],
    [],
    false,
  );
});
afterAll(() => engine.destroy());

const calc = (formula: string) => engine.calculate('=' + formula, 0);

function same(got: EngineValue | EngineValue[][], want: Expected): boolean {
  if (Array.isArray(want)) {
    return Array.isArray(got) && got.length === want.length && want.every((row, r) => row.length === got[r].length && row.every((w, c) => same(got[r][c], w)));
  }
  if (typeof want === 'number' && typeof got === 'number') return Math.abs(got - want) <= 1e-9 * Math.max(1, Math.abs(want));
  if (typeof want === 'object') return !!got && typeof got === 'object' && !Array.isArray(got) && got.error === want.error;
  return got === want;
}

/** Each formula with what it should give; lists every one that gives something else. */
function check(cases: [string, Expected][]) {
  const wrong = cases.flatMap(([f, want]) => {
    const got = calc(f);
    return same(got, want) ? [] : [`${f} → ${JSON.stringify(got)}, expected ${JSON.stringify(want)}`];
  });
  expect(wrong).toEqual([]);
}

const NA = { error: '#N/A' };
const VALUE = { error: '#VALUE!' };
const DIV0 = { error: '#DIV/0!' };
const NUM = { error: '#NUM!' };
const CALC = { error: '#CALC!' };

/**
 * What a formula entered in A1 shows in A1:D4: its whole result, as the
 * editor spills it (the session spills the engine's results, sized by it).
 */
function spill(formula: string, rows: (string | number | boolean | null)[][] = []) {
  const e = new NativeEngine([{ name: 'T', rows: [[null], ...rows] }], [], false);
  try {
    const got = e.calculate(formula, 0);
    const grid = Array.isArray(got) ? got : [[got]];
    return { values: [0, 1, 2, 3].map((r) => [0, 1, 2, 3].map((c) => grid[r]?.[c] ?? null)), isArray: Array.isArray(got) };
  } finally {
    e.destroy();
  }
}

describe('lookup functions', () => {
  it('LOOKUP finds the last value not above the key, in vector and array forms', () => {
    check([
      ['LOOKUP(4.19,A1:A5,B1:B5)', 'orange'],
      ['LOOKUP(5.75,A1:A5,B1:B5)', 'yellow'],
      ['LOOKUP(7.66,A1:A5,B1:B5)', 'blue'],
      ['LOOKUP(0,A1:A5,B1:B5)', NA],
      ['LOOKUP("C",{"a","b","c","d";1,2,3,4})', 3],
      ['LOOKUP("bump",{"a",1;"b",2;"c",3})', 2],
      ['LOOKUP(3,{1,"a";3,"c"})', 'c'],
      ['LOOKUP("b",{"a";"b";"c"})', 'b'],
      // The last match, and the last number or text, by searching past everything.
      ['LOOKUP(2,1/(B1:B5<>"blue"),A1:A5)', 5.77],
      ['LOOKUP(9.99*10^307,D1:D5)', 1],
      ['LOOKUP("zzz",B1:B5)', 'blue'],
      ['LOOKUP(6,A1:A5,B1:B2)', NA],
      ['LOOKUP(1,{1,2;3,4},{5,6;7,8})', NA],
      ['LOOKUP(NA(),A1:A5)', NA],
      ['LOOKUP(D4,{0,1},{"zero","one"})', NA],
      ['LOOKUP(1,{1},D4)', 0],
      ['LOOKUP({4.2,6},A1:A5,B1:B5)', [['orange', 'green']]],
    ]);
  });

  it('XMATCH matches exactly, by next smaller or larger, by wildcard, and by binary search', () => {
    check([
      ['XMATCH(4,{5,4,3,2,1})', 2],
      ['XMATCH(4.5,{5,4,3,2,1},1)', 1],
      ['XMATCH(4.5,{5,4,3,2,1},-1)', 2],
      ['XMATCH("ORANGE",B1:B5)', 2],
      ['XMATCH(5,A1:A5)', NA],
      ['XMATCH(5,A1:A5,1)', 3],
      ['XMATCH(5,A1:A5,-1)', 2],
      ['XMATCH(2.5,{3,4,3},1)', 1],
      ['XMATCH(2.5,{3,4,3},1,-1)', 3],
      ['XMATCH(1,{1,2,1})', 1],
      ['XMATCH(1,{1,2,1},0,-1)', 3],
      ['XMATCH("gr?en",B1:B5,2)', 4],
      ['XMATCH("*e",B1:B5,2)', 2],
      ['XMATCH("*e",B1:B5,2,-1)', 5],
      ['XMATCH(4.19,A1:A5,2)', 2],
      ['XMATCH(5.77,A1:A5,0,2)', 4],
      ['XMATCH(5.5,A1:A5,0,2)', NA],
      ['XMATCH(5.5,A1:A5,-1,2)', 3],
      ['XMATCH(5.5,A1:A5,1,2)', 4],
      ['XMATCH(1,A1:A5,-1,2)', NA],
      ['XMATCH(9,A1:A5,1,2)', NA],
      ['XMATCH(3,{5,4,3,2,1},0,-2)', 3],
      ['XMATCH(4.5,{5,4,3,2,1},-1,-2)', 2],
      ['XMATCH(4.5,{5,4,3,2,1},1,-2)', 1],
      ['XMATCH(1,{"1",1})', 2],
      ['XMATCH(TRUE(),{1,TRUE()})', 2],
      ['XMATCH(TRUE(),{FALSE()},-1)', 1],
      ['XMATCH(D4,{1,0})', NA],
      ['XMATCH(DATE(2026,2,1),E1:E2)', 2],
      // Wildcards can't be searched for by halves, whatever is looked for.
      ['XMATCH(5.77,A1:A5,2,2)', VALUE],
      ['XMATCH("z*",B1:B5,2)', NA],
      ['XMATCH({"red","blue"},B1:B5)', [[1, 5]]],
      ['XMATCH(1,{1,2},3)', VALUE],
      ['XMATCH(1,{1,2},0,3)', VALUE],
      ['XMATCH(1,{1,2;3,4})', VALUE],
      ['XMATCH(NA(),A1:A5)', NA],
    ]);
  });
});

describe('information functions', () => {
  it('TYPE names what kind of value it is given, errors and arrays included', () => {
    check([
      ['TYPE(1)', 1],
      ['TYPE(D4)', 1],
      ['TYPE("Smith")', 2],
      ['TYPE("MR. "&B1)', 2],
      ['TYPE(TRUE())', 4],
      ['TYPE(2+"Smith")', 16],
      ['TYPE({1,2;3,4})', 64],
      ['TYPE(B1:B5)', 64],
      ['TYPE()', VALUE],
    ]);
  });

  it('ERROR.TYPE numbers each error, and is #N/A for anything else', () => {
    check([
      ['ERROR.TYPE(1/0)', 2],
      ['ERROR.TYPE(D5)', 2],
      ['ERROR.TYPE("a"+1)', 3],
      ['ERROR.TYPE(#REF!)', 4],
      ['ERROR.TYPE(NOSUCHFUNCTION())', 5],
      ['ERROR.TYPE(SQRT(-1))', 6],
      ['ERROR.TYPE(NA())', 7],
      ['ERROR.TYPE(1)', NA],
    ]);
    // A cell reading itself reads 0, as Excel shows it, which has no error number.
    const e = new NativeEngine([{ name: 'T', rows: [['=A1', '=ERROR.TYPE(A1)']] }], [], false);
    expect(e.value(0, 0, 1)).toEqual({ error: '#N/A' });
    e.destroy();
  });
});

describe('text functions', () => {
  it('CONCAT joins values and ranges, writing numbers as Excel does', () => {
    check([
      ['CONCAT("The"," ","sun"," ","will"," ","come"," ","up"," ","tomorrow.")', 'The sun will come up tomorrow.'],
      ['CONCAT(B1:B3)', 'redorangeyellow'],
      ['CONCAT(A1:A2,"-",TRUE(),0.1+0.2)', '4.144.19-TRUE0.3'],
      ['CONCAT(D1:D4)', '1twoTRUE'],
      ['CONCAT(1/3)', '0.333333333333333'],
      ['CONCAT(0,"|",E1,"|",FALSE())', '0|46023|FALSE'],
      ['CONCAT(123456789012345678)', '1.23456789012346E+17'],
      ['CONCAT(10^20,"|",0.0001,"|",0.00001,"|",-1.5/10^7)', '1E+20|0.0001|1E-05|-1.5E-07'],
      ['CONCAT(D5)', DIV0],
      ['CONCAT(REPT("a",20000),REPT("b",20000))', VALUE],
    ]);
  });

  it('NUMBERVALUE reads numbers written with any separators', () => {
    check([
      ['NUMBERVALUE("2.500,27",",",".")', 2500.27],
      ['NUMBERVALUE("3.5%")', 0.035],
      ['NUMBERVALUE("9%%")', 0.0009],
      ['NUMBERVALUE(" 1 000 ")', 1000],
      ['NUMBERVALUE("1,234.5")', 1234.5],
      ['NUMBERVALUE("3,14",",")', 3.14],
      ['NUMBERVALUE("1.234,5",",",".")', 1234.5],
      ['NUMBERVALUE("1,5",,".")', VALUE],
      ['NUMBERVALUE("1.5e3")', 1500],
      ['NUMBERVALUE("-.5")', -0.5],
      ['NUMBERVALUE("")', 0],
      ['NUMBERVALUE(D4)', 0],
      ['NUMBERVALUE(5)', 5],
      ['NUMBERVALUE("1.234,5")', VALUE],
      ['NUMBERVALUE("1.2.3")', VALUE],
      ['NUMBERVALUE("abc")', VALUE],
      ['NUMBERVALUE("%")', VALUE],
      ['NUMBERVALUE(TRUE())', VALUE],
      ['NUMBERVALUE("1","")', VALUE],
      ['NUMBERVALUE("1",".",".")', VALUE],
      ['NUMBERVALUE(NA())', NA],
    ]);
  });

  it('TEXTSPLIT splits into columns and rows, padding short rows', () => {
    check([
      ['TEXTSPLIT("Dakota Lennon Sanchez"," ")', [['Dakota', 'Lennon', 'Sanchez']]],
      ['TEXTSPLIT("1,2,3;4,5,6",",",";")', [['1', '2', '3'], ['4', '5', '6']]],
      ['TEXTSPLIT(C2,",",";")', [['a', 'b'], ['c', 'd']]],
      ['TEXTSPLIT("a,b;c",",",";")', [['a', 'b'], ['c', NA]]],
      ['TEXTSPLIT("a,b;c",",",";",,,"-")', [['a', 'b'], ['c', '-']]],
      ['TEXTSPLIT("a,,b",",")', [['a', '', 'b']]],
      ['TEXTSPLIT("a,,b",",",,TRUE())', [['a', 'b']]],
      ['TEXTSPLIT("a;;b",,";")', [['a'], [''], ['b']]],
      ['TEXTSPLIT("a;;b",,";",TRUE())', [['a'], ['b']]],
      ['TEXTSPLIT("aXbxc","x")', [['aXb', 'c']]],
      ['TEXTSPLIT("aXbxc","x",,,1)', [['a', 'b', 'c']]],
      ['TEXTSPLIT("a-b_c",{"-","_"})', [['a', 'b', 'c']]],
      ['TEXTSPLIT("a--b",{"-","--"})', [['a', 'b']]],
      ['TEXTSPLIT(12.5,".")', [['12', '5']]],
      ['TEXTSPLIT("abc",",")', 'abc'],
      ['TEXTSPLIT(B1:B2,"e")', [['r'], ['orang']]],
      ['TEXTSPLIT({"a,b",",x"},",",,TRUE())', [['a', 'x']]],
      ['TEXTSPLIT({"a",""},",",,TRUE())', [['a', CALC]]],
      ['TEXTSPLIT(A1:A2/0,".")', [[DIV0], [DIV0]]],
      ['TEXTSPLIT(",",",",,TRUE())', CALC],
      ['TEXTSPLIT("abc","")', VALUE],
      ['TEXTSPLIT("abc",",",,,2)', VALUE],
      ['TEXTSPLIT(NA(),",")', NA],
      ['TEXTSPLIT("a",{",",#N/A})', NA],
      ['TEXTSPLIT("a",",",{";",#N/A})', NA],
    ]);
  });

  it('TEXTSPLIT splits text worked out by formulas as it does text written in', () => {
    expect(spill('=TEXTSPLIT("a,b;c",",",";")')).toEqual({ values: [['a', 'b', null, null], ['c', expect.objectContaining({ error: '#N/A' }), null, null], [null, null, null, null], [null, null, null, null]], isArray: true });
    expect(spill('=TEXTSPLIT({"a-b";"c-d";"e"},"-")').values.map((r) => r[0])).toEqual(['a', 'c', 'e', null]);
    expect(spill('=TEXTSPLIT("x;y",,";",FALSE(),0)').values.map((r) => r[0])).toEqual(['x', 'y', null, null]);
    expect(spill('=TEXTSPLIT("abc",",")')).toMatchObject({ values: [['abc', null, null, null], [null, null, null, null], [null, null, null, null], [null, null, null, null]], isArray: false });
    // Literal numbers, logicals and parenthesised arguments count as literals.
    expect(spill('=TEXTSPLIT(-12.5,".")').values[0]).toEqual(['-12', '5', null, null]);
    expect(spill('=TEXTSPLIT(TRUE(),"U")').values[0]).toEqual(['TR', 'E', null, null]);
    expect(spill('=TEXTSPLIT(FALSE(),"L")').values[0]).toEqual(['FA', 'SE', null, null]);
    expect(spill('=TEXTSPLIT(("a1b"),1)').values[0]).toEqual(['a', 'b', null, null]);
    expect(spill('=TEXTSPLIT("a,,b",",",,"TRUE")').values[0]).toEqual(['a', 'b', null, null]);
    expect(spill('=TEXTSPLIT("aXb","x",,,"1")').values[0]).toEqual(['a', 'b', null, null]);
    expect(spill('=TEXTSPLIT("a,,b",",",,"yes")').values[0]).toEqual([expect.objectContaining({ error: '#VALUE!' }), null, null, null]);
    expect(spill('=TEXTSPLIT("a,,b",",",,,2)').values[0]).toEqual([expect.objectContaining({ error: '#VALUE!' }), null, null, null]);
    expect(spill('=TEXTSPLIT("a,,b",{",",";"},,(+1))').values[0]).toEqual(['a', 'b', null, null]);
    // Text known only once worked out splits as written text does.
    expect(spill('=TEXTSPLIT(("a")&("-b"),"-")').values[0]).toEqual(['a', 'b', null, null]);
    expect(spill('=TEXTSPLIT(A2,",")', [['solo']]).values[0][0]).toBe('solo');
    expect(spill('=TEXTSPLIT(A2,",")', [['a,b']]).values[0]).toEqual(['a', 'b', null, null]);
    expect(spill('=TEXTSPLIT("a,b",B2)', [[null, ',']]).values[0]).toEqual(['a', 'b', null, null]);
    expect(spill('=TEXTSPLIT(UPPER("a,b"),",")').values[0]).toEqual(['A', 'B', null, null]);
    expect(spill('=TEXTSPLIT(-"a",",")').values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=TEXTSPLIT("a,b",",",,ABS(0))').values[0]).toEqual(['a', 'b', null, null]);
    expect(spill('=TEXTSPLIT(IF(TRUE(),"a,b",),",")').values[0]).toEqual(['a', 'b', null, null]);
    // Several texts each give their first part.
    expect(spill('=TEXTSPLIT(SEQUENCE(B2),",")', [[null, 2]]).values.map((r) => r[0])).toEqual(['1', '2', null, null]);
    expect(spill('=TEXTSPLIT("a,b","")').values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=TEXTSPLIT(",",",",,TRUE())').values[0][0]).toMatchObject({ error: '#CALC!' });
    expect(spill('=TEXTSPLIT()').values[0][0]).toMatchObject({ error: '#VALUE!' });
  });

  it('TEXTBEFORE and TEXTAFTER cut at the nth delimiter, from either end', () => {
    check([
      ['TEXTBEFORE(C1,"Red")', 'Little '],
      ['TEXTBEFORE(C1,"basket")', NA],
      ['TEXTBEFORE(C1,"red",2,1)', "Little Red Riding Hood's "],
      ['TEXTBEFORE(C1,"red",-2,1)', 'Little '],
      ['TEXTBEFORE(C1,"red",-1)', "Little Red Riding Hood's "],
      ['TEXTBEFORE(C1,"Red",,FALSE())', 'Little '],
      ['TEXTAFTER(C1,"Red")', " Riding Hood's red hood"],
      ['TEXTAFTER(C1,"red",-1)', ' hood'],
      ['TEXTAFTER(C1,"hood",2,1)', ''],
      ['TEXTAFTER(C1,"basket",1,,,"none")', 'none'],
      ['TEXTBEFORE("abc","")', ''],
      ['TEXTBEFORE("abc","",-1)', 'abc'],
      ['TEXTAFTER("abc","")', 'abc'],
      ['TEXTAFTER("abc","",-1)', ''],
      ['TEXTBEFORE("a-b","-",2,,1)', 'a-b'],
      ['TEXTAFTER("a-b","-",2,,1)', ''],
      ['TEXTBEFORE("a-b","-",-2,,1)', ''],
      ['TEXTAFTER("a-b","-",-2,,1)', 'a-b'],
      ['TEXTBEFORE("a-b","-",3,,1)', NA],
      ['TEXTBEFORE("a-b","-",-3,,1)', NA],
      ['TEXTBEFORE("a_b-c",{"-","_"})', 'a'],
      ['TEXTAFTER("a_b-c",{"-","_"},-1)', 'c'],
      ['TEXTBEFORE("aaa","aa",2)', NA],
      ['TEXTBEFORE(12.5,".")', '12'],
      ['TEXTBEFORE(B1:B2,"e")', [['r'], ['orang']]],
      ['TEXTBEFORE("abc","b",0)', VALUE],
      ['TEXTBEFORE("abc","b",4)', VALUE],
      ['TEXTBEFORE("abc","b",1,2)', VALUE],
      ['TEXTBEFORE("abc","b",1,0,2)', VALUE],
      ['TEXTBEFORE(NA(),"a")', NA],
      ['TEXTBEFORE("a",{"b",#N/A})', NA],
    ]);
  });

  it('VALUETOTEXT and ARRAYTOTEXT write values concisely or as a formula would', () => {
    check([
      ['VALUETOTEXT(TRUE())', 'TRUE'],
      ['VALUETOTEXT(1234.01234)', '1234.01234'],
      ['VALUETOTEXT("Hello")', 'Hello'],
      ['VALUETOTEXT("Hello",1)', '"Hello"'],
      ['VALUETOTEXT("say "&CHAR(34)&"hi"&CHAR(34),1)', '"say ""hi"""'],
      ['VALUETOTEXT(1/0)', '#DIV/0!'],
      ['VALUETOTEXT(D4)', ''],
      ['VALUETOTEXT(A1:A2)', [['4.14'], ['4.19']]],
      ['VALUETOTEXT(1,2)', VALUE],
      ['ARRAYTOTEXT({1,"a";TRUE(),2})', '1, a, TRUE, 2'],
      ['ARRAYTOTEXT({1,"a";TRUE(),2},1)', '{1,"a";TRUE,2}'],
      ['ARRAYTOTEXT(D1:D5)', '1, two, TRUE, , #DIV/0!'],
      ['ARRAYTOTEXT(D1:D5,1)', '{1;"two";TRUE;;#DIV/0!}'],
      ['ARRAYTOTEXT(5)', '5'],
      ['ARRAYTOTEXT(A1:A2,2)', VALUE],
    ]);
  });

  it('the byte-counting functions match their character-counting ones', () => {
    check([
      ['LENB("abc")', 3],
      ['LEFTB("abc",2)', 'ab'],
      ['RIGHTB("abc",1)', 'c'],
      ['MIDB("abcde",2,3)', 'bcd'],
      ['FINDB("c","abc")', 3],
      ['SEARCHB("C","abc")', 3],
      ['REPLACEB("abc",2,1,"X")', 'aXc'],
      ['LENB(B1:B2)', [[3], [6]]],
      ['LEFTB()', VALUE],
    ]);
    expect(spill('=LENB({"a","bb"})').values[0]).toEqual([1, 2, null, null]);
  });

  it('FIXED rounds and formats a number as text', () => {
    check([
      ['FIXED(1234.567,1)', '1,234.6'],
      ['FIXED(1234.567,-1)', '1,230'],
      ['FIXED(-1234.567,-1,TRUE())', '-1230'],
      ['FIXED(44.332)', '44.33'],
      ['FIXED(1234567.891)', '1,234,567.89'],
      ['FIXED(999.996,2)', '1,000.00'],
      ['FIXED(1.005,2)', '1.01'],
      ['FIXED(-0.5,0)', '-1'],
      ['FIXED(-0.001,2)', '0.00'],
      ['FIXED(0)', '0.00'],
      ['FIXED(0.001234,5)', '0.00123'],
      // Never more than Excel's 15 significant digits.
      ['FIXED(0.1+0.2,20)', '0.30000000000000000000'],
      ['FIXED(1234567890123456789,0)', '1,234,567,890,123,460,000'],
      ['FIXED(10^21,0,TRUE())', '1000000000000000000000'],
      ['LEN(FIXED(1,127))', 129],
      ['FIXED(1,128)', VALUE],
      ['FIXED("a")', VALUE],
    ]);
  });
});

describe('dynamic array functions', () => {
  const GRID = '{1,2,3;4,5,6;7,8,9}';

  it('TAKE and DROP keep or remove rows and columns from either end', () => {
    check([
      [`TAKE(${GRID},2)`, [[1, 2, 3], [4, 5, 6]]],
      [`TAKE(${GRID},,2)`, [[1, 2], [4, 5], [7, 8]]],
      [`TAKE(${GRID},-2)`, [[4, 5, 6], [7, 8, 9]]],
      [`TAKE(${GRID},2,2)`, [[1, 2], [4, 5]]],
      [`TAKE(${GRID},-2,-2)`, [[5, 6], [8, 9]]],
      [`TAKE(${GRID},10)`, [[1, 2, 3], [4, 5, 6], [7, 8, 9]]],
      [`TAKE(${GRID},1.9)`, [[1, 2, 3]]],
      ['TAKE(B1:B5,-1)', 'blue'],
      ['TAKE(5,1)', 5],
      [`TAKE(${GRID},0)`, VALUE],
      [`TAKE(${GRID},1,0)`, VALUE],
      ['TAKE(1/0,1)', DIV0],
      [`DROP(${GRID},2)`, [[7, 8, 9]]],
      [`DROP(${GRID},,2)`, [[3], [6], [9]]],
      [`DROP(${GRID},-2)`, [[1, 2, 3]]],
      [`DROP(${GRID},2,2)`, 9],
      [`DROP(${GRID},0)`, [[1, 2, 3], [4, 5, 6], [7, 8, 9]]],
      [`DROP(${GRID},3)`, VALUE],
      [`DROP(${GRID},,-5)`, VALUE],
    ]);
  });

  it('CHOOSEROWS and CHOOSECOLS pick by position, from either end, repeating as asked', () => {
    check([
      [`CHOOSEROWS(${GRID},1,3,1)`, [[1, 2, 3], [7, 8, 9], [1, 2, 3]]],
      [`CHOOSEROWS(${GRID},{1,3},-1)`, [[1, 2, 3], [7, 8, 9], [7, 8, 9]]],
      [`CHOOSEROWS(${GRID},-1,-2)`, [[7, 8, 9], [4, 5, 6]]],
      [`CHOOSEROWS(${GRID},"2")`, [[4, 5, 6]]],
      [`CHOOSECOLS(${GRID},1,3)`, [[1, 3], [4, 6], [7, 9]]],
      [`CHOOSECOLS(${GRID},3,-3)`, [[3, 1], [6, 4], [9, 7]]],
      [`CHOOSECOLS(${GRID},2,2)`, [[2, 2], [5, 5], [8, 8]]],
      [`CHOOSEROWS(${GRID},0)`, VALUE],
      [`CHOOSEROWS(${GRID},4)`, VALUE],
      [`CHOOSECOLS(${GRID},-4)`, VALUE],
      [`CHOOSEROWS(${GRID},"x")`, VALUE],
      [`CHOOSEROWS(${GRID},{1,#N/A})`, NA],
      [`CHOOSEROWS(${GRID},NA())`, NA],
    ]);
  });

  it('TOCOL and TOROW flatten by rows or columns, leaving out blanks or errors', () => {
    check([
      [`TOCOL(${GRID})`, [[1], [2], [3], [4], [5], [6], [7], [8], [9]]],
      [`TOROW(${GRID},,TRUE())`, [[1, 4, 7, 2, 5, 8, 3, 6, 9]]],
      ['TOROW(D1:D5,1)', [[1, 'two', true, DIV0]]],
      ['TOROW(D1:D5,3)', [[1, 'two', true]]],
      ['TOCOL(D4,1)', CALC],
      ['TOCOL(D1:D5,4)', VALUE],
      ['TOCOL(D1:D5,-1)', VALUE],
      ['TOROW(D1:D5,4)', VALUE],
    ]);
    expect(calc('TOROW(D1:D5,2)')).toEqual([[1, 'two', true, 0]]);
  });

  it('WRAPROWS and WRAPCOLS wrap a row or column, padding the last line', () => {
    check([
      ['WRAPROWS({"A","B","C","D","E","F","G"},3)', [['A', 'B', 'C'], ['D', 'E', 'F'], ['G', NA, NA]]],
      ['WRAPROWS({"A","B","C","D","E","F","G"},3,"x")', [['A', 'B', 'C'], ['D', 'E', 'F'], ['G', 'x', 'x']]],
      ['WRAPCOLS({"A","B","C","D","E","F","G"},3)', [['A', 'D', 'G'], ['B', 'E', NA], ['C', 'F', NA]]],
      ['WRAPROWS(A1:A5,2.5)', [[4.14, 4.19], [5.17, 5.77], [6.39, NA]]],
      ['WRAPROWS({1,2},5)', [[1, 2, NA, NA, NA]]],
      ['WRAPROWS({1,2;3,4},2)', VALUE],
      ['WRAPROWS({1,2},0)', NUM],
      ['WRAPCOLS({1,2},0)', NUM],
      ['WRAPROWS({1,2},2000000)', NUM],
    ]);
  });

  it('EXPAND pads an array out to a size', () => {
    check([
      ['EXPAND({1,2;3,4},3,3)', [[1, 2, NA], [3, 4, NA], [NA, NA, NA]]],
      ['EXPAND({1,2;3,4},3,,"-")', [[1, 2], [3, 4], ['-', '-']]],
      ['EXPAND({1,2;3,4},,3,0)', [[1, 2, 0], [3, 4, 0]]],
      ['EXPAND({1,2;3,4},1)', VALUE],
      ['EXPAND({1,2;3,4},2,1)', VALUE],
      ['EXPAND(1,2000,1000)', NUM],
    ]);
  });

  it('SORTBY sorts by other rows or columns, in either order, keeping ties in place', () => {
    const NAMES = '{"Tom";"Fred";"Amy";"Sal";"Fritz";"Xio"}';
    check([
      [`SORTBY(${NAMES},{52;65;22;73;19;39})`, [['Fritz'], ['Amy'], ['Xio'], ['Tom'], ['Fred'], ['Sal']]],
      [`SORTBY(${NAMES},{52;65;22;73;19;39},-1)`, [['Sal'], ['Fred'], ['Tom'], ['Xio'], ['Amy'], ['Fritz']]],
      [`SORTBY(${NAMES},{"b";"a";"b";"a";"b";"a"},1,{1;2;3;4;5;6},-1)`, [['Xio'], ['Sal'], ['Fred'], ['Fritz'], ['Amy'], ['Tom']]],
      [`SORTBY(${NAMES},{"b";"a";"b";"a";"b";"a"})`, [['Fred'], ['Sal'], ['Xio'], ['Tom'], ['Amy'], ['Fritz']]],
      ['SORTBY({1,2,3},{"c","b","a"})', [[3, 2, 1]]],
      // Numbers, then text, then logicals and errors, with blanks last either way.
      ['SORTBY({1;2;3;4;5},D1:D5)', [[1], [2], [3], [5], [4]]],
      ['SORTBY({1;2;3;4;5},D1:D5,-1)', [[5], [3], [2], [1], [4]]],
      ['SORTBY({1;2;3},{#N/A;1;#DIV/0!})', [[2], [1], [3]]],
      ['SORTBY({1,2},{1,2},2)', VALUE],
      ['SORTBY({1;2},{1,2;3,4})', VALUE],
      ['SORTBY({1;2},{1;2;3})', VALUE],
      ['SORTBY({1,2},{1,2})', [[1, 2]]],
      ['SORTBY({1,2;3,4},{2;1},1,{2,1},1)', VALUE],
    ]);
  });

  it('RANDARRAY fills an array with random numbers in range', () => {
    const decimals = calc('RANDARRAY(3,2,5,6)') as number[][];
    expect(decimals.length).toBe(3);
    expect(decimals.flat().every((n) => n >= 5 && n < 6 && decimals[0].length === 2)).toBe(true);
    const whole = (calc('RANDARRAY(4,4,1,3,TRUE())') as number[][]).flat();
    expect(whole.every((n) => Number.isInteger(n) && n >= 1 && n <= 3)).toBe(true);
    expect(typeof calc('RANDARRAY()')).toBe('number');
    check([
      ['RANDARRAY(0)', CALC],
      ['RANDARRAY(-1)', VALUE],
      ['RANDARRAY(1,1,2,1)', VALUE],
      ['RANDARRAY(1,1,1.5,3,TRUE())', VALUE],
      ['RANDARRAY(2000,1000)', NUM],
    ]);
  });

  it('spill to the size their literal arguments give, or as far as their input reaches', () => {
    const column = (s: { values: EngineValue[][] }, c = 0) => s.values.map((r) => r[c]);
    expect(spill('=TAKE({1,2,3;4,5,6;7,8,9},2,-1)').values.slice(0, 3)).toEqual([[3, null, null, null], [6, null, null, null], [null, null, null, null]]);
    expect(column(spill('=DROP({1;2;3;4},1)'))).toEqual([2, 3, 4, null]);
    expect(spill('=CHOOSECOLS({1,2,3},3,1,2)').values[0]).toEqual([3, 1, 2, null]);
    expect(column(spill('=CHOOSEROWS({1;2;3},{3,1})'))).toEqual([3, 1, null, null]);
    expect(spill('=TOROW({1;2;3})').values[0]).toEqual([1, 2, 3, null]);
    expect(column(spill('=TOCOL({1,2})'))).toEqual([1, 2, null, null]);
    expect(spill('=WRAPCOLS({1,2,3},2)').values.slice(0, 2)).toEqual([[1, 3, null, null], [2, expect.objectContaining({ error: '#N/A' }), null, null]]);
    expect(spill('=EXPAND(1,2,2,0)').values.slice(0, 2)).toEqual([[1, 0, null, null], [0, 0, null, null]]);
    expect(column(spill('=SORTBY({"b";"a"},{2;1})'))).toEqual(['a', 'b', null, null]);
    expect(spill('=RANDARRAY(2,3)').values.flat().filter((v) => typeof v === 'number')).toHaveLength(6);
    expect(spill('=RANDARRAY(,2)').values.flat().filter((v) => typeof v === 'number')).toHaveLength(2);
    expect(spill('=WRAPROWS({1,2,3},2)').values.slice(0, 2)).toEqual([[1, 2, null, null], [3, expect.objectContaining({ error: '#N/A' }), null, null]]);
    expect(spill('=EXPAND({1;2},,2,0)').values.slice(0, 2)).toEqual([[1, 0, null, null], [2, 0, null, null]]);
    expect(column(spill('=EXPAND({1;2},3,,0)'))).toEqual([1, 2, 0, null]);
    expect(column(spill('=TAKE({1;2;3},(2))'))).toEqual([1, 2, null, null]);
    expect(column(spill('=TAKE({1;2;3},-B2)', [[null, 1]]))).toEqual([3, null, null, null]);
    expect(spill('=TAKE({1,2;3,4},1,B2)', [[null, 1]]).values[0]).toEqual([1, null, null, null]);
    expect(spill('=TAKE({1;2;3},"a")').values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=RANDARRAY(2000,1000)').values[0][0]).toMatchObject({ error: '#NUM!' });
    // A count from a cell takes the whole input's room, the cells past the result left empty.
    expect(column(spill('=TAKE(C2:C4,D2)', [[null, null, 10, 2], [null, null, 20], [null, null, 30]]))).toEqual([10, 20, null, null]);
    expect(column(spill('=TOCOL(C2:C4,1)', [[null, null, 10], [null, null, null], [null, null, 30]]))).toEqual([10, 30, null, null]);
    // Sizes known only once worked out spill as any other.
    expect(column(spill('=TAKE(SEQUENCE(B2),2)', [[null, 3]]))).toEqual([1, 2, null, null]);
    expect(spill('=WRAPROWS({1,2,3},B2)', [[null, 2]]).values.slice(0, 2)).toEqual([[1, 2, null, null], [3, NA, null, null]]);
    expect(column(spill('=EXPAND(1,B2)', [[null, 2]]))).toEqual([1, NA, null, null]);
    expect(spill('=ROWS(RANDARRAY(B2))', [[null, 2]]).values[0][0]).toBe(2);
    expect(column(spill('=CHOOSEROWS(A2:A3,SEQUENCE(B2))', [[1, 2], [3]]))).toEqual([1, 3, null, null]);
    expect(spill('=EXPAND({1,2},1,1)').values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=WRAPROWS({1,2;3,4},2)').values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=DROP({1;2},2)').values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=TAKE({1;2},0)').values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=CHOOSEROWS({1;2})').values[0][0]).toMatchObject({ error: '#VALUE!' });
    for (const f of ['=TOCOL()', '=TOROW()', '=SORTBY()', '=WRAPROWS()', '=WRAPCOLS()', '=EXPAND()', '=TAKE()', '=DROP()']) {
      expect(spill(f).values[0][0]).toMatchObject({ error: expect.stringMatching(/^#/) });
    }
  });
});

describe('statistical functions', () => {
  it('AVERAGEIFS averages the values meeting every criterion', () => {
    check([
      ['AVERAGEIFS(Grades!B1:B4,Grades!B1:B4,">70",Grades!B1:B4,"<90")', 80.5],
      ['AVERAGEIFS(Grades!C1:C4,Grades!C1:C4,">95")', DIV0],
      ['AVERAGEIFS(Grades!D1:D4,Grades!D1:D4,"<>Incomplete",Grades!D1:D4,">80")', 87.5],
      ['AVERAGEIFS(Houses!B1:B6,Houses!C1:C6,"Bellevue",Houses!D1:D6,">2",Houses!E1:E6,"Yes")', 397839],
      ['AVERAGEIFS(Houses!B1:B6,Houses!C1:C6,"Issaquah",Houses!D1:D6,"<=3",Houses!E1:E6,"No")', 230000],
      ['AVERAGEIFS(Houses!B1:B6,Houses!A1:A6,"C*")', 323559.3333333333],
      ['AVERAGEIFS(D1:D5,A1:A5,">5")', DIV0],
      ['AVERAGEIFS(D1:D5,A1:A5,">6")', DIV0],
      ['AVERAGEIFS(D1:D5,A1:A5,">6.3")', DIV0],
      ['AVERAGEIFS(D1:D5,A1:A5,"<5")', 1],
      ['AVERAGEIFS(D1:D5,A1:A5,">6")', DIV0],
      ['AVERAGEIFS(D1:D5,A1:A5,">=6.39")', DIV0],
      ['AVERAGEIFS(A1:A3,D1:D3,"<>two")', 4.655],
      ['AVERAGEIFS(D1:D5,A1:A5,">6",A1:A5,"<7")', DIV0],
      ['AVERAGEIFS(D5,A5,">6")', DIV0],
      ['AVERAGEIFS(A1:A5,B1:B4,"red")', VALUE],
      ['AVERAGEIFS(A1:A5,A1:A5,D4)', DIV0],
    ]);
  });

  it('MODE and MODE.MULT find the most frequent numbers', () => {
    check([
      ['MODE({5.6,4,4,3,2,4})', 4],
      ['MODE.SNGL({5.6,4,4,3,2,4})', 4],
      ['MODE(1,2,2,1)', 1],
      ['MODE("1","1",2)', 1],
      ['MODE(0,,)', 0],
      ['MODE(1,2,3)', NA],
      ['MODE(D1:D4,1)', 1],
      ['MODE(D1:D5)', DIV0],
      ['MODE(1,"a")', VALUE],
      ['MODE.MULT({1,2,3,4,3,2,1,2,3,5,6,1})', [[1], [2], [3]]],
      ['MODE.MULT(1,2)', NA],
      ['MODE.MULT(D5)', DIV0],
    ]);
    expect(spill('=MODE.MULT({1,1,2})').values.map((r) => r[0])).toEqual([1, null, null, null]);
    expect(spill('=MODE.MULT({1,1,2,2})').values.map((r) => r[0])).toEqual([1, 2, null, null]);
  });

  it('FREQUENCY counts values into bins, whatever order the bins are in', () => {
    check([
      ['FREQUENCY({79,85,78,85,50,81,95,88,97},{70,79,89})', [[1], [2], [4], [2]]],
      ['FREQUENCY({1,2,3,4,5},{4,2})', [[2], [2], [1]]],
      ['FREQUENCY({1,2,3},{2,2})', [[2], [0], [1]]],
      ['FREQUENCY({1,2,"x"},{"a"})', 2],
      ['FREQUENCY({1,#N/A},{1})', NA],
      ['FREQUENCY({1},{1,#N/A})', NA],
    ]);
    expect(spill('=FREQUENCY({1,2,3},{1,2})').values.map((r) => r[0])).toEqual([1, 1, 1, null]);
    expect(spill('=FREQUENCY({1,2,3},SEQUENCE(B2))', [[null, 1]]).values.map((r) => r[0])).toEqual([1, 2, null, null]);
  });

  it('TRIMMEAN, KURT, PERMUT, PERMUTATIONA, GESTEP and RANK.AVG', () => {
    check([
      ['TRIMMEAN({4,5,6,7,2,3,4,5,1,2,3},0.2)', 3.7777777777777777],
      ['TRIMMEAN({1,2,3,100},0.5)', 2.5],
      ['TRIMMEAN({1,2,3,4},1)', NUM],
      ['TRIMMEAN({1,2},-0.1)', NUM],
      ['TRIMMEAN({1,2},1.1)', NUM],
      ['TRIMMEAN("a",0)', VALUE],
      ['TRIMMEAN(D1:D5,0)', DIV0],
      ['KURT(3,4,5,2,3,4,5,6,4,7)', -0.15179963720841627],
      ['KURT(1,2,3)', DIV0],
      ['KURT(1,1,1,1)', DIV0],
      ['KURT(D5,1,2,3)', DIV0],
      ['PERMUT(100,3)', 970200],
      ['PERMUT(3,2)', 6],
      ['PERMUT(5.9,2.1)', 20],
      ['PERMUT(2,3)', NUM],
      ['PERMUT(-1,0)', NUM],
      ['PERMUTATIONA(3,2)', 9],
      ['PERMUTATIONA(2,2)', 4],
      ['PERMUTATIONA(0,0)', 1],
      ['PERMUTATIONA(-1,2)', NUM],
      ['GESTEP(5,4)', 1],
      ['GESTEP(5,5)', 1],
      ['GESTEP(-4,-5)', 1],
      ['GESTEP(-1)', 0],
      ['GESTEP("a")', VALUE],
      ['RANK.AVG(94,{89;88;92;101;94;97;95})', 4],
      ['RANK.AVG(94,{89;88;92;101;94;97;95},1)', 4],
      ['RANK.AVG(2,{1,2,2,3})', 2.5],
      ['RANK.AVG(2,{1,2,2,3},1)', 2.5],
      ['RANK.AVG(5,{1,2})', NA],
      ['RANK.AVG(1,D1:D5)', DIV0],
    ]);
  });

  it('PERCENTRANK cuts its result down to the significant digits', () => {
    const DATA = '{13,12,11,8,4,3,2,1,1,1}';
    const EXC = '{1,2,3,6,6,6,7,8,9}';
    check([
      [`PERCENTRANK.INC(${DATA},2)`, 0.333],
      [`PERCENTRANK.INC(${DATA},4)`, 0.555],
      [`PERCENTRANK.INC(${DATA},8)`, 0.666],
      [`PERCENTRANK.INC(${DATA},5)`, 0.583],
      [`PERCENTRANK(${DATA},5,1)`, 0.5],
      [`PERCENTRANK.EXC(${EXC},7)`, 0.7],
      [`PERCENTRANK.EXC(${EXC},5.43)`, 0.381],
      [`PERCENTRANK.EXC(${EXC},5.43,1)`, 0.3],
      ['PERCENTRANK.INC({5},5)', 1],
      ['PERCENTRANK.EXC({5},5)', 0.5],
      [`PERCENTRANK.INC(${DATA},0)`, NA],
      [`PERCENTRANK.INC(${DATA},14)`, NA],
      [`PERCENTRANK.INC(${DATA},2,0)`, NUM],
      ['PERCENTRANK.INC({"a"},1)', NUM],
      ['PERCENTRANK.INC(D1:D5,1)', DIV0],
    ]);
  });

  it('PERCENTOF gives a subset sum as a share of the whole', () => {
    check([
      ['PERCENTOF({1,2,3},{1,2,3,4})', 0.6],
      ['PERCENTOF(A1:A2,A1:A5)', (4.14 + 4.19) / 25.66],
      ['PERCENTOF(10,100)', 0.1],
      ['PERCENTOF(D1:D4,{1,1})', 0.5],
      ['PERCENTOF(1,{0})', DIV0],
      ['PERCENTOF(D5,1)', DIV0],
      ['PERCENTOF(1,NA())', NA],
      ['PERCENTOF(1)', VALUE],
    ]);
    expect(toFileFormula('PERCENTOF(A1:A2,A1:A9)')).toBe('_xlfn.PERCENTOF(A1:A2,A1:A9)');
  });

  it('PROB, BINOM.DIST.RANGE, ERF.PRECISE and ERFC.PRECISE', () => {
    check([
      ['PROB({0,1,2,3},{0.2,0.3,0.1,0.4},2)', 0.1],
      ['PROB({0,1,2,3},{0.2,0.3,0.1,0.4},1,3)', 0.8],
      ['PROB({0,1,"x"},{0.5,0.5,"y"},0,1)', 1],
      ['PROB({0,1},{0.5,0.4},0)', NUM],
      ['PROB({0,1},{0,1},0)', NUM],
      ['PROB({0,1},{0.5},0)', NA],
      ['PROB({0,#N/A},{0.5,0.5},0)', NA],
      ['PROB({0,1},{0.5,#N/A},0)', NA],
      ['ROUND(BINOM.DIST.RANGE(60,0.75,48),3)', 0.084],
      ['ROUND(BINOM.DIST.RANGE(60,0.75,45,50),3)', 0.524],
      ['BINOM.DIST.RANGE(5,0,0)', 1],
      ['BINOM.DIST.RANGE(5,1,0,4)', 0],
      ['BINOM.DIST.RANGE(5,1,5)', 1],
      ['BINOM.DIST.RANGE(2000,0.5,0,2000)', 1],
      ['BINOM.DIST.RANGE(-1,0.5,0)', NUM],
      ['BINOM.DIST.RANGE(5,1.5,0)', NUM],
      ['BINOM.DIST.RANGE(5,0.5,6)', NUM],
      ['BINOM.DIST.RANGE(5,0.5,3,2)', NUM],
      ['BINOM.DIST.RANGE(20000000,0.5,1)', 0],
      ['ROUND(ERF.PRECISE(0.745),6)', 0.707929],
      ['ROUND(ERF.PRECISE(1),7)', 0.8427008],
      ['ROUND(ERFC.PRECISE(1),7)', 0.1572992],
      ['ERF.PRECISE(1,2)', VALUE],
    ]);
  });

  it('MDETERM, MINVERSE and MUNIT', () => {
    check([
      ['MDETERM({1,3,8,5;1,3,6,1;1,1,1,0;7,3,10,2})', 88],
      ['MDETERM({3,6,1;1,1,0;3,10,2})', 1],
      ['MDETERM({3,6;1,1})', -3],
      ['MDETERM({1,2;2,4})', 0],
      ['MDETERM({1,3,8,5;1,3,6,1})', VALUE],
      ['MDETERM({1,"a";1,2})', VALUE],
      ['MDETERM({1,#N/A;1,2})', NA],
      ['MDETERM(MUNIT(401))', NUM],
      ['MINVERSE({4,-1;2,0})', [[0, 0.5], [-1, 2]]],
      ['MINVERSE({1,2,1;3,4,-1;0,2,0})', [[0.25, 0.25, -0.75], [0, 0, 0.5], [0.75, -0.25, -0.25]]],
      ['MINVERSE({1,2;2,4})', NUM],
      ['MINVERSE({1,2})', VALUE],
      ['MUNIT(3)', [[1, 0, 0], [0, 1, 0], [0, 0, 1]]],
      ['MUNIT(0)', VALUE],
    ]);
    expect(spill('=MINVERSE({4,-1;2,0})').values.slice(0, 2)).toEqual([[0, 0.5, null, null], [-1, 2, null, null]]);
    expect(spill('=MUNIT(2)').values.slice(0, 2)).toEqual([[1, 0, null, null], [0, 1, null, null]]);
    expect(spill('=MINVERSE({1,2,3})').values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=MUNIT(B2)', [[null, 2]]).values.slice(0, 2)).toEqual([[1, 0, null, null], [0, 1, null, null]]);
    expect(spill('=MUNIT(2000)').values[0][0]).toMatchObject({ error: '#NUM!' });
    expect(spill('=MUNIT()').values[0][0]).toMatchObject({ error: '#VALUE!' });
  });

  it('AGGREGATE applies any of 19 functions, leaving out errors or subtotals as asked', () => {
    const DATA = '{1,2,3,4,5,#N/A}';
    check([
      ['AGGREGATE(4,6,Agg!A1:A11)', 96],
      ['AGGREGATE(14,6,Agg!A1:A11,3)', 72],
      ['AGGREGATE(15,6,Agg!A1:A11)', VALUE],
      ['AGGREGATE(12,6,Agg!A1:A11,Agg!B1:B11)', 68],
      ['AGGREGATE(9,4,Agg!A1:A11)', DIV0],
      ['AGGREGATE(9,4,Agg!B1:B11)', 847],
      [`AGGREGATE(1,6,${DATA})`, 3],
      [`AGGREGATE(2,4,${DATA})`, 5],
      [`AGGREGATE(3,4,${DATA})`, 6],
      [`AGGREGATE(3,6,${DATA},"a")`, 6],
      [`AGGREGATE(4,6,${DATA})`, 5],
      [`AGGREGATE(5,6,${DATA})`, 1],
      [`AGGREGATE(6,6,${DATA})`, 120],
      [`AGGREGATE(7,6,${DATA})`, 1.5811388300841898],
      [`AGGREGATE(8,6,${DATA})`, 1.4142135623730951],
      [`AGGREGATE(9,6,${DATA})`, 15],
      [`AGGREGATE(10,6,${DATA})`, 2.5],
      [`AGGREGATE(11,6,${DATA})`, 2],
      [`AGGREGATE(12,6,${DATA})`, 3],
      ['AGGREGATE(12,6,{1,4,2,3})', 2.5],
      [`AGGREGATE(13,6,${DATA})`, NA],
      ['AGGREGATE(13,6,{1,2,2,1,3})', 1],
      [`AGGREGATE(14,6,${DATA},2)`, 4],
      [`AGGREGATE(15,6,${DATA},1.2)`, 2],
      [`AGGREGATE(16,6,${DATA},0.3)`, 2.2],
      [`AGGREGATE(17,6,${DATA},1)`, 2],
      [`AGGREGATE(18,6,${DATA},0.3)`, 1.8],
      [`AGGREGATE(19,6,${DATA},1)`, 1.5],
      [`AGGREGATE(16,6,${DATA},1)`, 5],
      ['AGGREGATE(1,6,{"a"})', DIV0],
      ['AGGREGATE(4,6,{"a"})', 0],
      ['AGGREGATE(5,6,{"a"})', 0],
      ['AGGREGATE(6,6,{"a"})', 0],
      ['AGGREGATE(7,6,{1})', DIV0],
      ['AGGREGATE(8,6,{"a"})', DIV0],
      ['AGGREGATE(10,6,{1})', DIV0],
      ['AGGREGATE(11,6,{"a"})', DIV0],
      ['AGGREGATE(12,6,{"a"})', NUM],
      ['AGGREGATE(14,6,{1,2},3)', NUM],
      ['AGGREGATE(15,6,{1},0)', NUM],
      ['AGGREGATE(16,6,{1},2)', NUM],
      ['AGGREGATE(16,6,{"a"},0.5)', NUM],
      ['AGGREGATE(17,6,{1,2},5)', NUM],
      ['AGGREGATE(18,6,{1,2},0.1)', NUM],
      ['AGGREGATE(19,6,{1,2,3},4)', NUM],
      ['AGGREGATE(14,6,{1},NA())', NA],
      ['AGGREGATE(14,6,{1},"x")', VALUE],
      ['AGGREGATE(9,,{1,2})', 3],
      ['AGGREGATE(9,4,1,2)', 3],
      ['AGGREGATE(9,6,1,NA())', 1],
      ['AGGREGATE(9,4,1/0)', DIV0],
      ['AGGREGATE(NA(),6,{1})', NA],
      ['AGGREGATE(20,6,{1})', VALUE],
      ['AGGREGATE(9,8,{1})', VALUE],
      ['AGGREGATE(9,6)', VALUE],
      // Cells holding SUBTOTAL or AGGREGATE are left out with options 0 to 3.
      ['AGGREGATE(9,0,Sub!A1:A6)', 10],
      ['AGGREGATE(9,4,Sub!A1:A6)', 17],
      ['AGGREGATE(4,3,Sub!A1:A6)', 4],
      ['AGGREGATE(14,0,Sub!A1:A6,1)', 4],
      ['AGGREGATE(9,0,(Sub!A1:A6)*1)', 17],
    ]);
  });
});

describe('regression functions', () => {
  const COSTS = '{133890;135000;135790;137300;138130;139100;139900;141120;141890;143230;144000;145290}';
  const MONTHS = '{1;2;3;4;5;6;7;8;9;10;11;12}';
  const UNITS = '{33100;47300;69000;102000;150000;220000}';

  it('FORECAST, FORECAST.LINEAR and INTERCEPT fit a straight line to the number pairs', () => {
    check([
      ['FORECAST(30,{6;7;9;15;21},{20;28;31;38;40})', 10.607253086419753],
      ['FORECAST.LINEAR(30,{6;7;9;15;21},{20;28;31;38;40})', 10.607253086419753],
      ['INTERCEPT({2;3;9;1;8},{6;5;11;7;5})', 0.04838709677419217],
      ['FORECAST(4,{1;"a";3},{1;2;3})', 4],
      ['FORECAST({4,5},{1;2},{1;2})', [[4, 5]]],
      ['FORECAST(1,{1;2},{1;2;3})', NA],
      ['FORECAST(1,{1;2},{3;3})', DIV0],
      ['FORECAST(1,{"a"},{"b"})', DIV0],
      ['FORECAST(1,{1;#N/A},{1;2})', NA],
      ['FORECAST(1,{1;2},{1;#N/A})', NA],
      ['INTERCEPT({1;2},{1;2;3})', NA],
    ]);
  });

  it('TREND and GROWTH project the fitted line or curve', () => {
    check([
      [
        `TREND(${COSTS},${MONTHS})`,
        [133953.33333333334, 134971.51515151517, 135989.69696969696, 137007.87878787878, 138026.0606060606, 139044.24242424243, 140062.42424242425, 141080.60606060608, 142098.78787878787, 143116.9696969697, 144135.15151515152, 145153.33333333334].map((v) => [v]),
      ],
      [`TREND(${COSTS},${MONTHS},{13;14;15;16;17})`, [[146171.51515151517], [147189.69696969696], [148207.87878787878], [149226.0606060606], [150244.24242424243]]],
      ['TREND(Offices!E1:E11,Offices!A1:D11,{2500,3,2,25})', 158261.09563260485],
      ['TREND(Offices!E1:E11,Offices!A1:D11,{2500,3,2,25;2310,2,2,20})', [[158261.09563260485], [141650.64968833]]],
      ['TREND({2;4;6},{1;2;3},{4},FALSE())', 8],
      ['TREND({1;2;3})', [[1], [2], [3]]],
      ['TREND({1,2,3})', [[1, 2, 3]]],
      ['TREND({1,2,3},,{4,5})', [[4, 5]]],
      ['TREND({1,2;3,4},{1,2;3,4})', [[1, 2], [3, 4]]],
      ['TREND({1,2,3,4},{1,2,3,4;1,0,1,0},{5;1})', 5],
      ['TREND({1,2,3,4},{1,2,3,4;1,0,1,0},{5,1})', { error: '#REF!' }],
      ['TREND(Offices!E1:E11,Offices!A1:D11,{2500,3})', { error: '#REF!' }],
      ['TREND({1;2;3},{1;2})', { error: '#REF!' }],
      ['TREND({1;"a";3})', VALUE],
      ['TREND({1;2;3},{1;#N/A;3})', NA],
      ['TREND({1;2;3},{1;2;3},{1;#N/A})', NA],
      ['TREND({1,2,3},{1,2})', { error: '#REF!' }],
      ['TREND({1,2;3,4},{1,2,3})', { error: '#REF!' }],
      [`GROWTH(${UNITS},{11;12;13;14;15;16})`, [[32618.20377353971], [47729.42261474776], [69841.30085621739], [102197.07337883174], [149542.48674004574], [218821.87621459612]]],
      [`GROWTH(${UNITS},{11;12;13;14;15;16},{17;18})`, [[320196.71836347267], [468536.0541840476]]],
      ['GROWTH({1;2;0})', NUM],
    ]);
  });

  it('LINEST and LOGEST give the coefficients, constant last, and the statistics of the fit', () => {
    check([
      [
        'LINEST(Offices!E1:E11,Offices!A1:D11,TRUE(),TRUE())',
        [
          [-234.23716447120218, 2553.2106603915443, 12529.768167086753, 27.641387366020326, 52317.83050729132],
          [13.268011475500292, 530.669151930375, 400.0668381939511, 5.429374041545291, 12237.361602862317],
          [0.9967479933845114, 970.5784629285008, NA, NA, NA],
          [459.7536742253976, 6, NA, NA, NA],
          [1732393319.229254, 5652135.316203905, NA, NA, NA],
        ],
      ],
      ['LINEST({1;9;5;7},{0;4;2;3},,FALSE())', [[2, 1]]],
      ['LINEST({1;9;5;7},{0;4;2;3})', [[2, 1]]],
      [`LOGEST(${UNITS},{11;12;13;14;15;16},TRUE(),FALSE())`, [[1.4632756281161785, 495.3047701587281]]],
      ['LINEST({1;2;3},{1;2;3},FALSE(),TRUE())', [[1, 0], [0, NA], [1, 0], [NUM, 2], [14, 0]]],
      ['LINEST({2;4;6},{1;2;3},TRUE(),TRUE())', [[2, 0], [0, 0], [1, 0], [NUM, 1], [8, 0]]],
      ['LINEST({5;5;5},{1;2;3},TRUE(),TRUE())', [[0, 5], [0, 0], [NUM, 0], [NUM, 1], [0, 0]]],
      ['LINEST(1,,TRUE(),TRUE())', [[0, 1], [0, NUM], [NUM, NUM], [NUM, 0], [0, 0]]],
      // A variable that is a multiple of another is dropped, with 0 for its coefficient.
      ['LINEST({2;4;6;8},{1,2;2,4;3,6;4,8})', [[0, 2, 0]]],
      ['LINEST({1,2,3},{1,2,3})', [[1, 0]]],
      ['LOGEST({1;0})', NUM],
    ]);
  });

  it('spill to the size their arguments give', () => {
    const shape = (s: { values: EngineValue[][] }) => s.values.map((r) => r.map((v) => (v === null ? '.' : 'x')).join('')).join(' ');
    expect(shape(spill('=TREND({1;2;3})'))).toBe('x... x... x... ....');
    expect(shape(spill('=TREND({1;2;3},{1;2;3},{4,5})'))).toBe('xx.. .... .... ....');
    expect(shape(spill('=TREND({1;2;3},{1,1;2,0;3,1},{4,1;5,0})'))).toBe('x... x... .... ....');
    expect(shape(spill('=TREND({1,2,3},{1,2,3;1,0,1},{4,5;1,0})'))).toBe('xx.. .... .... ....');
    expect(shape(spill('=LINEST({1;2;3},{1;2;3},TRUE(),TRUE())'))).toBe('xx.. xx.. xx.. xx..');
    expect(shape(spill('=LINEST({1;2;4},{1,2;2,1;3,3})'))).toBe('xxx. .... .... ....');
    expect(shape(spill('=LOGEST(B2:B4)', [[null, 1], [null, 2], [null, 4]]))).toBe('xx.. .... .... ....');
    // Statistics asked for, or not, by a cell's value.
    expect(shape(spill('=LINEST({1;2;4},{1;2;3},TRUE(),C2)', [[null, null, false]]))).toBe('xx.. .... .... ....');
    expect(spill('=TREND(SEQUENCE(B2))', [[null, 3]]).values.map((r) => r[0])).toEqual([1, 2, 3, null]);
    expect(spill('=TREND({1;2},SEQUENCE(B2))', [[null, 2]]).values.map((r) => r[0])).toEqual([1, 2, null, null]);
    expect(spill('=TREND({1;2},{1;2},SEQUENCE(B2))', [[null, 2]]).values.map((r) => r[0])).toEqual([1, 2, null, null]);
    expect(shape(spill('=LINEST({1;2},SEQUENCE(B2))', [[null, 2]]))).toBe('xx.. .... .... ....');
    expect(spill('=TREND({1;2},{1;2;3})').values[0][0]).toMatchObject({ error: '#REF!' });
  });
});

describe('bond and depreciation functions', () => {
  const S = 'DATE(2011,1,25)';
  const M = 'DATE(2011,11,15)';

  it('the COUP functions find the coupon period around settlement, in each day-count basis', () => {
    check([
      [`COUPDAYBS(${S},${M},2,1)`, 71],
      [`COUPDAYS(${S},${M},2,1)`, 181],
      [`COUPDAYSNC(${S},${M},2,1)`, 110],
      [`COUPNCD(${S},${M},2,1)`, 40678],
      [`COUPPCD(${S},${M},2,1)`, 40497],
      [`COUPNUM(${S},${M},2,1)`, 2],
      ['COUPNUM(DATE(2007,1,25),DATE(2008,11,15),2,1)', 4],
      [`COUPNUM(${S},${M},4)`, 4],
      [`COUPNUM(${S},${M},1)`, 1],
      [`COUPDAYBS(${S},${M},2)`, 70],
      [`COUPDAYS(${S},${M},2,0)`, 180],
      [`COUPDAYSNC(${S},${M},2,0)`, 110],
      [`COUPDAYBS(${S},${M},2,2)`, 71],
      [`COUPDAYS(${S},${M},2,2)`, 180],
      [`COUPDAYS(${S},${M},2,3)`, 182.5],
      [`COUPDAYBS(${S},${M},2,4)`, 70],
      // Month-end maturities keep their coupons at month end; others keep their day.
      ['COUPPCD(DATE(2011,3,15),DATE(2011,8,31),2,1)', 40602],
      ['COUPNCD(DATE(2011,3,15),DATE(2011,8,31),2,1)', 40786],
      ['COUPDAYS(DATE(2011,3,15),DATE(2011,8,31),2,1)', 184],
      ['COUPDAYSNC(DATE(2011,3,15),DATE(2011,8,31),2,1)', 169],
      ['COUPPCD(DATE(2011,12,1),DATE(2012,8,30),2,1)', 40785],
      ['COUPPCD(DATE(2011,5,15),DATE(2011,11,15),2,1)', 40678],
      // 30/360 day counts, US and European, at the ends of months.
      ['COUPDAYBS(DATE(2011,3,15),DATE(2011,8,31),2,0)', 15],
      ['COUPDAYBS(DATE(2011,3,15),DATE(2011,8,31),2,4)', 17],
      ['COUPDAYBS(DATE(2011,2,28),DATE(2011,11,30),2,4)', 88],
      ['COUPDAYBS(DATE(2011,1,31),DATE(2011,11,15),2,0)', 76],
      ['COUPDAYBS(DATE(2011,2,28),DATE(2011,11,15),2,0)', 103],
      ['COUPDAYBS(DATE(2011,5,20),DATE(2011,11,15),2,0)', 5],
      [`COUPDAYBS(${M},${S},2,1)`, NUM],
      [`COUPDAYBS(${S},${M},3,1)`, NUM],
      [`COUPDAYBS(${S},${M},2,5)`, NUM],
      [`COUPDAYBS(${S},${M},2,-1)`, NUM],
      [`COUPDAYBS("x",${M},2,1)`, VALUE],
      [`COUPDAYS(-1,${M},2,1)`, NUM],
      [`COUPDAYS(${S},2958466,2,1)`, NUM],
    ]);
    // Dates are serial numbers of the workbook's own date system.
    const e = new NativeEngine([{ name: 'S', rows: [] }], [], true);
    expect(e.calculate(`=COUPNCD(${S},${M},2,1)-DATE(2011,5,15)`, 0)).toBe(0);
    e.destroy();
  });

  it('PRICE and YIELD value a bond and invert each other, with simple interest in the last period', () => {
    check([
      ['PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)', 94.6343616213221],
      ['ROUND(YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0),6)', 0.065],
      ['YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0),100,2,0)', 0.065],
      ['PRICE(DATE(2008,2,15),DATE(2008,5,15),0.0575,0.065,100,2,0)', 99.792512300123],
      ['YIELD(DATE(2008,2,15),DATE(2008,5,15),0.0575,PRICE(DATE(2008,2,15),DATE(2008,5,15),0.0575,0.065,100,2,0),100,2,0)', 0.065],
      ['YIELD(DATE(2008,2,15),DATE(2016,11,15),0,95,100,2,0)', 0.00587069021396558],
      ['YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,10^10,100,2,0)', NUM],
      ['PRICE(DATE(2008,2,15),DATE(2017,11,15),-0.1,0.065,100,2,0)', NUM],
      ['PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,-0.1,100,2,0)', NUM],
      ['PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,0,2,0)', NUM],
      ['PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,3,0)', NUM],
      ['YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,0,100,2,0)', NUM],
      ['YIELD(DATE(2008,2,15),DATE(2016,11,15),-0.1,95,100,2,0)', NUM],
      ['YIELD(DATE(2016,2,15),DATE(2008,11,15),0.0575,95,100,2,0)', NUM],
    ]);
  });

  it('DURATION and MDURATION', () => {
    check([
      ['DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1)', 10.9191452815919],
      ['MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)', 5.73566981391884],
      ['DURATION(DATE(2018,7,1),DATE(2048,1,1),-0.08,0.09,2,1)', NUM],
      ['MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,-0.09,2,1)', NUM],
      ['DURATION(DATE(2048,7,1),DATE(2018,1,1),0.08,0.09,2,1)', NUM],
    ]);
  });

  it('ACCRINT accrues over each coupon period, from issue or from the first interest date', () => {
    const DOC = 'DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2';
    check([
      [`ACCRINT(${DOC},0)`, 16.666666666666668],
      ['ACCRINT(DATE(2008,3,5),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0,FALSE())', 15.555555555555555],
      ['ACCRINT(DATE(2008,4,5),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0,TRUE())', 7.222222222222222],
      [`ACCRINT(${DOC},1)`, 16.576086956521738],
      [`ACCRINT(${DOC},2)`, 16.944444444444443],
      [`ACCRINT(${DOC},3)`, 16.71232876712329],
      [`ACCRINT(${DOC},4)`, 16.666666666666668],
      ['ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,,2)', 16.666666666666668],
      ['ACCRINT(DATE(2020,1,1),DATE(2020,7,1),DATE(2021,3,1),0.1,1000,2,0,TRUE())', 116.66666666666667],
      ['ACCRINT(DATE(2020,1,1),DATE(2020,7,1),DATE(2021,3,1),0.1,1000,2,0,FALSE())', 66.66666666666667],
      ['ACCRINT(DATE(2019,10,1),DATE(2020,7,1),DATE(2020,3,1),0.1,1000,2,0)', 41.666666666666664],
      ['ACCRINT(DATE(2009,3,15),DATE(2008,8,31),DATE(2009,5,15),0.1,1000,2,0)', 16.666666666666668],
      ['ACCRINT(DATE(2008,5,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2)', NUM],
      [`ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0,1000,2)`, NUM],
      [`ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,0,2)`, NUM],
      [`ACCRINT(${DOC},7)`, NUM],
      ['ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,3)', NUM],
    ]);
  });

  it('ACCRINTM and the discounted securities count the fraction of a year as YEARFRAC does', () => {
    check([
      ['ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,3)', 20.54794520547945],
      ['ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,1)', 20.491803278688526],
      ['ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1)', 20.555555555555557],
      // US 30/360 at the ends of months, and actual/actual across years.
      ['ACCRINTM(DATE(2010,1,31),DATE(2010,3,31),0.36,1000,0)', 60],
      ['ACCRINTM(DATE(2010,1,31),DATE(2010,3,15),0.36,1000,0)', 45],
      ['ACCRINTM(DATE(2010,1,30),DATE(2010,3,31),0.36,1000,0)', 60],
      ['ACCRINTM(DATE(2010,2,28),DATE(2011,2,28),0.36,1000,0)', 360],
      ['ACCRINTM(DATE(2010,2,28),DATE(2010,3,31),0.36,1000,0)', 31],
      ['ACCRINTM(DATE(2010,1,31),DATE(2010,3,31),0.36,1000,4)', 60],
      ['ACCRINTM(DATE(2010,1,1),DATE(2012,1,1),0.1,1000,1)', 199.81751824817516],
      ['ACCRINTM(DATE(2011,6,1),DATE(2012,5,1),0.1,1000,1)', 91.53005464480874],
      ['ACCRINTM(DATE(2011,3,1),DATE(2012,2,29),0.1,1000,1)', 99.72677595628416],
      ['ACCRINTM(DATE(2012,2,1),DATE(2013,1,15),0.1,1000,1)', 95.35519125683061],
      ['ACCRINTM(DATE(2011,1,1),DATE(2011,7,1),0.1,1000,1)', 49.58904109589041],
      ['ACCRINTM(DATE(2008,6,15),DATE(2008,4,1),0.1,1000,3)', NUM],
      ['ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0,1000,3)', NUM],
      ['ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,5)', NUM],
      ['ACCRINTM(-5,DATE(2008,6,15),0.1)', NUM],
      ['DISC(DATE(2018,7,1),DATE(2048,1,1),97.975,100,1)', 0.000686384169121347],
      ['INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,2)', 0.05768],
      ['RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,0.0575,2)', 1014584.6544071],
      ['PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)', 99.79583333333333],
      ['YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)', 0.0528225719868583],
      ['PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)', 99.9844988755569],
      ['PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2008,2,15),0.061,0.061,0)', 100],
      ['PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2008,3,1),0.061,0.061,0)', 99.4604162670447],
      ['YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)', 0.0609543336915387],
      ['RECEIVED(DATE(2008,2,15),DATE(2018,5,15),1000,0.5,2)', NUM],
      ['DISC(DATE(2048,1,1),DATE(2018,7,1),97.975,100,1)', NUM],
      ['DISC(DATE(2018,7,1),DATE(2048,1,1),0,100,1)', NUM],
      ['INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,9)', NUM],
      ['PRICEMAT(DATE(2008,4,13),DATE(2008,2,15),DATE(2007,11,11),0.061,0.061,0)', NUM],
      ['PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),-0.061,0.061,0)', NUM],
      ['YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,0,0)', NUM],
      ['YIELDMAT(DATE(2008,11,3),DATE(2008,3,15),DATE(2007,11,8),0.0625,100,0)', NUM],
    ]);
  });

  it('VDB depreciates by declining balance, switching to straight-line unless told not to', () => {
    check([
      ['VDB(2400,300,10*365,0,1)', 1.3150684931506849],
      ['VDB(2400,300,10*12,0,1)', 40],
      ['VDB(2400,300,10,0,1)', 480],
      ['VDB(2400,300,10*12,6,18)', 396.306053264752],
      ['VDB(2400,300,10*12,6,18,1.5)', 311.808936658233],
      ['VDB(2400,300,10,0,0.875,1.5)', 315],
      ['VDB(2400,300,10,0,3,2,TRUE())', 1171.2],
      ['VDB(2400,300,10,0,3)', 1171.2],
      ['VDB(2400,300,10,0.5,2.5,2,TRUE())', 777.6],
      ['VDB(2400,300,10,0,10)', 2100],
      ['VDB(2400,300,10,8,10)', 102.653184],
      ['VDB(2400,300,10,1.5,3.25)', 560.64],
      ['VDB(1000,100,2,0,1,2)', 900],
      ['VDB(1000,100,2,1,2,2)', 0],
      ['VDB(2400,300,10,2,1)', NUM],
      ['VDB(2400,300,10,0,11)', NUM],
      ['VDB(2400,3000,10,0,1)', NUM],
      ['VDB(2400,300,10,0,1,0)', NUM],
      ['VDB(-1,0,10,0,1)', NUM],
      ['VDB(1,0,2000000,0,2000000)', NUM],
    ]);
  });
});

describe('large inputs', () => {
  it('are worked through without overflowing the stack', () => {
    const rows = Array.from({ length: 200_000 }, (_, i) => [i % 1000, i]);
    const e = new NativeEngine([{ name: 'S', rows }], [], false);
    const at = (f: string) => e.calculate('=' + f, 0);
    // An array of keys looked up in one formula searches the range once, not once per key.
    expect(at('SUM(XMATCH(A1:A20000,B1:B200000))')).toBe(10_010_000);
    expect(at('SUM(LOOKUP(A1:A20000,B1:B200000))')).toBe(9_990_000);
    expect(at('MODE(A1:A200000)')).toBe(0);
    expect(at('AGGREGATE(4,6,A1:A200000)')).toBe(999);
    expect(at('AGGREGATE(5,6,A1:A200000)')).toBe(0);
    expect(at('PERCENTOF(A1:A1000,A1:A200000)')).toBeCloseTo(1 / 200);
    expect(at('ROWS(TEXTSPLIT(REPT("a;",16000),,";"))')).toBe(16_001);
    e.destroy();
  });

  it('a whole column of an empty sheet counts as an empty cell rather than breaking the engine', () => {
    const formulas = [
      'TAKE(E!X:X,2)',
      'DROP(E!X:X,1)',
      'TOCOL(E!X:X)',
      'TOROW(E!X:X,1)',
      'CHOOSEROWS(E!X:X,1)',
      'CHOOSECOLS(E!X:X,1)',
      'WRAPROWS(E!X:X,2)',
      'EXPAND(E!X:X,2)',
      'SORTBY(E!X:X,E!Y:Y)',
      'TEXTSPLIT(E!X:X,",")',
      'ARRAYTOTEXT(E!X:X)',
      'LOOKUP(1,E!X:X)',
      'XMATCH(1,E!X:X)',
      'FREQUENCY(E!X:X,E!Y:Y)',
      'TREND(E!X:X)',
      'LINEST(E!X:X)',
      'MDETERM(E!X:X)',
      'MINVERSE(E!X:X)',
      'AGGREGATE(9,0,E!X:X)',
      'TRIMMEAN(E!X:X,0)',
      'PERCENTRANK(E!X:X,1)',
      'MODE(E!X:X)',
      'AVERAGEIFS(E!X:X,E!Y:Y,1)',
      'PROB(E!X:X,E!Y:Y,1)',
      'FORECAST(1,E!X:X,E!Y:Y)',
    ];
    const e = new NativeEngine([{ name: 'S', rows: [[1]] }, { name: 'E', rows: [] }], [], false);
    e.write(formulas.map((f, i) => ({ sheet: 0, row: i + 1, col: 2, raw: '=' + f })));
    for (let i = 0; i < formulas.length; i++) expect(e.value(0, i + 1, 2), formulas[i]).not.toBeUndefined();
    e.destroy();
  });
});
