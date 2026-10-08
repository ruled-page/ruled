import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { CalcEngine, type EngineValue } from '../src/core/calc';

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

let engine: CalcEngine;
beforeAll(() => {
  engine = new CalcEngine([{ name: 'S', rows: ROWS }], [], false);
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
const CALC = { error: '#CALC!' };

/** A sheet where a formula is entered in A1 and spills, so the cells it fills can be read. */
function spill(formula: string, rows: (string | number | null)[][] = []) {
  const e = new CalcEngine([{ name: 'T', rows: [[formula], ...rows] }], [], false);
  try {
    const out: EngineValue[][] = [];
    for (let r = 0; r < 4; r++) out.push([0, 1, 2, 3].map((c) => e.value(0, r, c)));
    return { values: out, isArray: e.isArray(0, 0, 0) };
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
      ['LOOKUP(D4,{0,1},{"zero","one"})', 'zero'],
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
      ['XMATCH(D4,{1,0})', 2],
      ['XMATCH(DATE(2026,2,1),E1:E2)', 2],
      ['XMATCH(5.77,A1:A5,2,2)', 4],
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
      ['TYPE()', NA],
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
    // The engine's circular-reference error has no number and passes through.
    const e = new CalcEngine([{ name: 'T', rows: [['=A1', '=ERROR.TYPE(A1)']] }], [], false);
    expect(e.value(0, 0, 1)).toMatchObject({ error: '#CYCLE!' });
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

  it('TEXTSPLIT spills by the size of its literal arguments', () => {
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
    expect(spill('=TEXTSPLIT(("a")&("-b"),"-")').values[0][0]).toMatchObject({ error: '#VALUE!' });
    // With text that is only known once calculated, the size can't be predicted: one part fits, more can't spill.
    expect(spill('=TEXTSPLIT(A2,",")', [['solo']]).values[0][0]).toBe('solo');
    expect(spill('=TEXTSPLIT(A2,",")', [['a,b']]).values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=TEXTSPLIT("a,b",B2)', [[null, ',']]).values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=TEXTSPLIT(UPPER("a,b"),",")').values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=TEXTSPLIT("a,b",",",,ABS(0))').values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=TEXTSPLIT(IF(TRUE(),"a,b",),",")').values[0][0]).toMatchObject({ error: '#VALUE!' });
    // A size that hangs on another formula's values (here SEQUENCE's) is never guessed.
    expect(spill('=TEXTSPLIT(SEQUENCE(B2),",")', [[null, 2]]).values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=TEXTSPLIT(A2:AN30000,",")').isArray).toBe(false);
    expect(spill('=TEXTSPLIT("a,b","")').values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=TEXTSPLIT(",",",",,TRUE())').values[0][0]).toMatchObject({ error: '#CALC!' });
    expect(spill('=TEXTSPLIT()').values[0][0]).toMatchObject({ error: '#N/A' });
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
      [`TAKE(${GRID},0)`, CALC],
      [`TAKE(${GRID},1,0)`, CALC],
      ['TAKE(1/0,1)', DIV0],
      [`DROP(${GRID},2)`, [[7, 8, 9]]],
      [`DROP(${GRID},,2)`, [[3], [6], [9]]],
      [`DROP(${GRID},-2)`, [[1, 2, 3]]],
      [`DROP(${GRID},2,2)`, 9],
      [`DROP(${GRID},0)`, [[1, 2, 3], [4, 5, 6], [7, 8, 9]]],
      [`DROP(${GRID},3)`, CALC],
      [`DROP(${GRID},,-5)`, CALC],
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
    expect(calc('TOROW(D1:D5,2)')).toEqual([[1, 'two', true, null]]);
  });

  it('WRAPROWS and WRAPCOLS wrap a row or column, padding the last line', () => {
    check([
      ['WRAPROWS({"A","B","C","D","E","F","G"},3)', [['A', 'B', 'C'], ['D', 'E', 'F'], ['G', NA, NA]]],
      ['WRAPROWS({"A","B","C","D","E","F","G"},3,"x")', [['A', 'B', 'C'], ['D', 'E', 'F'], ['G', 'x', 'x']]],
      ['WRAPCOLS({"A","B","C","D","E","F","G"},3)', [['A', 'D', 'G'], ['B', 'E', NA], ['C', 'F', NA]]],
      ['WRAPROWS(A1:A5,2.5)', [[4.14, 4.19], [5.17, 5.77], [6.39, NA]]],
      ['WRAPROWS({1,2},5)', [[1, 2, NA, NA, NA]]],
      ['WRAPROWS({1,2;3,4},2)', VALUE],
      ['WRAPROWS({1,2},0)', { error: '#NUM!' }],
      ['WRAPCOLS({1,2},0)', { error: '#NUM!' }],
    ]);
  });

  it('EXPAND pads an array out to a size', () => {
    check([
      ['EXPAND({1,2;3,4},3,3)', [[1, 2, NA], [3, 4, NA], [NA, NA, NA]]],
      ['EXPAND({1,2;3,4},3,,"-")', [[1, 2], [3, 4], ['-', '-']]],
      ['EXPAND({1,2;3,4},,3,0)', [[1, 2, 0], [3, 4, 0]]],
      ['EXPAND({1,2;3,4},1)', VALUE],
      ['EXPAND({1,2;3,4},2,1)', VALUE],
      ['EXPAND(1,2000,1000)', { error: '#NUM!' }],
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
      ['RANDARRAY(2000,1000)', { error: '#NUM!' }],
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
    expect(column(spill('=TAKE({1;2;3},(2))'))).toEqual([1, 2, null, null]);
    expect(column(spill('=TAKE({1;2;3},-B2)', [[null, 1]]))).toEqual([3, null, null, null]);
    expect(spill('=TAKE({1,2;3,4},1,B2)', [[null, 1]]).values[0]).toEqual([1, null, null, null]);
    expect(spill('=TAKE({1;2;3},"a")').values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=RANDARRAY(2000,1000)').values[0][0]).toMatchObject({ error: '#NUM!' });
    // A count from a cell takes the whole input's room, the cells past the result left empty.
    expect(column(spill('=TAKE(C2:C4,D2)', [[null, null, 10, 2], [null, null, 20], [null, null, 30]]))).toEqual([10, 20, null, null]);
    expect(column(spill('=TOCOL(C2:C4,1)', [[null, null, 10], [null, null, null], [null, null, 30]]))).toEqual([10, 30, null, null]);
    // Sizes that can't be known before calculating don't spill.
    expect(spill('=TAKE(SEQUENCE(B2),2)', [[null, 3]]).values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=WRAPROWS({1,2,3},B2)', [[null, 2]]).values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=EXPAND(1,B2)', [[null, 2]]).values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=RANDARRAY(B2)', [[null, 2]]).values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=CHOOSEROWS(A2:A3,SEQUENCE(B2))', [[1, 2], [3]]).values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=EXPAND({1,2},1,1)').values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=WRAPROWS({1,2;3,4},2)').values[0][0]).toMatchObject({ error: '#VALUE!' });
    expect(spill('=DROP({1;2},2)').values[0][0]).toMatchObject({ error: '#CALC!' });
    expect(spill('=TAKE({1;2},0)').values[0][0]).toMatchObject({ error: '#CALC!' });
    expect(spill('=CHOOSEROWS({1;2})').values[0][0]).toMatchObject({ error: '#N/A' });
    for (const f of ['=TOCOL()', '=TOROW()', '=SORTBY()', '=WRAPROWS()', '=WRAPCOLS()', '=EXPAND()', '=TAKE()', '=DROP()']) {
      expect(spill(f).values[0][0]).toMatchObject({ error: expect.stringMatching(/^#/) });
    }
  });
});
