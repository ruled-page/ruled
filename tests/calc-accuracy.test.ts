import ExcelJS from 'exceljs';
import { beforeAll, describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import type { CalcReport } from '../src/core/types';

/**
 * Formulas with the results Excel gives, stored as each cell's saved value,
 * the way Excel writes them. Opening the workbook compares the engine with
 * every one, exactly as it does for real files, so the report says where the
 * engine and Excel agree. Data!A1:A10 = 1..10; B1:B5 = apple, Banana, cherry,
 * apple, date; E1:F4 = a 1, b 2, c 3, d 4; E6 is empty.
 */
type Expected = number | string | boolean | { error: string };
const CASES: [string, Expected][] = [
  // Arithmetic and rounding
  ['SUM(1,2,3)', 6],
  ['SUM("1",2)', 3],
  ['10/3', 10 / 3],
  ['0.1+0.2=0.3', true],
  ['ROUND(2.5,0)', 3],
  ['ROUND(-2.5,0)', -3],
  ['ROUND(1.005,2)', 1.01],
  ['ROUNDUP(3.14159,2)', 3.15],
  ['ROUNDDOWN(-3.14159,1)', -3.1],
  ['MROUND(10,3)', 9],
  ['CEILING(4.2,1)', 5],
  ['FLOOR(4.8,1)', 4],
  ['INT(-1.5)', -2],
  ['INT(5.9)', 5],
  ['ROUND(1234.5,-2)', 1200],
  ['ROUNDUP(-1.21,1)', -1.3],
  ['ROUNDDOWN(1.29,1)', 1.2],
  ['ROUND(2.675,2)', 2.68],
  ['MOD(5.5,2)', 1.5],
  ['MOD(-7,-3)', -1],
  ['MOD(1,0)', { error: '#DIV/0!' }],
  ['TRUNC(-1.5)', -1],
  ['MOD(-3,2)', 1],
  ['MOD(3,-2)', -1],
  ['POWER(2,10)', 1024],
  ['SQRT(16)', 4],
  ['PRODUCT(2,3,4)', 24],
  ['QUOTIENT(7,2)', 3],
  ['Data!E6+1', 1],
  // Aggregates and conditions
  ['SUMPRODUCT(Data!A1:A3,Data!A2:A4)', 20],
  ['SUMIF(Data!A1:A10,">5")', 40],
  ['SUMIFS(Data!A1:A10,Data!A1:A10,">=3",Data!A1:A10,"<=5")', 12],
  ['COUNTIF(Data!B1:B5,"apple")', 2],
  ['COUNTIF(Data!B1:B5,"a*")', 2],
  ['COUNTIF(Data!B1:B5,"*an*")', 1],
  ['COUNTA(Data!B1:B5)', 5],
  ['COUNTBLANK(Data!E5:E6)', 2],
  ['AVERAGE(Data!A1:A10)', 5.5],
  ['MEDIAN(1,2,3,4)', 2.5],
  ['STDEV.S(Data!A1:A10)', 3.0276503540974917],
  ['LARGE(Data!A1:A10,2)', 9],
  ['SMALL(Data!A1:A10,3)', 3],
  ['RANK.EQ(7,Data!A1:A10)', 4],
  ['RANK.EQ(7,Data!A1:A10,1)', 7],
  ['RANK(3,Data!A1:A10)', 8],
  ['RANK.EQ(11,Data!A1:A10)', { error: '#N/A' }],
  // Text
  ['LEFT("Hello",2)', 'He'],
  ['MID("Hello",2,3)', 'ell'],
  ['RIGHT("Hello",3)', 'llo'],
  ['LEN("Hello")', 5],
  ['UPPER("abc")', 'ABC'],
  ['PROPER("hello world")', 'Hello World'],
  ['TRIM("  a  b  ")', 'a b'],
  ['SUBSTITUTE("a-b-c","-","+")', 'a+b+c'],
  ['SUBSTITUTE("a-b-c","-","+",2)', 'a-b+c'],
  ['FIND("l","Hello")', 3],
  ['SEARCH("L","Hello")', 3],
  ['TEXT(1234.567,"#,##0.00")', '1,234.57'],
  ['TEXT(0.256,"0.0%")', '25.6%'],
  ['TEXT(DATE(2026,3,1),"yyyy-mm-dd")', '2026-03-01'],
  ['TEXT(DATE(2026,3,1),"dddd")', 'Sunday'],
  ['TEXT(-1234.5,"#,##0;(#,##0)")', '(1,235)'],
  ['TEXT(45352,"d mmm yyyy")', '1 Mar 2024'],
  ['TEXT(0.75,"h:mm AM/PM")', '6:00 PM'],
  ['TEXT("abc","@")', 'abc'],
  ['CONCATENATE("a","b",1)', 'ab1'],
  ['"a"&1&TRUE', 'a1TRUE'],
  ['VALUE("12.5")', 12.5],
  ['REPT("ab",3)', 'ababab'],
  ['EXACT("a","A")', false],
  // Logic and errors
  ['IF(1>2,"y","n")', 'n'],
  ['AND(TRUE,1)', true],
  ['OR(FALSE,0)', false],
  ['NOT(0)', true],
  ['IFERROR(1/0,"div")', 'div'],
  ['IFS(FALSE,1,TRUE,2)', 2],
  ['SWITCH(2,1,"a",2,"b")', 'b'],
  ['ISBLANK(Data!E6)', true],
  ['ISNUMBER("1")', false],
  ['1/0', { error: '#DIV/0!' }],
  ['NA()', { error: '#N/A' }],
  ['"a"+1', { error: '#VALUE!' }],
  // Dates (serial numbers in the 1900 system)
  ['DATE(2024,2,30)', 45352],
  ['DATE(2026,1,1)', 46023],
  ['YEAR(46023)', 2026],
  ['MONTH(45352)', 3],
  ['DAY(45352)', 1],
  ['EOMONTH(DATE(2026,1,15),1)', 46081],
  ['EDATE(DATE(2026,1,31),1)', 46081],
  ['WEEKDAY(DATE(2026,3,1))', 1],
  ['WEEKDAY(DATE(2026,3,1),2)', 7],
  ['NETWORKDAYS(DATE(2026,3,2),DATE(2026,3,13))', 10],
  ['WORKDAY(DATE(2026,3,6),1)', 46090],
  ['DATEDIF(DATE(2020,1,15),DATE(2026,3,1),"y")', 6],
  ['DATEDIF(DATE(2026,1,31),DATE(2026,3,1),"m")', 1],
  ['DATEDIF(DATE(2020,1,15),DATE(2026,3,1),"ym")', 1],
  ['DATEDIF(DATE(2025,6,15),DATE(2026,3,1),"yd")', 259],
  ['DATEDIF(DATE(2026,3,1),DATE(2026,1,1),"d")', { error: '#NUM!' }],
  ['DAYS(DATE(2026,3,1),DATE(2026,1,1))', 59],
  ['YEARFRAC(DATE(2026,1,1),DATE(2026,7,1))', 0.5],
  ['TIME(12,30,0)', 0.5208333333333334],
  ['HOUR(0.75)', 18],
  // Lookups
  ['VLOOKUP("c",Data!E1:F4,2,FALSE)', 3],
  ['VLOOKUP("C",Data!E1:F4,2,FALSE)', 3],
  ['INDEX(Data!F1:F4,MATCH("b",Data!E1:E4,0))', 2],
  ['MATCH(5,{1,3,5,7},0)', 3],
  ['MATCH(6,{1,3,5,7},1)', 3],
  ['CHOOSE(2,"a","b","c")', 'b'],
  ['XLOOKUP("c",Data!E1:E4,Data!F1:F4)', 3],
  // Finance (Excel's documented examples)
  ['PMT(0.05/12,360,200000)', -1073.6432460242797],
  ['FV(0.06/12,10,-200,-500,1)', 2581.4033740601],
  ['NPV(0.1,-10000,3000,4200,6800)', 1188.4434123352207],
  ['IRR({-70000,12000,15000,18000,21000})', -0.021244848272],
];

/**
 * Where this editor's engine (HyperFormula) is known to differ from Excel.
 * Opening a real file keeps Excel's saved value for these and marks them.
 */
const KNOWN_DIFFERENCES = new Set<string>([]);

async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet('Data');
  for (let i = 1; i <= 10; i++) data.getCell(`A${i}`).value = i;
  ['apple', 'Banana', 'cherry', 'apple', 'date'].forEach((v, i) => (data.getCell(`B${i + 1}`).value = v));
  ['a', 'b', 'c', 'd'].forEach((v, i) => {
    data.getCell(`E${i + 1}`).value = v;
    data.getCell(`F${i + 1}`).value = i + 1;
  });
  const calc = wb.addWorksheet('Calc');
  CASES.forEach(([formula, result], i) => {
    calc.getCell(`A${i + 1}`).value = formula;
    calc.getCell(`B${i + 1}`).value = { formula, result } as ExcelJS.CellFormulaValue;
  });
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let report: CalcReport;
beforeAll(async () => {
  const s = WorkbookSession.open(await workbook(), 'accuracy.xlsx');
  report = s.ensureEngine();
});

describe('calculation accuracy', () => {
  it('matches Excel on every formula in the catalogue, apart from the known differences', () => {
    const differ = report.samples.map((x) => `${x.formula} → Excel ${x.excel ?? ''}, here ${x.engine ?? x.reason}`);
    const unexpected = report.samples.filter((x) => !KNOWN_DIFFERENCES.has(x.formula));
    expect(unexpected.map((x) => `${x.formula} → Excel ${x.excel ?? ''}, here ${x.engine ?? x.reason}`)).toEqual([]);
    expect(report.formulaCells).toBe(CASES.length);
    expect(report.liveCells).toBe(CASES.length - differ.length);
  });
});
