import { describe, expect, it } from 'vitest';
import { createWorkbook, dateToSerial, serialToDate } from '../src';

const prices = () =>
  createWorkbook({
    sheets: {
      Prices: [
        ['Tea', 4.5],
        ['Cake', 3],
        ['Total', '=SUM(B1:B2)'],
      ],
      'Q1 Sales': [['=SEQUENCE(3)', null, '=Prices!B3*2']],
    },
    names: { Rate: '=Prices!$B$1' },
  });

describe('a workbook on its own', () => {
  it('works out the formulas it is made with, spilling those whose results are arrays', () => {
    const book = prices();
    expect(book.sheets).toEqual(['Prices', 'Q1 Sales']);
    expect(book.get('Prices!B3')).toBe(7.5);
    expect(book.get("'Q1 Sales'!C1")).toBe(15);
    expect(book.getRange("'q1 sales'!A1:A4")).toEqual([[1], [2], [3], [null]]);
    expect(book.get('A1')).toBe('Tea');
    expect(book.formula('Prices!B3')).toBe('=SUM(B1:B2)');
    expect(book.formula('Prices!B1')).toBeUndefined();
    expect(createWorkbook().sheets).toEqual(['Sheet1']);
  });

  it('reports every value a change makes, by cell', () => {
    const book = prices();
    expect(book.set('Prices!B1', 10)).toEqual([
      { cell: 'Prices!B1', value: 10 },
      { cell: 'Prices!B3', value: 13 },
      { cell: "'Q1 Sales'!C1", value: 26 },
    ]);
    expect(book.set({ 'Prices!B1': 1, 'Prices!B2': 2 })).toContainEqual({ cell: 'Prices!B3', value: 3 });
    book.set('Prices!D1', [
      [1, 2],
      [3, '=D1+E1+D2'],
    ]);
    expect(book.get('Prices!E2')).toBe(6);
  });

  it('spills as Excel does: blocked, cleared, grown and shrunk', () => {
    const book = prices();
    expect(book.set("'Q1 Sales'!A2", 'x')).toEqual([
      { cell: "'Q1 Sales'!A2", value: 'x' },
      { cell: "'Q1 Sales'!A3", value: null },
      { cell: "'Q1 Sales'!A1", value: { error: '#SPILL!' } },
    ]);
    book.set("'Q1 Sales'!A2", null);
    expect(book.getRange("'Q1 Sales'!A1:A3")).toEqual([[1], [2], [3]]);
    book.set("'Q1 Sales'!A1", '=SEQUENCE(4)');
    expect(book.get("'Q1 Sales'!A4")).toBe(4);
    expect(book.set("'Q1 Sales'!A1", '=SEQUENCE(2)')).toEqual(
      expect.arrayContaining([
        { cell: "'Q1 Sales'!A3", value: null },
        { cell: "'Q1 Sales'!A4", value: null },
      ]),
    );
    // A result that becomes an array when what it reads changes.
    book.set('Prices!E1', '=FILTER(B1:B2,B1:B2>4)');
    expect(book.set('Prices!B2', 5)).toContainEqual({ cell: 'Prices!E2', value: 5 });
    expect(book.set('Prices!B2', 1)).toContainEqual({ cell: 'Prices!E2', value: null });
    // A value over a spilling formula ends its spill.
    expect(book.set("'Q1 Sales'!A1", 7)).toContainEqual({ cell: "'Q1 Sales'!A2", value: null });
  });

  it('sizes stacked arrays as their parts grow and shrink', () => {
    const NA = { error: '#N/A' };
    const book = createWorkbook({ sheets: { S: [[2, 'a', '=VSTACK(SEQUENCE(A1),HSTACK(B1,"b"))', null, '=HSTACK(SEQUENCE(1,A1),VSTACK(B1,"c"))']] } });
    expect(book.getRange('C1:D3')).toEqual([[1, NA], [2, NA], ['a', 'b']]);
    expect(book.getRange('E1:G2')).toEqual([[1, 2, 'a'], [NA, NA, 'c']]);
    book.set('A1', 4);
    expect(book.getRange('C1:D6')).toEqual([[1, NA], [2, NA], [3, NA], [4, NA], ['a', 'b'], [null, null]]);
    expect(book.getRange('E1:J2')).toEqual([[1, 2, 3, 4, 'a', null], [NA, NA, NA, NA, 'c', null]]);
    book.set('A1', 1);
    expect(book.getRange('C1:D3')).toEqual([[1, NA], ['a', 'b'], [null, null]]);
    expect(book.getRange('E1:H2')).toEqual([[1, 'a', null, null], [NA, 'c', null, null]]);
  });

  it('works out formulas without putting them in cells', () => {
    const book = prices();
    expect(book.evaluate('=Rate*2')).toBe(9);
    // A workbook's name read through a sheet that has none of its own.
    expect(book.evaluate("='Q1 Sales'!Rate")).toBe(4.5);
    expect(book.evaluate('=B1+1', 'Prices')).toBe(5.5);
    expect(book.evaluate('=SEQUENCE(2,2)')).toEqual([
      [1, 2],
      [3, 4],
    ]);
    expect(book.evaluate('=LET(x,2,x*3)')).toBe(6);
    expect(book.evaluate('=MAP(SEQUENCE(3),LAMBDA(v,v*2))')).toEqual([[2], [4], [6]]);
    expect(book.evaluate('=XLOOKUP("Cake",Prices!A1:A2,Prices!B1:B2)')).toBe(3);
    expect(book.evaluate('=NOTAFUNCTION(1)')).toEqual({ error: '#NAME?' });
  });

  it('defines, changes and removes names', () => {
    const book = prices();
    book.set('Prices!C1', '=Rate*10');
    expect(book.setName('Rate', 'Prices!$B$2')).toEqual([{ cell: 'Prices!C1', value: 30 }]);
    expect(book.setName('Rate', null)).toEqual([{ cell: 'Prices!C1', value: { error: '#NAME?' } }]);
    expect(book.setName('Rate', 'Prices!#REF!')).toEqual([{ cell: 'Prices!C1', value: { error: '#REF!' } }]);
  });

  it('adds sheets, which formulas that named them read', () => {
    const book = prices();
    expect(book.set('Prices!A5', '=Later!A1+1')).toEqual([{ cell: 'Prices!A5', value: { error: '#REF!' } }]);
    expect(book.addSheet('Later', [[41, '=SEQUENCE(1,2)']])).toEqual(expect.arrayContaining([{ cell: 'Prices!A5', value: 42 }, { cell: 'Later!C1', value: 2 }]));
    expect(book.sheets).toEqual(['Prices', 'Q1 Sales', 'Later']);
    expect(() => book.addSheet('later')).toThrow('already a sheet named');
    expect(() => book.addSheet('a/b')).toThrow("isn't a sheet name");
  });

  it('says what is wrong with a reference', () => {
    const book = prices();
    expect(() => book.get('Nowhere!A1')).toThrow('no sheet named "Nowhere"');
    expect(() => book.get('Prices!1A')).toThrow("isn't a cell or range reference");
    expect(() => book.get('XFE1')).toThrow("isn't a cell or range reference");
    expect(() => book.get('Prices!A1:B2')).toThrow('getRange');
  });

  it('takes dates as Excel stores them', () => {
    const book = createWorkbook({ sheets: { S: [[new Date('2026-10-13T12:00:00Z'), '=TEXT(A1,"yyyy-mm-dd hh:mm")']] } });
    expect(book.get('A1')).toBe(46308.5);
    expect(book.get('B1')).toBe('2026-10-13 12:00');
    expect(createWorkbook({ date1904: true, sheets: { S: [[new Date('2026-10-13')]] } }).get('A1')).toBe(44846);
    expect(dateToSerial(new Date('1900-02-28'))).toBe(59);
    expect(dateToSerial(new Date('1900-03-01'))).toBe(61);
    expect(serialToDate(46308.5).toISOString()).toBe('2026-10-13T12:00:00.000Z');
    expect(serialToDate(44846, true).toISOString()).toBe('2026-10-13T00:00:00.000Z');
  });

  it('inserts and deletes rows, references following the cells as Excel moves them', () => {
    const book = prices();
    book.set('Prices!C3', '=ROW()');
    expect(book.insertRows('Prices!2')).toEqual([{ cell: 'Prices!C4', value: 4 }]);
    expect(book.getRange('Prices!A1:B4')).toEqual([
      ['Tea', 4.5],
      [null, null],
      ['Cake', 3],
      ['Total', 7.5],
    ]);
    expect(book.formula('Prices!B4')).toBe('=SUM(B1:B3)');
    expect(book.formula("'Q1 Sales'!C1")).toBe('=Prices!B4*2');
    expect(book.set('Prices!B2', 1)).toEqual([
      { cell: 'Prices!B2', value: 1 },
      { cell: 'Prices!B4', value: 8.5 },
      { cell: "'Q1 Sales'!C1", value: 17 },
    ]);
    expect(book.deleteRows('Prices!1:2')).toEqual([
      { cell: 'Prices!B2', value: 3 },
      { cell: 'Prices!C2', value: 2 },
      { cell: "'Q1 Sales'!C1", value: 6 },
    ]);
    expect(book.evaluate('=Rate')).toEqual({ error: '#REF!' });
    expect(() => book.insertRows('Prices!B')).toThrow('isn\'t a reference to rows');
  });

  it('inserts and deletes columns', () => {
    const book = prices();
    expect(book.insertColumns('Prices!A')).toEqual([]);
    expect(book.formula('Prices!C3')).toBe('=SUM(C1:C2)');
    expect(book.formula("'Q1 Sales'!C1")).toBe('=Prices!C3*2');
    expect(book.evaluate('=Rate')).toBe(4.5);
    expect(book.deleteColumns('Prices!C')).toEqual([{ cell: "'Q1 Sales'!C1", value: { error: '#REF!' } }]);
    expect(book.formula("'Q1 Sales'!C1")).toBe('=#REF!*2');
    expect(() => book.deleteColumns('Prices!3')).toThrow('isn\'t a reference to columns');
  });

  it('reports what a host shifting its cells by the rows inserted needs to know to show them as they are', () => {
    const book = createWorkbook({ sheets: { S: [['=SEQUENCE(3)', 'a', '=ROWS(A1:A3)'], [null, 'b'], [null, 'c']] } });
    const shown = book.getRange('A1:C5');
    shown.splice(1, 0, [null, null, null]);
    for (const { cell, value } of book.insertRows('2')) {
      const at = /^S!([A-Z])(\d+)$/.exec(cell)!;
      shown[Number(at[2]) - 1][at[1].charCodeAt(0) - 65] = value;
    }
    expect(shown.slice(0, 5)).toEqual(book.getRange('A1:C5'));
    expect(book.getRange('A1:C4')).toEqual([
      [1, 'a', 4],
      [2, null, null],
      [3, 'b', null],
      [null, 'c', null],
    ]);
  });

  it('renames, moves and deletes sheets, references following them', () => {
    const book = prices();
    book.set('Prices!C1', '=INDIRECT("Prices!B1")');
    expect(book.renameSheet('Prices', 'Costs')).toEqual([{ cell: 'Costs!C1', value: { error: '#REF!' } }]);
    expect(book.sheets).toEqual(['Costs', 'Q1 Sales']);
    expect(book.formula("'Q1 Sales'!C1")).toBe('=Costs!B3*2');
    expect(book.evaluate('=Rate')).toBe(4.5);
    expect(book.renameSheet('costs', 'COSTS')).toEqual([]);
    expect(() => book.renameSheet('COSTS', 'q1 sales')).toThrow('already a sheet named');
    expect(book.moveSheet('Q1 Sales', 0)).toEqual([]);
    expect(book.sheets).toEqual(['Q1 Sales', 'COSTS']);
    expect(book.get('A1')).toBe(1);
    expect(() => book.moveSheet('COSTS', 2)).toThrow('numbered from 0 to 1');
    expect(book.deleteSheet('COSTS')).toEqual([{ cell: "'Q1 Sales'!C1", value: { error: '#REF!' } }]);
    expect(book.formula('C1')).toBe('=#REF!*2');
    expect(() => book.deleteSheet('Q1 Sales')).toThrow("only sheet");
  });

  it('works out references across sheets by where the sheets are', () => {
    const book = createWorkbook({ sheets: { Jan: [[1]], Feb: [[2]], Mar: [[4]], Summary: [['=SUM(Jan:Feb!A1)']] } });
    expect(book.get('Summary!A1')).toBe(3);
    expect(book.moveSheet('Mar', 1)).toEqual([{ cell: 'Summary!A1', value: 7 }]);
    expect(book.deleteSheet('Jan')).toEqual([{ cell: 'Summary!A1', value: 6 }]);
    expect(book.formula('Summary!A1')).toBe('=SUM(Mar:Feb!A1)');
  });
});
