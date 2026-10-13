import { beforeAll, describe, expect, it } from 'vitest';
import { MAX_COLS, MAX_ROWS } from '@ruled/formula/refs';
import {
  argumentAt,
  completionAt,
  cycleReference,
  formulaReferences,
  insertFunction,
  loadFunctionHelp,
  REF_COLORS,
  signature,
  toggleAbsolute,
} from '../src/core/formulaAssist';

beforeAll(() => loadFunctionHelp());
import { functionHelp } from '../src/core/functionCatalog';

/** The text with "|" marking the caret, split into the text and the caret's position. */
function at(marked: string): [string, number] {
  return [marked.replace('|', ''), marked.indexOf('|')];
}

const names = (marked: string) => completionAt(...at(marked))?.matches.map((f) => f.name);

/** The signature as text, the current argument in **. */
function shown(name: string, index: number): string {
  return signature(functionHelp(name)!, index)
    .map((p) => (p.current ? `**${p.text}**` : p.text))
    .join('');
}

describe('function names as they are typed', () => {
  it('are offered after "=", an operator, a comma, a bracket or a space', () => {
    expect(names('=VL|')).toEqual(['VLOOKUP']);
    expect(names('=sumi|')).toEqual(['SUMIF', 'SUMIFS']);
    for (const f of ['=1+VL|', '=IF(A1,VL|', '=(VL|', '=A1 & VL|', '=-VL|', '={1,VL|', '=A1>=VL|']) expect(names(f), f).toEqual(['VLOOKUP']);
    expect(completionAt(...at('=IF(A1,VL|'))).toMatchObject({ start: 7, end: 9 });
  });

  it('are not offered elsewhere', () => {
    for (const f of [
      'VL|', // not a formula
      '=V|L', // the caret is inside the word
      '="VL|', // in text
      "='VL|", // in a sheet name
      '=Table1[VL|', // in a table reference
      '=Sheet1!VL|', // a name on a sheet
      '=A1:VL|', // part of a range
      '=$VL|',
      '=1VL|',
      '=A1|', // no function starts so
      '=|',
      '=SUM(A1)|',
    ]) {
      expect(names(f), f).toBeUndefined();
    }
    // Text that is closed again doesn't count.
    expect(names('="a ""b"" c"&VL|')).toEqual(['VLOOKUP']);
    expect(names("='My Sheet'!A1+VL|")).toEqual(['VLOOKUP']);
  });

  it('are inserted with an opening bracket, the caret inside it', () => {
    const [text, caret] = at('=1+vl|*2');
    const c = completionAt(text, caret)!;
    expect(insertFunction(text, c, 'VLOOKUP')).toEqual({ text: '=1+VLOOKUP(*2', caret: 11 });
    // A bracket already there is used.
    expect(insertFunction('=SU(A1)', { start: 1, end: 3 }, 'SUM')).toEqual({ text: '=SUM(A1)', caret: 5 });
  });
});

describe('argument hints', () => {
  const arg = (marked: string) => {
    const a = argumentAt(...at(marked));
    return a && [a.fn.name, a.index];
  };

  it('follow the caret through the arguments of the innermost function', () => {
    expect(arg('=VLOOKUP(|')).toEqual(['VLOOKUP', 0]);
    expect(arg('=VLOOKUP(A1, |')).toEqual(['VLOOKUP', 1]);
    expect(arg('=VLOOKUP(A1,B:C,|2)')).toEqual(['VLOOKUP', 2]);
    expect(arg('=IF(SUM(A1,|B1)>2,1)')).toEqual(['SUM', 1]);
    expect(arg('=IF(SUM(A1,B1)>2,|1)')).toEqual(['IF', 1]);
    expect(arg('=_xlfn.XLOOKUP(A1,|')).toEqual(['XLOOKUP', 1]);
    expect(argumentAt(...at('=1+IF(SUM(A1,|B1)>2,1)'))?.start).toBe(6);
    expect(argumentAt(...at('=1+IF((A1|'))?.start).toBe(3);
  });

  it("aren't moved on by commas in text, sheet names, tables or array constants, and plain brackets keep their function's", () => {
    expect(arg('=IF(A1=",,",|')).toEqual(['IF', 1]);
    expect(arg("=SUM('a,b'!A1,|")).toEqual(['SUM', 1]);
    expect(arg('=SUM(T[[#This Row],[a]],|')).toEqual(['SUM', 1]);
    expect(arg('=INDEX({1,2,3},|')).toEqual(['INDEX', 1]);
    expect(arg('=ROUND((1+2)*3,|')).toEqual(['ROUND', 1]);
    expect(arg('=ROUND((1,|')).toEqual(['ROUND', 0]);
  });

  it('are only for functions Excel has, and only inside their brackets', () => {
    expect(arg('=SUM(A1)|')).toBeUndefined();
    expect(arg('=SUM(MYFUNC(|')).toBeUndefined();
    expect(arg('=(1+|')).toBeUndefined();
    expect(arg('SUM(|')).toBeUndefined();
  });

  it("show the signature as Excel does, the argument being typed marked", () => {
    expect(shown('VLOOKUP', 1)).toBe('VLOOKUP(lookup_value, **table_array**, col_index_num, [range_lookup])');
    expect(shown('VLOOKUP', 3)).toBe('VLOOKUP(lookup_value, table_array, col_index_num, [**range_lookup**])');
    expect(shown('VLOOKUP', 4)).toBe('VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])');
    expect(shown('PI', 0)).toBe('PI()');
    expect(shown('BESSELJ', 0)).toBe('BESSELJ');
  });

  it('number repeating arguments on, up to the one being typed', () => {
    expect(shown('SUM', 0)).toBe('SUM(**number1**, [number2], ...)');
    expect(shown('SUM', 1)).toBe('SUM(number1, [**number2**], ...)');
    expect(shown('SUM', 3)).toBe('SUM(number1, [number2], [number3], [**number4**], ...)');
    expect(shown('SUMIFS', 4)).toBe('SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, **criteria2**], ...)');
    expect(shown('SUMIFS', 5)).toBe(
      'SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], [**criteria_range3**, criteria3], ...)',
    );
    expect(shown('LAMBDA', 1)).toBe('LAMBDA([parameter_or_calculation], [**parameter_or_calculation**], ...)');
  });
});

describe('references while a formula is edited', () => {
  it('are found with their cells, the same reference in the same colour', () => {
    const text = "=SUM(A1:B2,Sheet2!$C$3,'My Sheet'!D4)+a1:b2+A:B+$2:$3";
    const refs = formulaReferences(text);
    expect(refs.map((r) => text.slice(r.start, r.end))).toEqual(['A1:B2', 'Sheet2!$C$3', "'My Sheet'!D4", 'a1:b2', 'A:B', '$2:$3']);
    expect(refs.map((r) => r.color)).toEqual([0, 1, 2, 0, 3, 4]);
    expect(refs[0].rect).toEqual({ r1: 0, c1: 0, r2: 1, c2: 1 });
    expect(refs[1]).toMatchObject({ sheet: 'Sheet2', rect: { r1: 2, c1: 2, r2: 2, c2: 2 } });
    expect(refs[2].sheet).toBe('My Sheet');
    expect(refs[4].rect).toEqual({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: 1 });
    expect(refs[5].rect).toEqual({ r1: 1, c1: 0, r2: 2, c2: MAX_COLS - 1 });
  });

  it('take the colours in turn', () => {
    const refs = formulaReferences(`=${Array.from({ length: REF_COLORS + 1 }, (_, i) => `A${i + 1}`).join('+')}`);
    expect(refs.map((r) => r.color)).toEqual([0, 1, 2, 3, 4, 5, 6, 0]);
  });

  it("have no cells when they are in another workbook or span sheets, and aren't found outside formulas", () => {
    const refs = formulaReferences('=[1]Sheet1!A1+Jan:Mar!B2+Name+"A1"');
    expect(refs).toHaveLength(2);
    expect(refs.every((r) => !r.rect)).toBe(true);
    expect(formulaReferences('A1+B1')).toEqual([]);
  });
});

describe('F4', () => {
  it("moves a reference's $ signs on as Excel does", () => {
    const seen = ['A1'];
    for (let i = 0; i < 4; i++) seen.push(cycleReference(seen[i]));
    expect(seen).toEqual(['A1', '$A$1', 'A$1', '$A1', 'A1']);
    expect(cycleReference('a1:b2')).toBe('$a$1:$b$2');
    expect(cycleReference('$A$1:B2')).toBe('A$1:B$2');
    expect(cycleReference("'Q1 Sales'!B$3")).toBe("'Q1 Sales'!$B3");
    expect(cycleReference('A:C')).toBe('$A:$C');
    expect(cycleReference('$A:$C')).toBe('A:C');
    expect(cycleReference('2:5')).toBe('$2:$5');
    expect(cycleReference('$2:$5')).toBe('2:5');
    expect(cycleReference('Name')).toBe('Name');
  });

  it('works on the reference the caret is in or just after, and leaves the caret after it', () => {
    expect(toggleAbsolute('=A1+B2', 2, 2)).toEqual({ text: '=$A$1+B2', start: 5, end: 5 });
    expect(toggleAbsolute('=A1+B2', 3, 3)).toEqual({ text: '=$A$1+B2', start: 5, end: 5 });
    expect(toggleAbsolute('=A1+B2', 6, 6)).toEqual({ text: '=A1+$B$2', start: 8, end: 8 });
    expect(toggleAbsolute('=SUM(Data!A1:A9)', 8, 8)).toEqual({ text: '=SUM(Data!$A$1:$A$9)', start: 19, end: 19 });
    expect(toggleAbsolute('=A1+1', 5, 5)).toBeUndefined();
    expect(toggleAbsolute('A1', 1, 1)).toBeUndefined();
  });

  it('works on every reference a selection touches, and keeps them selected', () => {
    expect(toggleAbsolute('=A1+B2*C3', 0, 6)).toEqual({ text: '=$A$1+$B$2*C3', start: 0, end: 10 });
    expect(toggleAbsolute('=A1+B2*C3', 4, 9)).toEqual({ text: '=A1+$B$2*$C$3', start: 4, end: 13 });
  });
});
