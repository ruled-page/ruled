import ExcelJS from 'exceljs';
import { beforeAll, describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import type { StructuralOp } from '../src/core/refTransform';
import { english, errorText } from '../src/i18n';
import { loadedFormula, loadedValue, parts, sample, sheetIndex } from './helpers';

/** What a call throws, in English. */
function thrown(fn: () => unknown): string | undefined {
  try {
    fn();
  } catch (e) {
    return english.text(errorText(e));
  }
  return undefined;
}

/**
 * A workbook with something of every kind that refers to cells:
 *   Data!A1:A5 = 1, 2, 3, "4", 5, B1 = SUM(A1:A5), C1 = A3*2, D1:E2 merged,
 *   validation and conditional formatting on A1:A5, a note on A5, a link on A4,
 *   names Nums = Data!$A$1:$A$5 and Third = Data!$A$3,
 *   Calc!A1 = Data!A5, A2 = SUM(Data!A1:A5), A3 = Data!A3, A4 = Data!A1.
 */
async function build(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet('Data');
  const calc = wb.addWorksheet('Calc');
  for (let i = 1; i <= 5; i++) data.getCell(`A${i}`).value = i;
  data.getCell('B1').value = { formula: 'SUM(A1:A5)', result: 11 };
  data.getCell('C1').value = { formula: 'A3*2', result: 6 };
  data.getCell('D1').value = 'Merged';
  data.mergeCells('D1:E2');
  data.getCell('A4').value = { text: '4', hyperlink: 'https://example.com' };
  data.getCell('A5').note = 'Five';
  // ExcelJS supports range validations but its types don't declare them.
  (data as unknown as { dataValidations: { add(ref: string, v: object): void } }).dataValidations.add('A1:A5', {
    type: 'whole',
    operator: 'between',
    formulae: [0, 100],
    showErrorMessage: true,
  });
  data.addConditionalFormatting({
    ref: 'A1:A5',
    rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: ['3'], priority: 1, style: { font: { bold: true } } }],
  });
  calc.getCell('A1').value = { formula: 'Data!A5', result: 5 };
  calc.getCell('A2').value = { formula: 'SUM(Data!A1:A5)', result: 11 };
  calc.getCell('A3').value = { formula: 'Data!A3', result: 3 };
  calc.getCell('A4').value = { formula: 'Data!A1', result: 1 };
  wb.definedNames.add('Data!$A$1:$A$5', 'Nums');
  wb.definedNames.add('Data!$A$3', 'Third');
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await build();
});

/** Applies a change and reopens the saved result, as Excel would see it. */
function after(op: StructuralOp, edit?: (s: WorkbookSession) => void) {
  const s = WorkbookSession.open(bytes, 'refs.xlsx');
  edit?.(s);
  const next = s.restructure(op);
  const saved = next.save();
  return { before: s, next, file: WorkbookSession.open(saved, 'refs.xlsx'), xml: parts(saved) };
}

const DATA = 0;
const CALC = 1;
const meta = (xml: Record<string, string>) => xml['xl/worksheets/sheet1.xml'].slice(xml['xl/worksheets/sheet1.xml'].indexOf('</sheetData>'));

describe('inserting rows', () => {
  it('moves cells down and grows the ranges they were inside', () => {
    const { file, xml } = after({ type: 'insertRows', sheet: DATA, at: 2, count: 2 });
    expect([1, 2, 3, 4, 5, 6, 7].map((r) => loadedValue(file, 'Data', `A${r}`))).toEqual([1, 2, null, null, 3, '4', 5]);
    expect(loadedFormula(file, 'Data', 'B1')).toBe('SUM(A1:A7)');
    expect(loadedFormula(file, 'Data', 'C1')).toBe('A5*2');
    expect(loadedFormula(file, 'Calc', 'A1')).toBe('Data!A7');
    expect(loadedFormula(file, 'Calc', 'A2')).toBe('SUM(Data!A1:A7)');
    expect(loadedFormula(file, 'Calc', 'A3')).toBe('Data!A5');
    expect(xml['xl/workbook.xml']).toContain('>Data!$A$1:$A$7<');
    expect(xml['xl/workbook.xml']).toContain('>Data!$A$5<');
    const tail = meta(xml);
    expect(tail).toContain('ref="D1:E2"');
    expect(tail).toMatch(/<dataValidation [^>]*sqref="A1:A7"/);
    expect(tail).toMatch(/<conditionalFormatting sqref="A1:A7"/);
    expect(tail).toMatch(/<hyperlink ref="A6"/);
    expect(Object.entries(xml).find(([k]) => /comments\d*\.xml$/.test(k))?.[1]).toContain('ref="A7"');
  });

  it('keeps the values the engine calculates', () => {
    const { next } = after({ type: 'insertRows', sheet: DATA, at: 2, count: 2 });
    next.ensureEngine();
    next.apply({ type: 'input', sheet: DATA, row: 2, col: 0, text: '10' });
    const updates = next.stateUpdates();
    const b1 = updates.find((u) => u.sheet === DATA && u.row === 0 && u.col === 1);
    // A4 holds the text "4" (it is a link), which SUM skips.
    expect(b1?.n).toBe(21);
  });

  it('refuses to push data off the sheet', () => {
    const s = WorkbookSession.open(bytes, 'refs.xlsx');
    expect(thrown(() => s.restructure({ type: 'insertRows', sheet: DATA, at: 0, count: 1_048_575 }))).toMatch(/off the bottom/);
  });
});

describe('deleting rows', () => {
  it('pulls cells up, shrinks ranges, and turns references to deleted cells into #REF!', () => {
    const { file, xml } = after({ type: 'deleteRows', sheet: DATA, at: 2, count: 2 });
    expect([1, 2, 3].map((r) => loadedValue(file, 'Data', `A${r}`))).toEqual([1, 2, 5]);
    expect(loadedFormula(file, 'Data', 'B1')).toBe('SUM(A1:A3)');
    expect(loadedFormula(file, 'Data', 'C1')).toBe('#REF!*2');
    expect(loadedFormula(file, 'Calc', 'A1')).toBe('Data!A3');
    expect(loadedFormula(file, 'Calc', 'A3')).toBe('#REF!');
    expect(xml['xl/workbook.xml']).toContain('>Data!$A$1:$A$3<');
    expect(xml['xl/workbook.xml']).toMatch(/name="Third"[^>]*>#REF!</);
    // The link on the deleted row A4 goes with it.
    expect(meta(xml)).not.toContain('<hyperlink');
  });
});

describe('columns', () => {
  it('insert and delete with the same rules, including column widths', () => {
    const ins = after({ type: 'insertCols', sheet: DATA, at: 1, count: 1 });
    expect(loadedFormula(ins.file, 'Data', 'C1')).toBe('SUM(A1:A5)');
    expect(loadedFormula(ins.file, 'Data', 'D1')).toBe('A3*2');
    expect(meta(ins.xml)).toContain('ref="E1:F2"');
    const del = after({ type: 'deleteCols', sheet: DATA, at: 1, count: 1 });
    expect(loadedFormula(del.file, 'Data', 'B1')).toBe('A3*2');
    expect(loadedFormula(del.file, 'Calc', 'A2')).toBe('SUM(Data!A1:A5)');
  });
});

describe('moving cells', () => {
  it('takes references with it, as cut and paste does', () => {
    const { file } = after({ type: 'move', sheet: DATA, src: { r1: 0, c1: 0, r2: 0, c2: 0 }, toSheet: DATA, row: 9, col: 6 });
    expect(loadedValue(file, 'Data', 'G10')).toBe(1);
    expect(loadedValue(file, 'Data', 'A1')).toBeNull();
    expect(loadedFormula(file, 'Calc', 'A4')).toBe('Data!G10');
    // A range only partly moved stays as it was.
    expect(loadedFormula(file, 'Data', 'B1')).toBe('SUM(A1:A5)');
  });

  it('moves formulas without shifting what they point at', () => {
    const { file } = after({ type: 'move', sheet: DATA, src: { r1: 0, c1: 2, r2: 0, c2: 2 }, toSheet: DATA, row: 5, col: 5 });
    expect(loadedFormula(file, 'Data', 'F6')).toBe('A3*2');
  });

  it('qualifies references when cells move to another sheet', () => {
    const { file } = after({ type: 'move', sheet: DATA, src: { r1: 0, c1: 1, r2: 0, c2: 2 }, toSheet: CALC, row: 0, col: 3 });
    expect(loadedFormula(file, 'Calc', 'D1')).toBe('SUM(Data!A1:A5)');
    expect(loadedFormula(file, 'Calc', 'E1')).toBe('Data!A3*2');
  });

  it('breaks references to cells that were pasted over', () => {
    const { file } = after({ type: 'move', sheet: DATA, src: { r1: 0, c1: 3, r2: 1, c2: 4 }, toSheet: DATA, row: 0, col: 0 });
    expect(loadedFormula(file, 'Calc', 'A4')).toBe('#REF!');
  });

  it('refuses to split merged cells', () => {
    const s = WorkbookSession.open(bytes, 'refs.xlsx');
    expect(thrown(() => s.restructure({ type: 'move', sheet: DATA, src: { r1: 0, c1: 3, r2: 0, c2: 3 }, toSheet: DATA, row: 8, col: 8 }))).toMatch(/part of the merged cell/);
  });
});

describe('change tracking across structural changes', () => {
  it('carries earlier edits to where their cells went, and lists the change', () => {
    const { next } = after({ type: 'insertRows', sheet: DATA, at: 0, count: 1 }, (s) => s.apply({ type: 'input', sheet: DATA, row: 1, col: 0, text: '20' }));
    expect(next.changes().map((c) => (c.note ? { ...c, note: english.text(c.note) } : c))).toEqual([
      { sheet: 'Data', ref: '1:1', before: '', after: '', note: 'Inserted 1 row' },
      { sheet: 'Data', ref: 'A3', before: '2', after: '20' },
    ]);
    expect(next.editedCount).toBe(2);
    next.apply({ type: 'input', sheet: DATA, row: 2, col: 0, text: '2' });
    expect(next.changes().filter((c) => !c.note)).toEqual([]);
  });
});

describe('real files', () => {
  it('inserts rows into the kitchen sink without losing anything', () => {
    const s = WorkbookSession.open(sample('kitchen-sink.xlsx'), 'k.xlsx');
    const data = sheetIndex(s, 'Data');
    const next = s.restructure({ type: 'insertRows', sheet: data, at: 5, count: 3 });
    const saved = next.save();
    const xml = parts(saved);
    // The table and the summary formulas that read it grow.
    expect(Object.values(xml).some((x) => /<table [^>]*ref="A1:F204"/.test(x))).toBe(true);
    const file = WorkbookSession.open(saved, 'k.xlsx');
    expect(file.model.sheets[data].info.tables[0].rect).toEqual({ r1: 0, c1: 0, r2: 203, c2: 5 });
    // Other sheets' formulas move exactly once.
    expect(loadedFormula(file, 'Summary', 'B5')).toBe(
      'SUMIFS(Data!$F$2:$F$204,Data!$B$2:$B$204,$A5,Data!$A$2:$A$204,">="&DATE(2026,1,1),Data!$A$2:$A$204,"<"&DATE(2026,4,1))',
    );
    expect(file.loadResult().findings.length).toBeGreaterThan(0);
  });
});

describe('sorting', () => {
  it("orders rows Excel's way and moves formulas as if copied", async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.addRow(['Name', 'Score', 'Double']);
    const rows: [string | null, number | string | null][] = [['b', 3], ['A', 1], [null, null], ['c', 'x'], ['a', 2]];
    rows.forEach(([n, v], i) => {
      ws.getCell(`A${i + 2}`).value = n;
      ws.getCell(`B${i + 2}`).value = v;
      ws.getCell(`C${i + 2}`).value = { formula: `B${i + 2}*2`, result: typeof v === 'number' ? v * 2 : 0 };
    });
    const s = WorkbookSession.open(new Uint8Array(await wb.xlsx.writeBuffer()), 's.xlsx');
    const rect = { r1: 0, c1: 0, r2: 5, c2: 2 };
    s.apply({ type: 'sort', sheet: 0, rect, keys: [{ col: 1, descending: false }], header: true });
    const col = (c: number) => [1, 2, 3, 4, 5].map((r) => s.current(0, r, c).content);
    // Numbers, then text, with the blank last; formulas point at their own row.
    expect(col(1).map((c) => ('value' in c ? c.value : null))).toEqual([1, 2, 3, 'x', null]);
    expect(col(2).map((c) => ('formula' in c ? c.formula : null))).toEqual(['B2*2', 'B3*2', 'B4*2', 'B5*2', 'B6*2']);
    s.apply({ type: 'sort', sheet: 0, rect, keys: [{ col: 0, descending: true }], header: true });
    // Text ignores case; descending still leaves the blank row last.
    expect(col(0).map((c) => ('value' in c ? c.value : null))).toEqual(['c', 'b', 'A', 'a', null]);
  });
});
