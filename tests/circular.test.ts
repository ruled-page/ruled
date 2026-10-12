import { readFileSync } from 'node:fs';
import { strToU8, zipSync } from 'fflate';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { engineFormula } from '../src/core/formula';
import { WorkbookSession } from '../src/core/session';
import { nativeEngine } from '../src/engine/native';
import { english } from '../src/i18n';

describe('circular references', () => {
  it('read as 0, and the engine lists where each closes until it is broken', () => {
    const e = nativeEngine.create([{ name: 'S', rows: [[engineFormula('B1+1'), engineFormula('A1*2'), engineFormula('C1')], [5]] }], [], false);
    // The circle closes at A1, which B1 read as 0.
    expect([e.value(0, 0, 0), e.value(0, 0, 1)]).toEqual([0, 0]);
    expect(e.circular()).toEqual([
      { sheet: 0, row: 0, col: 0 },
      { sheet: 0, row: 0, col: 2 },
    ]);
    e.write([{ sheet: 0, row: 0, col: 1, raw: engineFormula('A2*2') }]);
    expect(e.value(0, 0, 0)).toBe(11);
    expect(e.circular()).toEqual([{ sheet: 0, row: 0, col: 2 }]);
    e.write([{ sheet: 0, row: 0, col: 2, raw: 3 }]);
    expect(e.circular()).toEqual([]);
    e.destroy();
  });

  it('are warned of when an edit makes one, as Excel warns, and listed for the status bar', () => {
    const s = WorkbookSession.open(new Uint8Array(readFileSync(join(__dirname, 'corpus', 'templates.xlsx'))), 'templates.xlsx', 'en-US', { recalc: 'on' });
    expect(s.ensureEngine().circular).toEqual([]);
    const sheet = s.model.sheets.findIndex((m) => m.info.name === 'Arrays');
    const first = s.apply({ type: 'input', sheet, row: 200, col: 0, text: '=A202' });
    expect(first.warnings).toEqual([]);
    const closed = s.apply({ type: 'input', sheet, row: 201, col: 0, text: '=A201+1' });
    expect(closed.warnings.map((w) => english.text(w))).toEqual([expect.stringMatching(/^There are circular references/)]);
    expect(closed.circular).toEqual([{ sheet, row: 201, col: 0 }]);
    // Another edit that leaves the circle as it is doesn't warn again.
    expect(s.apply({ type: 'input', sheet, row: 205, col: 0, text: '7' }).warnings).toEqual([]);
    const broken = s.apply({ type: 'input', sheet, row: 201, col: 0, text: '4' });
    expect(broken.circular).toEqual([]);
  });
});

describe('iterative calculation', () => {
  const ITERATE = { count: 100, delta: 0.001 };
  const engine = (rows: (string | number)[][], iteration = ITERATE) =>
    nativeEngine.create([{ name: 'S', rows: rows.map((r) => r.map((v) => (typeof v === 'string' ? engineFormula(v) : v))) }], [], false, undefined, iteration);

  it('works a circle round until no value moves by more than the maximum change', () => {
    // Interest on a balance that includes the interest.
    const e = engine([['B1*0.1', '100+A1']]);
    expect(e.value(0, 0, 1) as number).toBeCloseTo(1000 / 9, 3);
    expect(e.value(0, 0, 0) as number).toBeCloseTo(100 / 9, 3);
    const changed = e.write([{ sheet: 0, row: 0, col: 1, raw: engineFormula('200+A1') }]);
    expect(changed.map((c) => c.col).sort()).toEqual([0, 1]);
    expect(e.value(0, 0, 1) as number).toBeCloseTo(2000 / 9, 3);
    e.destroy();
  });

  it('stops at the first pass that moves nothing by more than the maximum change, or when the passes run out', () => {
    // Halving the distance to 2 each pass: 1, 1.5, 1.75 and so on.
    expect(engine([['A1/2+1']]).value(0, 0, 0)).toBe(2 - 2 ** -10);
    expect(engine([['A1/2+1']], { count: 3, delta: 0.001 }).value(0, 0, 0)).toBe(1.75);
    // A formula counting itself up counts once a pass.
    expect(engine([['A1+1']]).value(0, 0, 0)).toBe(100);
  });

  it('is read from the workbook, and a circle that iterates is neither warned of nor named', () => {
    const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const rels = (list: [string, string, string][]) =>
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list
        .map(([id, type, target]) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`)
        .join('')}</Relationships>`;
    const file = (calcPr: string, [a, b]: number[]) =>
      zipSync({
        '[Content_Types].xml': strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'),
        '_rels/.rels': strToU8(rels([['rId1', 'officeDocument', 'xl/workbook.xml']])),
        'xl/workbook.xml': strToU8(`<workbook xmlns="${MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>${calcPr}</workbook>`),
        'xl/_rels/workbook.xml.rels': strToU8(rels([['rId1', 'worksheet', 'worksheets/sheet1.xml']])),
        'xl/worksheets/sheet1.xml': strToU8(`<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1"><f>B1*0.1</f><v>${a}</v></c><c r="B1"><f>100+A1</f><v>${b}</v></c></row></sheetData></worksheet>`),
      });
    // Values Excel saved once the circle moved by less than the maximum change.
    const s = WorkbookSession.open(file('<calcPr iterate="1" iterateCount="50" iterateDelta="0.0001"/>', [11.11111, 111.1111]), 'i.xlsx', 'en-US', { recalc: 'on' });
    expect(s.model.iteration).toEqual({ count: 50, delta: 0.0001 });
    const report = s.ensureEngine();
    expect(report.circular).toEqual([]);
    expect(s.model.findings.map((f) => f.kind)).not.toContain('formulaEngine');
    expect(s.apply({ type: 'input', sheet: 0, row: 1, col: 0, text: '=A2+1' }).warnings).toEqual([]);
    // Without iteration the same circle reads 0, and is named.
    const plain = WorkbookSession.open(file('<calcPr/>', [0, 100]), 'p.xlsx', 'en-US', { recalc: 'on' });
    expect(plain.model.iteration).toBeUndefined();
    expect(plain.ensureEngine().circular).toEqual([{ sheet: 0, row: 0, col: 0 }]);
  });
});
