import { readFileSync } from 'node:fs';
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
