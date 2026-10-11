import { describe, expect, it } from 'vitest';
import type { EngineSheet } from '../src/core/engine';
import { engineFormula } from '../src/core/formula';
import { nativeEngine } from '../src/engine/native';

const sheet = (name: string, rows: EngineSheet['rows'] = []): EngineSheet => ({ name, rows });

describe('three-dimensional references and unions in the native engine', () => {
  it('are calculated, and worked out again when a cell they read changes', () => {
    expect(nativeEngine.unsupportedReason('SUM(Jan:Mar!B2)')).toBeUndefined();
    const e = nativeEngine.create(
      [sheet('S', [[engineFormula('SUM(Jan:Mar!A1)'), engineFormula('SUM((Feb!A1:A2,Feb!B1))')]]), sheet('Jan', [[1]]), sheet('Feb', [[2, 5], [3]]), sheet('Mar', [[4]])],
      [],
      false,
    );
    expect([e.value(0, 0, 0), e.value(0, 0, 1)]).toEqual([7, 10]);
    const changes = e.write([{ sheet: 2, row: 0, col: 0, raw: 20 }]);
    expect(changes).toContainEqual({ sheet: 0, row: 0, col: 0, value: 25 });
    expect(changes).toContainEqual({ sheet: 0, row: 0, col: 1, value: 28 });
    e.destroy();
  });

  it('read a sheet that arrives within their run of sheets', () => {
    const e = nativeEngine.create([sheet('S', [[engineFormula('SUM(Jan:Mar!A1)')]]), sheet('Jan', [[1]]), null, sheet('Mar', [[4]])], [], false);
    expect(e.value(0, 0, 0)).toBe(5);
    expect(e.addSheet(2, sheet('Feb', [[2]]))).toEqual([{ sheet: 0, row: 0, col: 0, value: 7 }]);
    e.destroy();
  });
});

describe('a sheet that arrives after the formulas naming it', () => {
  it('is read by them', () => {
    const e = nativeEngine.create([sheet('S', [[engineFormula('Other!A1*2')]]), null], [], false);
    expect(e.value(0, 0, 0)).toEqual({ error: '#REF!' });
    expect(e.addSheet(1, sheet('Other', [[21]]))).toEqual([{ sheet: 0, row: 0, col: 0, value: 42 }]);
    e.destroy();
  });
});
