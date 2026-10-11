import { describe, expect, it } from 'vitest';
import type { EngineSheet } from '../src/core/engine';
import { engineFormula } from '../src/core/formula';
import { nativeEngine } from '../src/engine/native';

const sheet = (name: string, rows: EngineSheet['rows'] = []): EngineSheet => ({ name, rows });

describe('LAMBDA in the native engine', () => {
  it('calls names defined as LAMBDAs, which the workbook tells it of', () => {
    expect(nativeEngine.unsupportedReason('TWICE(2)')).toEqual({ kind: 'function', name: 'TWICE' });
    expect(nativeEngine.unsupportedReason('TWICE(2)', new Set(['TWICE']))).toBeUndefined();
    expect(nativeEngine.unsupportedReason('_xlfn.LAMBDA(_xlpm.f,_xlpm.f(2))(_xlfn.LAMBDA(_xlpm.x,_xlpm.x))')).toBeUndefined();
  });

  it('works a name out again when it is defined afresh, and stops a LAMBDA that calls itself for ever', () => {
    const e = nativeEngine.create(
      [sheet('S', [[engineFormula('TWICE(A2)'), engineFormula('FOREVER(1)')], [5]])],
      [
        { name: 'TWICE', formula: '_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)' },
        { name: 'FOREVER', formula: '_xlfn.LAMBDA(_xlpm.n,FOREVER(_xlpm.n+1))' },
      ],
      false,
    );
    expect([e.value(0, 0, 0), e.value(0, 0, 1)]).toEqual([10, { error: '#NUM!' }]);
    expect(e.write([{ sheet: 0, row: 1, col: 0, raw: 6 }])).toContainEqual({ sheet: 0, row: 0, col: 0, value: 12 });
    expect(e.setName('TWICE', undefined, '_xlfn.LAMBDA(_xlpm.x,_xlpm.x*3)').changes).toContainEqual({ sheet: 0, row: 0, col: 0, value: 18 });
    e.destroy();
  });
});
