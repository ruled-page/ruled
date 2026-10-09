import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { engineFormula, toFileFormula } from '../src/core/formula';
import { nativeEngine } from '../src/engine/native';
import { FAMILIES } from './excel-checks/cases';

/**
 * Functions the native engine was checked against Excel on: answers Excel
 * agreed with when it graded the check workbook (npm run excel:check), held
 * here so they stay as they are. Excel's searched results (IRR, XIRR, RATE,
 * YIELD) agreed within its documented accuracy; these are the exact ones.
 */
describe('functions checked against Excel', () => {
  for (const [family, formulas] of Object.entries(FAMILIES)) {
    const file = join(__dirname, 'excel-checks', `${family}.json`);
    it.skipIf(!existsSync(file))(`${family}: gives the answers Excel agreed with`, () => {
      const expected = JSON.parse(readFileSync(file, 'utf8')) as Record<string, unknown>;
      const engine = nativeEngine.create([{ name: 'Check', rows: formulas.map((f) => [engineFormula(toFileFormula(f).slice(1))]) }], [], false);
      const got = Object.fromEntries(formulas.map((f, i) => [f, engine.value(0, i, 0)]));
      engine.destroy();
      expect(got).toEqual(expected);
    });
  }
});
