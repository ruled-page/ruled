import { it } from 'vitest';
import { engineFormula } from '../src/core/formula';
import { nativeEngine } from '../src/engine/native';

/**
 * Lookup-heavy workbooks, larger and larger, to find where the engine's time
 * stops growing with the work: `npm run lookup:bench`. A Data sheet holds a
 * table of n rows (an id, a name, a price and one of 20 categories) and a
 * Calc sheet n formulas looking it up, one kind of lookup at a time. For each
 * size: building the engine and calculating, editing one formula's key, and
 * editing a price every formula reads.
 */

const SIZES = [1_000, 2_000, 5_000, 10_000, 20_000];
/** A kind of lookup stops growing once a step takes longer than this. */
const LIMIT_MS = 15_000;

const KINDS: Record<string, (n: number, r: number) => string> = {
  'VLOOKUP exact': (n, r) => `VLOOKUP(A${r},Data!$A$1:$D$${n},3,FALSE)`,
  'VLOOKUP exact, whole columns': (_n, r) => `VLOOKUP(A${r},Data!$A:$D,3,FALSE)`,
  'VLOOKUP sorted': (n, r) => `VLOOKUP(A${r},Data!$A$1:$D$${n},3,TRUE)`,
  XLOOKUP: (n, r) => `XLOOKUP(A${r},Data!$A$1:$A$${n},Data!$C$1:$C$${n})`,
  'INDEX/MATCH': (n, r) => `INDEX(Data!$C$1:$C$${n},MATCH(A${r},Data!$A$1:$A$${n},0))`,
  SUMIFS: (n, r) => `SUMIFS(Data!$C$1:$C$${n},Data!$D$1:$D$${n},"Cat "&MOD(A${r},20))`,
};

const time = (f: () => unknown) => {
  const t = performance.now();
  f();
  return performance.now() - t;
};

it('measures lookups as workbooks grow', { timeout: 1_800_000 }, () => {
  const rows: string[] = [];
  for (const [kind, formula] of Object.entries(KINDS)) {
    for (const n of SIZES) {
      // Ids in order, so the sorted lookup is valid; the keys looked up in a scattered order.
      const data = Array.from({ length: n }, (_, i) => [i + 1, `Item ${i + 1}`, ((i * 37) % 1000) + 0.5, `Cat ${i % 20}`]);
      const calc = Array.from({ length: n }, (_, i) => [((i * 7919) % n) + 1, engineFormula(formula(n, i + 1))]);
      let engine!: ReturnType<typeof nativeEngine.create>;
      const build = time(() => {
        engine = nativeEngine.create([{ name: 'Calc', rows: calc }, { name: 'Data', rows: data }], [], false);
      });
      const key = time(() => engine.write([{ sheet: 0, row: 0, col: 0, raw: n }]));
      const price = time(() => engine.write([{ sheet: 1, row: n - 1, col: 2, raw: 1234.5 }]));
      engine.destroy();
      const ms = (t: number) => `${t.toFixed(0)} ms`.padStart(9);
      rows.push(`${kind.padEnd(30)} ${String(n).padStart(6)}  build ${ms(build)}  key edit ${ms(key)}  table edit ${ms(price)}`);
      console.log(rows[rows.length - 1]);
      if (Math.max(build, price) > LIMIT_MS) break;
    }
  }
  console.log(`\n${rows.join('\n')}`);
});
