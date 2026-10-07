import { describe, expect, it } from 'vitest';
import { verifyWorkbook } from '../src/node/verify';
import { runVerify } from '../src/node/verifyCli';
import { sample } from './helpers';

describe('round-trip verification', () => {
  it('passes every check on the kitchen sink, with a row inserted too', () => {
    const r = verifyWorkbook(sample('kitchen-sink.xlsx'), 'kitchen-sink.xlsx', { structure: true, python: false });
    const failed = r.checks.filter((c) => c.status === 'fail');
    expect(failed).toEqual([]);
    expect(r.checks.map((c) => c.name)).toContain('Charts, pivots, pictures and other parts are kept');
    expect(r.probes.length).toBeGreaterThan(5);
    expect(r.outputs.edited?.length).toBeGreaterThan(0);
    expect(r.outputs.restructured?.length).toBeGreaterThan(0);
    expect(r.calc!.formulas).toBeGreaterThan(200);
  });

  it('fails a file that is not a workbook', () => {
    const r = verifyWorkbook(new TextEncoder().encode('nope'), 'x.xlsx', { python: false });
    expect(r.ok).toBe(false);
    expect(r.checks[0]).toMatchObject({ name: 'Opens', status: 'fail' });
  });

  it('writes edited copies and a report', async () => {
    const written: Record<string, Uint8Array | string> = {};
    const out: string[] = [];
    const code = await runVerify(['samples', '--out', 'out', '--no-python'], {
      list: async () => ['samples/orders.xlsx', 'samples/pivot.xlsx'],
      readFile: async (p) => sample(p.split('/').pop()!),
      writeFile: async (p, d) => void (written[p] = d),
      out: (t) => out.push(t),
      err: () => {},
    });
    expect(code).toBe(0);
    expect(Object.keys(written).sort()).toEqual(['out/orders.edited.xlsx', 'out/pivot.edited.xlsx', 'out/report.json', 'out/report.md']);
    expect(String(written['out/report.md'])).toMatch(/The step that needs Excel/);
  });
});
