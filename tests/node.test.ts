import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { run, type CliIO } from '../src/node/cli';
import { schemaFromJSON, schemaToJSON, ServerWorkbook, validateWorkbook } from '../src/node/index';
import { forecastSchema, priceListSchema } from '../src/demo/schemas';
import { sample } from './helpers';

describe('ServerWorkbook', () => {
  it('checks a workbook, async checks included', async () => {
    const wb = ServerWorkbook.open(sample('price-list.xlsx'), { fileName: 'price-list.xlsx', schemas: [forecastSchema, priceListSchema] });
    expect(wb.schema?.id).toBe('demo.price-list');
    const issues = await wb.issues();
    expect(issues).toContainEqual({ severity: 'warning', message: 'The catalogue lists this product as discontinued.', sheet: 'Prices', ref: 'A14' });
    expect(issues.filter((i) => i.severity === 'error').map((i) => i.ref)).toEqual(['A8', 'B11', 'C12', 'D13']);
  }, 10_000);

  it('types values in as a user would, running business logic', async () => {
    const wb = ServerWorkbook.open(sample('price-list.xlsx'), { schemas: [priceListSchema] });
    const r = await wb.set('Prices', 'A16', 'SKU-2001');
    expect(r.rejected).toBeUndefined();
    expect(wb.cell('Prices', 'B16')).toBe('Barcode scanner');
    expect(wb.cell('Prices', 'C16')).toBe('Hardware');
    expect(wb.values()!.Prices.tables.prices.at(-1)).toMatchObject({ sku: 'SKU-2001', description: 'Barcode scanner', category: 'Hardware' });
    // As in the editor, a wrong value goes in and is flagged rather than refused.
    await wb.set('Prices', 'D2', 'cheap');
    await wb.set('Prices', 'E2', '1');
    const issues = await wb.issues();
    expect(issues.filter((i) => i.ref === 'D2' || i.ref === 'E2').map((i) => [i.ref, i.message])).toEqual([
      ['D2', 'Expected a number.'],
      ['E2', 'Should be calculated by a formula.'],
    ]);
  }, 10_000);

  it('reports why a workbook matches no schema', async () => {
    const wb = ServerWorkbook.open(sample('budget.xlsx'), { schemas: [priceListSchema] });
    expect(wb.schema).toBeUndefined();
    expect(wb.mismatch?.problems[0]).toMatch(/Prices/);
    const r = await validateWorkbook(sample('budget.xlsx'), { schemas: [priceListSchema] });
    expect(r.ok).toBe(false);
  });

  it('saves with the schema version stamped', async () => {
    // The JSON form, without the code rule that the sample's quarters break on purpose.
    const schema = { ...schemaFromJSON(schemaToJSON(forecastSchema)), version: '7' };
    const wb = ServerWorkbook.open(sample('forecast.xlsx'), { schemas: [schema] });
    await wb.set('Forecast', 'E3', '5');
    const again = ServerWorkbook.open(wb.save(), { schemas: [{ ...schema, version: '8' }] });
    expect(again.note).toMatch(/version 7 .* version 8/);
    const r = await validateWorkbook(wb.save(), { schemas: [schema] });
    expect(r).toMatchObject({ ok: true, errors: 0, schema: { id: schema.id, version: '7' } });
  });
});

describe('workbook-check', () => {
  const io = (files: Record<string, Uint8Array>) => {
    const out: string[] = [];
    const err: string[] = [];
    const written: Record<string, Uint8Array> = {};
    const cli: CliIO = {
      readFile: async (p) => {
        if (!files[p]) throw new Error(`ENOENT: ${p}`);
        return files[p];
      },
      writeFile: async (p, d) => void (written[p] = d),
      out: (t) => out.push(t),
      err: (t) => err.push(t),
    };
    return { cli, out, err, written };
  };
  const schemaFile = (name: string) => new Uint8Array(readFileSync(`examples/schemas/${name}.schema.json`));

  it('prints issues and exits 1 when there are errors', async () => {
    const t = io({ 'f.xlsx': sample('forecast.xlsx'), 's.json': schemaFile('forecast') });
    expect(await run(['f.xlsx', '--schema', 's.json'], t.cli)).toBe(1);
    expect(t.out.join('\n')).toMatch(/matches “Quarterly forecast”[\s\S]*Forecast!E3  Must be at least 0/);
  });

  it('applies --set, saves with --out, and exits 0 once clean', async () => {
    const t = io({ 'f.xlsx': sample('forecast.xlsx'), 's.json': schemaFile('forecast') });
    expect(await run(['f.xlsx', '--schema', 's.json', '--set', 'Forecast!E3=5', '--out', 'o.xlsx', '--json', '--values'], t.cli)).toBe(0);
    const result = JSON.parse(t.out[0]);
    expect(result).toMatchObject({ ok: true, errors: 0, saved: 'o.xlsx', changes: [{ sheet: 'Forecast', ref: 'E3', before: '-5', after: '5' }] });
    expect(result.values.Forecast.tables).toBeDefined();
    expect(t.written['o.xlsx'].length).toBeGreaterThan(1000);
  });

  it('exits 2 for usage and unreadable files', async () => {
    expect(await run([], io({}).cli)).toBe(2);
    expect(await run(['--bogus'], io({}).cli)).toBe(2);
    const t = io({ 'x.xlsx': new TextEncoder().encode('not a workbook') });
    expect(await run(['x.xlsx'], t.cli)).toBe(2);
    expect(t.err[0]).toMatch(/not an Excel workbook/);
  });
});
