import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { forecastSchema } from '../src/demo/schemas';
import { run, USAGE, type CliIO } from '../src/node/cli';
import { schemaFromJSON, schemaToJSON, ServerWorkbook } from '../src/node';
import { sample } from './helpers';

const json = (text: string) => new TextEncoder().encode(text);
const example = (name: string) => new Uint8Array(readFileSync(`examples/schemas/${name}.schema.json`));

/** The budget example with a version, and a warning the sample sets off: January's spending is over 50,000. */
function budgetWithWarning(version?: string) {
  const schema = JSON.parse(readFileSync('examples/schemas/budget.schema.json', 'utf8'));
  schema.sheets[0].tables[0].checks = [{ rule: 'sum', field: 'jan', max: 50000, severity: 'warning' }];
  return json(JSON.stringify(version ? { ...schema, version } : schema));
}

/** Runs the command line over files held in memory. */
async function cli(argv: string[], files: Record<string, Uint8Array> = {}) {
  const out: string[] = [];
  const err: string[] = [];
  const written: Record<string, Uint8Array> = {};
  const io: CliIO = {
    readFile: async (p) => {
      if (!files[p]) throw new Error(`ENOENT: no such file, open '${p}'`);
      return files[p];
    },
    writeFile: async (p, d) => void (written[p] = d),
    out: (t) => out.push(t),
    err: (t) => err.push(t),
  };
  const code = await run(argv, io);
  return { code, out: out.join('\n'), err: err.join('\n'), written };
}

describe('workbook-check’s options', () => {
  it('shows the usage for --help and -h', async () => {
    for (const flag of ['--help', '-h']) {
      expect(await cli([flag])).toEqual({ code: 0, out: USAGE, err: '', written: {} });
    }
  });

  it('says what is wrong with the command line, then the usage', async () => {
    const cases: [string[], string][] = [
      [['f.xlsx', '--schema'], '--schema needs a value.'],
      [['f.xlsx', '--out'], '--out needs a value.'],
      [['f.xlsx', '--set', 'A1=3'], '--set takes Sheet!A1=value.'],
      [['f.xlsx', '--set', 'Sheet!A=3'], '--set takes Sheet!A1=value.'],
      [['a.xlsx', 'b.xlsx'], 'Only one workbook at a time (got a.xlsx and b.xlsx).'],
      [['f.xlsx', '--verbose'], 'Unknown option --verbose.'],
    ];
    for (const [argv, message] of cases) {
      const r = await cli(argv);
      expect(r.code, argv.join(' ')).toBe(2);
      expect(r.err).toBe(`${message}\n\n${USAGE}`);
    }
  });

  it('names a schema file it can’t use', async () => {
    const files = { 'f.xlsx': sample('forecast.xlsx'), 'bad.json': json('{ nope'), 'odd.json': json('{"id":"x"}') };
    const bad = await cli(['f.xlsx', '--schema', 'bad.json'], files);
    expect(bad.code).toBe(2);
    expect(bad.err).toMatch(/^bad\.json: The schema file isn't valid JSON: /);
    expect((await cli(['f.xlsx', '--schema', 'odd.json'], files)).err).toMatch(/^odd\.json: /);
    expect(await cli(['f.xlsx', '--schema', 'missing.json'], files)).toMatchObject({ code: 2, err: "ENOENT: no such file, open 'missing.json'" });
    expect(await cli(['gone.xlsx'], files)).toMatchObject({ code: 2, err: "ENOENT: no such file, open 'gone.xlsx'" });
  });
});

describe('workbook-check’s report', () => {
  it('reads a workbook without schemas, naming it without its folder', async () => {
    expect(await cli(['data/in/orders.xlsx'], { 'data/in/orders.xlsx': sample('orders.xlsx') })).toMatchObject({ code: 0, out: 'orders.xlsx: read 2 sheets; no schema given.' });
    expect(await cli(['data\\forecast.xlsx', '--values'], { 'data\\forecast.xlsx': sample('forecast.xlsx') })).toMatchObject({
      code: 0,
      out: 'forecast.xlsx: read 1 sheet; no schema given.\nnull',
    });
  });

  it('names a schema written in several languages in English', async () => {
    const r = await cli(['f.xlsx', '--schema', 's.json'], { 'f.xlsx': sample('forecast.xlsx'), 's.json': json(schemaToJSON(forecastSchema)) });
    expect(r.out.split('\n')[0]).toBe('f.xlsx: matches “Quarterly forecast”.');
    const asJson = await cli(['f.xlsx', '--schema', 's.json', '--json'], { 'f.xlsx': sample('forecast.xlsx'), 's.json': json(schemaToJSON(forecastSchema)) });
    expect(JSON.parse(asJson.out).schema).toEqual({ id: 'demo.forecast', name: 'Quarterly forecast' });
  });

  it('says when a workbook matches no schema, and why', async () => {
    const files = { 'old.xlsx': sample('price-list-old.xlsx'), 'b.xlsx': sample('budget.xlsx'), 'p.json': example('price-list') };
    expect(await cli(['old.xlsx', '--schema', 'p.json'], files)).toMatchObject({
      code: 1,
      out: 'old.xlsx: doesn\'t match “Price list”.\n  Prices!C1 should be "Category" (field label) but is "Group".',
    });
    expect(await cli(['b.xlsx', '--schema', 'p.json'], files)).toMatchObject({
      code: 1,
      out: 'b.xlsx: doesn\'t match any schema given.\n  The workbook has no sheet named "Prices".',
    });
  });

  it('counts errors and warnings, and gives the schema’s version', async () => {
    const r = await cli(['b.xlsx', '--schema', 'b.json'], { 'b.xlsx': sample('budget.xlsx'), 'b.json': budgetWithWarning('2026.1') });
    expect(r.code).toBe(1);
    expect(r.out.split('\n')).toEqual([
      'b.xlsx: matches “Half-year budget” (version 2026.1).',
      '  error    Budget!B8  Required.',
      '  warning  Budget!B1  “Jan” adds up to 60,200, over the limit of 50,000.',
      '1 error, 1 warning.',
    ]);
  });

  it('mentions a file saved under another version of the schema', async () => {
    const v1 = { ...schemaFromJSON(readFileSync('examples/schemas/forecast.schema.json', 'utf8')), version: '1' };
    const wb = ServerWorkbook.open(sample('forecast.xlsx'), { schemas: [v1] });
    await wb.set('Forecast', 'E3', '5');
    const r = await cli(['f.xlsx', '--schema', 's.json'], { 'f.xlsx': wb.save(), 's.json': json(schemaToJSON({ ...v1, version: '2' })) });
    expect(r.code).toBe(0);
    expect(r.out.split('\n')[1]).toMatch(/version 1 .* version 2/);
    expect(r.out.split('\n')[2]).toBe('0 errors, 0 warnings.');
  });
});

describe('workbook-check’s changes', () => {
  it('takes quoted sheet names, absolute references and values with = in them', async () => {
    const r = await cli(['o.xlsx', '--set', "'Orders'!$z$1==1+1", '--set', 'Orders!Z2=a=b', '--json'], { 'o.xlsx': sample('orders.xlsx') });
    expect(r.code).toBe(0);
    expect(JSON.parse(r.out).changes).toEqual([
      { sheet: 'Orders', ref: 'Z1', before: '', after: '=1+1' },
      { sheet: 'Orders', ref: 'Z2', before: '', after: 'a=b' },
    ]);
    const text = await cli(['o.xlsx', '--set', "'Orders'!$z$1==1+1", '--set', 'Orders!Z2=a=b', '--out', 'o2.xlsx'], { 'o.xlsx': sample('orders.xlsx') });
    expect(text.out.split('\n').at(-1)).toBe('Saved o2.xlsx (2 changes).');
    expect(ServerWorkbook.open(text.written['o2.xlsx']).cell('Orders', 'Z1')).toBe(2);
  });

  it('lists the cells it couldn’t change, and exits 1', async () => {
    const files = { 'p.xlsx': sample('price-list.xlsx'), 'p.json': example('price-list') };
    const r = await cli(['p.xlsx', '--schema', 'p.json', '--set', 'Prices!A1=Code', '--set', 'Nope!A1=1', '--set', 'Prices!XFE1=1', '--set', 'Prices!D2==1+1'], files);
    expect(r.code).toBe(1);
    expect(r.out.split('\n').filter((l) => l.startsWith('  not done'))).toEqual([
      '  not done  Prices!A1: Column names in this table are set by the “Price list” schema.',
      '  not done  Nope!A1: There is no sheet called "Nope".',
      '  not done  Prices!XFE1: "XFE1" isn\'t a cell reference.',
      "  not done  Prices!D2: Formulas aren't allowed in “Unit price”. Type a value instead.",
    ]);
  });

  it('doesn’t save over errors the schema blocks', async () => {
    const files = { 'f.xlsx': sample('forecast.xlsx'), 'f.json': example('forecast') };
    const r = await cli(['f.xlsx', '--schema', 'f.json', '--out', 'o.xlsx'], files);
    expect(r.code).toBe(1);
    expect(r.out).toContain('  not done  Not saved: 1 cell breaks the “Quarterly forecast” schema.');
    expect(r.out).not.toContain('Saved o.xlsx');
    expect(r.written).toEqual({});
    const asJson = JSON.parse((await cli(['f.xlsx', '--schema', 'f.json', '--out', 'o.xlsx', '--json'], files)).out);
    expect(asJson.saved).toBeUndefined();
    expect(asJson.rejected).toEqual(['Not saved: 1 cell breaks the “Quarterly forecast” schema.']);
  });

  it('saves over errors when the schema allows it, and says how many changes', async () => {
    const files = { 'b.xlsx': sample('budget.xlsx'), 'b.json': budgetWithWarning() };
    const r = await cli(['b.xlsx', '--schema', 'b.json', '--set', 'Budget!G5=4000', '--out', 'out.xlsx'], files);
    expect(r.code).toBe(1);
    expect(r.out.split('\n').slice(-2)).toEqual(['1 error, 1 warning.', 'Saved out.xlsx (1 change).']);
    const saved = ServerWorkbook.open(r.written['out.xlsx']);
    expect(saved.cell('Budget', 'G5')).toBe(4000);
  });

  it('prints the values after the report', async () => {
    const files = { 'b.xlsx': sample('budget.xlsx'), 'b.json': example('budget') };
    const r = await cli(['b.xlsx', '--schema', 'b.json', '--set', 'Budget!B8=Dana', '--out', 'o.xlsx', '--values'], files);
    expect(r.code).toBe(0);
    const [report, values] = [r.out.slice(0, r.out.indexOf('{')), JSON.parse(r.out.slice(r.out.indexOf('{')))];
    expect(report.trim().split('\n')).toEqual(['b.xlsx: matches “Half-year budget”.', '0 errors, 0 warnings.', 'Saved o.xlsx (1 change).']);
    expect(values.Budget.cells).toEqual({ approvedBy: 'Dana', total: 389650.5 });
  });
});
