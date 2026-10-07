import ExcelJS from 'exceljs';
import { beforeAll, describe, expect, it } from 'vitest';
import { parseA1 } from '../src/core/refs';
import { WorkbookSession } from '../src/core/session';
import { bindSchema } from '../src/schema/binding';
import { SchemaController } from '../src/schema/controller';
import { generateSchema } from '../src/schema/generate';
import { schemaFromJSON } from '../src/schema/json';
import type { WorkbookSchema } from '../src/schema/types';
import { SheetView } from '../src/state/sheetView';
import { sample } from './helpers';

/**
 * Tasks!A1:F5: Task, Owner, Start, End, Status, Done on
 *   Write spec | Ann | 2026-01-05 | 2026-01-09 | Done | 2026-01-09
 *   Build it   | Bob | 2026-01-12 | 2026-01-10 | Open |              (ends before it starts)
 *   Test it    | Cy  | 2026-01-20 | 2026-01-25 | Done |              (done, no date)
 *   Write spec | Ann | 2026-02-01 | 2026-02-02 | Open |              (same task and owner as row 2)
 */
async function tasks(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Tasks');
  ws.addRow(['Task', 'Owner', 'Start', 'End', 'Status', 'Done on']);
  const d = (s: string) => new Date(`${s}T00:00:00Z`);
  ws.addRow(['Write spec', 'Ann', d('2026-01-05'), d('2026-01-09'), 'Done', d('2026-01-09')]);
  ws.addRow(['Build it', 'Bob', d('2026-01-12'), d('2026-01-10'), 'Open', null]);
  ws.addRow(['Test it', 'Cy', d('2026-01-20'), d('2026-01-25'), 'Done', null]);
  ws.addRow(['Write spec', 'Ann', d('2026-02-01'), d('2026-02-02'), 'Open', null]);
  for (const c of ['C', 'D', 'F']) ws.getColumn(c).numFmt = 'yyyy-mm-dd';
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await tasks();
});

const schema = (extra: Partial<WorkbookSchema> = {}): WorkbookSchema => ({
  id: 'tasks',
  name: 'Tasks',
  version: '2',
  sheets: [
    {
      name: 'Tasks',
      tables: [
        {
          key: 'tasks',
          at: 'A1',
          fields: [
            { key: 'task', label: 'Task', type: 'text', maxLength: 8 },
            { key: 'owner', label: 'Owner', type: 'text' },
            { key: 'start', label: 'Start', type: 'date' },
            { key: 'end', label: 'End', type: 'date' },
            { key: 'status', label: 'Status', type: 'enum', values: ['Open', 'Done'] },
            { key: 'doneOn', label: 'Done on', type: 'date' },
          ],
          checks: [
            { rule: 'compare', field: 'end', op: '>=', other: 'start' },
            { rule: 'requiredIf', field: 'doneOn', when: 'status', equals: 'Done' },
            { rule: 'uniqueTogether', field: 'task', fields: ['task', 'owner'] },
          ],
        },
      ],
    },
  ],
  ...extra,
});

function control(s: WorkbookSchema, file = bytes) {
  const session = WorkbookSession.open(file, 'tasks.xlsx');
  const r = session.loadResult();
  const views = r.sheets.map((info, i) => new SheetView(info, r.grids[i]));
  const bound = bindSchema(s, views);
  if (!bound.binding) throw new Error(bound.problems.join('\n'));
  const ctl = new SchemaController(bound.binding, { views: () => views, styles: () => r.styles, date1904: () => false, followUp: async () => {}, changed: () => {} });
  ctl.validate();
  const at = (ref: string) => {
    const { row, col } = parseA1(ref)!;
    return ctl.cellIssues(0, row, col).map((i) => i.message);
  };
  return { session, ctl, at };
}

describe('more checks', () => {
  it('check text length and compare fields within and across records', () => {
    const { at } = control(schema());
    expect(at('A2')).toEqual(['At most 8 characters.']);
    expect(at('A3')).toEqual([]);
    expect(at('D3')).toEqual(['Should be no earlier than “Start”.']);
    expect(at('F4')).toEqual(['Required when “Status” is Done.']);
    expect(at('F3')).toEqual([]);
    expect(at('A5')).toEqual(['At most 8 characters.', 'Same “Task” and “Owner” as record 1.']);
  });

  it('are rejected on import when they name a field the table lacks', () => {
    const bad = schema();
    bad.sheets[0].tables![0].checks = [{ rule: 'compare', field: 'end', op: '>', other: 'finish' }];
    expect(() => schemaFromJSON(JSON.stringify(bad))).toThrow(/names "finish"/);
  });
});

describe('hook time limits', () => {
  it('report a slow field check instead of waiting for ever', async () => {
    const s = schema({ hookTimeoutMs: 30 });
    s.sheets[0].tables![0].fields[1].validate = () => new Promise(() => {});
    const { ctl, at } = control(s);
    expect(ctl.summary().pending).toBeGreaterThan(0);
    await ctl.settled();
    await new Promise((r) => setTimeout(r, 60));
    expect(at('B2')).toEqual(["Couldn't be checked: the check didn't finish within 30 ms."]);
  });

  it('stop a save whose last check runs too long', async () => {
    const s = schema({ hookTimeoutMs: 20, beforeSave: () => new Promise(() => {}) });
    // Without the data's own errors, which stop a save first.
    const t = s.sheets[0].tables![0];
    t.checks = [];
    t.fields[0] = { key: 'task', label: 'Task', type: 'text' };
    const { ctl } = control(s);
    const r = await ctl.checkBeforeSave();
    expect(r).toMatchObject({ ok: false, message: "The checks before saving didn't finish within 20 ms. Try saving again." });
  });
});

describe('schema versions', () => {
  it('stamp the saved file, which reads back', () => {
    const { session } = control(schema());
    session.apply({ type: 'input', sheet: 0, row: 2, col: 4, text: 'Done' });
    const saved = session.save({ WorkbookSchema: 'tasks@2' });
    const again = WorkbookSession.open(saved, 'tasks.xlsx');
    expect(again.loadResult().properties).toEqual({ WorkbookSchema: 'tasks@2' });
    // Saving again updates the stamp rather than adding another.
    again.apply({ type: 'input', sheet: 0, row: 2, col: 4, text: 'Open' });
    const third = WorkbookSession.open(again.save({ WorkbookSchema: 'tasks@3' }), 'tasks.xlsx');
    expect(third.loadResult().properties).toEqual({ WorkbookSchema: 'tasks@3' });
  });

  it('keep existing custom properties', () => {
    const s = WorkbookSession.open(sample('kitchen-sink.xlsx'), 'k.xlsx');
    s.apply({ type: 'input', sheet: 0, row: 40, col: 0, text: 'x' });
    const props = WorkbookSession.open(s.save({ A: '1' }), 'k.xlsx').loadResult().properties;
    const again = WorkbookSession.open(WorkbookSession.open(s.save({ A: '1' }), 'k.xlsx').save({ B: '2' }), 'k.xlsx');
    expect(props.A).toBe('1');
    expect(again.loadResult().properties.A).toBe('1');
  });
});

describe('generating', () => {
  it('finds every table and the label/value cells on a sheet', () => {
    const s = WorkbookSession.open(sample('kitchen-sink.xlsx'), 'k.xlsx');
    const r = s.loadResult();
    const views = r.sheets.map((info, i) => new SheetView(info, r.grids[i]));
    const summary = generateSchema(views, r.styles, 'k.xlsx').sheets.find((x) => x.name === 'Summary')!;
    expect(summary.tables!.map((t) => t.at)).toEqual(['A4', 'A21']);
    expect(summary.cells!.map((c) => [c.ref, c.label?.text, c.type])).toEqual(
      expect.arrayContaining([
        ['B2', 'Report date', 'date'],
        ['B11', 'Tax rate', 'percent'],
        ['B13', 'Top region', 'text'],
      ]),
    );
    expect(bindSchema(generateSchema(views, r.styles, 'k.xlsx'), views).problems).toEqual([]);
  });
});

describe('exporting schemas written in code', () => {
  it('leaves hooks out, including table rules, which JSON would turn into nulls', async () => {
    const { forecastSchema } = await import('../src/demo/schemas');
    const { schemaToJSON } = await import('../src/schema/json');
    expect(forecastSchema.sheets[0].tables![0].rules?.length).toBeGreaterThan(0);
    const json = schemaToJSON(forecastSchema);
    expect(json).not.toContain('null');
    expect(JSON.parse(json).sheets[0].tables[0].rules).toBeUndefined();
  });
});
