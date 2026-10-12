import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import type { EditOp } from '../src/core/types';
import { english, type Text } from '../src/i18n';
import type { ChangeContext, FieldContext, FieldValue, IssueInput, WorkbookData } from '../src/schema/types';
import { control, controllerWorkbook, FIXED, fixedSheet, FREE, LOG, logSheet, PLAN, planSheet, rc, workbookSchema } from './schema-fixture';

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await controllerWorkbook();
});

afterEach(() => {
  vi.useRealTimers();
});

const later = <T>(ms: number, value: T) => new Promise<T>((resolve) => setTimeout(() => resolve(value), ms));
const plan = (extra = {}) => workbookSchema([planSheet()], extra);
const messages = (list: { message: Text }[]) => list.map((i) => english.text(i.message));

describe('field check hooks', () => {
  it('report what a check returns or throws, sync or async', async () => {
    const s = workbookSchema([
      logSheet({
        note: {
          validate: (v) => {
            if (v === 'x') throw new Error('boom');
            return v === 'y' ? 'Not y' : undefined;
          },
        },
        kind: {
          validate: async (v) => {
            if (v === 'C') throw new Error('offline');
            return v === 'B' ? { message: { en: 'B is rare', de: 'B ist selten' }, severity: 'warning' } : null;
          },
        },
      }),
    ]);
    const h = control(bytes, s);
    expect(h.at(LOG, 'G4')).toEqual(['Check failed: boom']);
    expect(h.at(LOG, 'G5')).toEqual(['Not y']);
    expect(h.ctl.isPending(LOG, ...rc('F3'))).toBe(true);
    expect(h.ctl.summary().pending).toBe(4);
    await h.ctl.settled();
    expect(h.at(LOG, 'F3')).toEqual(['B is rare']);
    expect(h.ctl.cellIssues(LOG, ...rc('F3'))[0]).toMatchObject({ severity: 'warning', sheetName: 'Log', source: 'field' });
    expect(h.at(LOG, 'F5')).toEqual(['Check failed: offline']);
    expect(h.at(LOG, 'F2')).toEqual([]);
    expect(h.changes()).toBeGreaterThan(1);
  });

  it('may answer in several languages, shown in the one that fits', () => {
    const h = control(bytes, workbookSchema([logSheet({ seq: { validate: (v) => (v === 4 ? { en: 'Out of order', de: 'Falsche Reihenfolge' } : undefined) } })]));
    expect(h.ctl.cellIssues(LOG, ...rc('A4')).map((i) => i.message)).toEqual([{ en: 'Out of order', de: 'Falsche Reihenfolge' }]);
    expect(h.at(LOG, 'A4')).toEqual(['Out of order']);
  });

  it('get the record they belong to, and run on single cells too', () => {
    const seen: [FieldValue, Omit<FieldContext, 'signal'>][] = [];
    const s = workbookSchema([
      logSheet(
        { amount: { validate: (v, { signal: _signal, ...ctx }) => void seen.push([v, ctx]) } },
        {},
        { cells: [{ key: 'code', ref: 'J2', type: 'number', validate: (v, { signal: _signal, ...ctx }) => void seen.push([v, ctx]) }] },
      ),
    ], {});
    control(bytes, s, { language: 'de' });
    expect(seen[0]).toEqual([10.5, { sheet: 'Log', table: 'log', record: 0, language: 'de', values: expect.objectContaining({ seq: 1, amount: 10.5, note: 'hello' }) }]);
    expect(seen.at(-1)).toEqual([7.5, { sheet: 'Log', language: 'de' }]);
  });

  it('abort a check whose value changes, and keep a running check while its value stays the same', async () => {
    const calls: { value: FieldValue; signal: AbortSignal }[] = [];
    const s = workbookSchema([
      logSheet({
        note: {
          validate: (value, { signal }) =>
            new Promise((resolve, reject) => {
              calls.push({ value, signal });
              signal.addEventListener('abort', () => reject(new Error('aborted')));
              setTimeout(() => resolve(value === 'later' ? 'Too late' : null), 20);
            }),
        },
      }),
    ]);
    const h = control(bytes, s);
    expect(calls.map((c) => c.value)).toEqual(['hello', 'x', 'y']);
    h.type(LOG, 'G2', 'later');
    h.type(LOG, 'G4', '');
    expect(calls.map((c) => c.value)).toEqual(['hello', 'x', 'y', 'later']);
    expect(calls.map((c) => c.signal.aborted)).toEqual([true, true, false, false]);
    expect(h.ctl.isPending(LOG, ...rc('G4'))).toBe(false);
    await h.ctl.settled();
    // The aborted checks report nothing, not a failure.
    expect(h.ctl.issues().filter((i) => i.col === 6).map((i) => [i.row, english.text(i.message)])).toEqual([[1, 'Too late']]);
    // A finished check isn't run again for the same value.
    h.ctl.validate();
    expect(calls.length).toBe(4);
  });
});

describe('table, sheet and workbook hooks', () => {
  it('place issues by cell reference, record and field, falling back to the table', () => {
    const rule = (): IssueInput[] => [
      { message: 'by ref', ref: 'B3' },
      { message: 'on another sheet', ref: "'Plan'!C2:D3" },
      { message: 'by record', record: 1 },
      { message: 'by record and field', record: 2, field: 'note', severity: 'warning' },
      { message: 'by field', field: 'amount' },
      { message: 'unknown field', field: 'nothing' },
      { message: 'unknown sheet', ref: 'Nowhere!E5' },
      { message: { en: 'nowhere in particular', de: 'nirgends' } },
    ];
    const h = control(bytes, workbookSchema([logSheet({}, { rules: [rule] }), planSheet()]));
    expect(h.at(LOG, 'B3')).toEqual(['by ref']);
    expect(h.at(PLAN, 'C2')).toEqual(['on another sheet']);
    expect(h.at(LOG, 'A3')).toEqual(['by record']);
    expect(h.at(LOG, 'G4')).toEqual(['by record and field']);
    expect(h.at(LOG, 'C1')).toEqual(['by field']);
    expect(h.at(LOG, 'A1')).toEqual(['unknown field']);
    expect(h.at(LOG, 'E5')).toEqual(['unknown sheet']);
    const loose = h.ctl.issues().find((i) => english.text(i.message) === 'nowhere in particular');
    expect(loose).toMatchObject({ sheet: LOG, sheetName: 'Log', row: undefined, col: undefined, source: 'table', severity: 'error' });
  });

  it('report a hook that throws or rejects as a failed check', async () => {
    const h = control(
      bytes,
      workbookSchema(
        [
          logSheet({}, { rules: [() => { throw new Error('bad rule'); }, async () => Promise.reject(new Error('service down'))] }, { validate: () => [{ message: 'sheet says no', ref: 'A2' }] }),
          planSheet(),
        ],
        { validate: () => { throw 'plain'; } },
      ),
    );
    expect(h.at(LOG, 'A2')).toEqual(['sheet says no']);
    const wb = h.ctl.issues().filter((i) => i.source === 'workbook');
    expect(wb.map((i) => [english.text(i.message), i.sheet, i.sheetName])).toEqual([['Check failed: plain', -1, '']]);
    expect(messages(h.ctl.issues().filter((i) => i.source === 'table'))).toEqual(['Check failed: bad rule']);
    await h.ctl.settled();
    expect(messages(h.ctl.issues().filter((i) => i.source === 'table'))).toEqual(['Check failed: bad rule', 'Check failed: service down']);
  });

  it('give up on hooks that run past the time limit, saying how long they had', async () => {
    vi.useFakeTimers();
    const never = () => new Promise<IssueInput[]>(() => {});
    const h = control(bytes, workbookSchema([logSheet({}, { rules: [never] }, { validate: never }), planSheet()], { hookTimeoutMs: 2000, validate: never }));
    expect(h.ctl.summary().pending).toBe(3);
    await vi.advanceTimersByTimeAsync(2000);
    // Warnings, after the sheet's own error; workbook-wide ones first.
    expect(messages(h.ctl.issues().filter((i) => i.severity === 'warning'))).toEqual([
      "A workbook check didn't finish within 2 seconds, so its result is ignored.",
      "A table check didn't finish within 2 seconds, so its result is ignored.",
      "A sheet check didn't finish within 2 seconds, so its result is ignored.",
    ]);
    expect(h.ctl.summary().pending).toBe(0);
  });

  it('drop results and failures from runs that a newer one replaced', async () => {
    let run = 0;
    const rule = () => {
      run++;
      if (run === 1) return later(30, [{ message: 'stale result' }]);
      if (run === 2) return later(30, null).then(() => Promise.reject(new Error('stale failure')));
      return later(5, [{ message: 'fresh result' }]);
    };
    const h = control(bytes, workbookSchema([logSheet({}, { rules: [rule] })]));
    h.ctl.validate();
    h.ctl.validate();
    await h.ctl.settled();
    await later(40, null);
    expect(messages(h.ctl.issues().filter((i) => i.source === 'table'))).toEqual(['fresh result']);
  });

  it('see the workbook through the schema: sheets by name, tables, records, labels and cells', () => {
    let data: WorkbookData | undefined;
    control(bytes, workbookSchema([logSheet(), { ...planSheet(), name: /^pl/i }, fixedSheet()], { validate: (ctx) => void (data = ctx) }), { language: 'de' });
    expect(data!.language).toBe('de');
    expect(data!.signal.aborted).toBe(false);
    const fixed = data!.sheet('FIXED')!.table('fixed')!;
    expect(fixed.records.map((r) => [r.index, r.label, r.values, r.ref('price'), r.ref('nothing')])).toEqual([
      [0, 'North', { units: 5, price: 1.5 }, 'B3', 'B2'],
      [1, 'South', { units: 6, price: 2 }, 'C3', 'C2'],
    ]);
    // A sheet can be found by its name in the workbook or by the schema's pattern.
    expect(data!.sheet('plan')?.table('plan')?.records.length).toBe(3);
    expect(data!.sheet('Plan')?.name).toBe('Plan');
    expect(data!.sheet('Nowhere')).toBeUndefined();
    const log = data!.sheet('Log')!;
    expect(log.cell('owner')).toBe('Ann');
    expect(log.cell('nothing')).toBeNull();
    expect(log.table('nothing')).toBeUndefined();
    expect(log.table('log')!.records[0].label).toBeUndefined();
  });
});

describe('business logic after an edit', () => {
  it('hears about each change and writes values back as typed input, in one step', async () => {
    let heard: ChangeContext['changes'] = [];
    const onChange = (ctx: ChangeContext) => {
      heard = ctx.changes;
      const rec = (record: number, field: string) => ({ sheet: 'log', table: 'log', record, field });
      ctx.set(rec(1, 'kind'), 'c');
      ctx.set(rec(1, 'note'), new Date(Date.UTC(2026, 2, 1)));
      ctx.set(rec(1, 'paid'), true);
      ctx.set(rec(0, 'paid'), false);
      ctx.set(rec(2, 'note'), null);
      ctx.set(rec(3, 'note'), { error: '#N/A' });
      ctx.set({ sheet: 'Log', ref: 'J2' }, 12);
      // Refused like typing: a formula where formulas are forbidden.
      ctx.set(rec(0, 'amount'), '=1+1');
      // Places that aren't there are skipped.
      ctx.set({ sheet: 'Nowhere', ref: 'A1' }, 1);
      ctx.set({ sheet: 'Log', table: 'nothing', record: 0, field: 'seq' }, 1);
      ctx.set(rec(0, 'nothing'), 1);
      ctx.set({ sheet: 'Log', ref: 'not a cell' }, 1);
    };
    const h = control(bytes, workbookSchema([logSheet({ amount: { formula: 'forbid' } }), planSheet()], { onChange }), { applyFollowUps: false });
    const before = h.ctl.valueAt(LOG, ...rc('C2'));
    h.apply({ type: 'input', sheet: LOG, row: 1, col: 2, text: '11' });
    await h.ctl.afterEdit([
      { sheet: LOG, row: 1, col: 2, before },
      { sheet: LOG, row: 1, col: 9, before: 7.5 },
      { sheet: LOG, row: 20, col: 20, before: null },
      { sheet: FREE, row: 0, col: 0, before: 'anything' },
    ]);
    expect(heard).toEqual([
      { sheet: 'Log', table: 'log', record: 0, field: 'amount', ref: 'C2', before: 10.5, after: 11 },
      { sheet: 'Log', table: undefined, record: undefined, field: 'code', ref: 'J2', before: 7.5, after: 7.5 },
      { sheet: 'Log', table: undefined, record: undefined, field: undefined, ref: 'U21', before: null, after: null },
    ]);
    expect(h.followed).toEqual([
      {
        type: 'batch',
        ops: [
          {
            type: 'inputs',
            sheet: LOG,
            cells: [
              { row: 2, col: 5, text: "'C" },
              { row: 2, col: 6, text: '2026-03-01' },
              { row: 2, col: 4, text: 'TRUE' },
              { row: 1, col: 4, text: 'FALSE' },
              { row: 3, col: 6, text: '' },
              { row: 4, col: 6, text: '#N/A' },
              { row: 1, col: 9, text: '12' },
            ],
          },
        ],
      },
    ]);
  });

  it('is not run for edits outside the schema, and nothing is written when it sets nothing', async () => {
    const onChange = vi.fn();
    const h = control(bytes, workbookSchema([logSheet()], { onChange }));
    await h.ctl.afterEdit([{ sheet: FREE, row: 0, col: 0, before: null }]);
    expect(onChange).not.toHaveBeenCalled();
    await h.ctl.afterEdit([{ sheet: LOG, row: 1, col: 0, before: 1 }]);
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(h.followed).toEqual([]);
  });

  it('reports logic that fails or runs too long, and clears the report once it works again', async () => {
    let mode: 'fail' | 'slow' | 'ok' = 'fail';
    let signal: AbortSignal | undefined;
    const onChange = (ctx: ChangeContext) => {
      signal = ctx.signal;
      if (mode === 'fail') throw new Error('lookup failed');
      if (mode === 'slow') return new Promise<void>(() => {});
      ctx.set({ sheet: 'Log', ref: 'J2' }, 3);
    };
    const h = control(bytes, workbookSchema([logSheet()], { onChange, hookTimeoutMs: 20 }));
    const edit = [{ sheet: LOG, row: 1, col: 0, before: 1 as FieldValue }];
    await h.ctl.afterEdit(edit);
    expect(messages(h.ctl.issues().filter((i) => i.source === 'workbook'))).toEqual(['Business logic failed: lookup failed']);
    mode = 'slow';
    await h.ctl.afterEdit(edit);
    expect(messages(h.ctl.issues().filter((i) => i.source === 'workbook'))).toEqual([
      "Business logic didn't finish within 20 ms, so the changes it would have made weren't applied.",
    ]);
    expect(signal?.aborted).toBe(true);
    mode = 'ok';
    await h.ctl.afterEdit(edit);
    expect(h.ctl.issues().filter((i) => i.source === 'workbook')).toEqual([]);
    expect(h.ctl.valueAt(LOG, ...rc('J2'))).toBe(3);
  });

  it('carries the calculated field and formats into a record started below the table, keeping what was typed', async () => {
    const h = control(bytes, workbookSchema([logSheet(), fixedSheet()]), { applyFollowUps: false });
    const typed: [string, string][] = [['A6', '5'], ['C6', '=1+1'], ['D6', '7'], ['E6', 'TRUE'], ['F6', '#N/A'], ['G6', 'z']];
    for (const [ref, text] of typed) h.apply({ type: 'input', sheet: LOG, row: rc(ref)[0], col: rc(ref)[1], text });
    await h.ctl.afterEdit([{ sheet: LOG, row: 5, col: 0, before: null }]);
    const [batch] = h.followed as Extract<EditOp, { type: 'batch' }>[];
    expect(batch.ops[0]).toEqual({ type: 'copy', srcSheet: LOG, src: { r1: 4, c1: 7, r2: 4, c2: 7 }, sheet: LOG, dst: { r1: 5, c1: 7, r2: 5, c2: 7 } });
    // Cells left in the default format take the one above; what they hold stays.
    const restored = (batch.ops[1] as Extract<EditOp, { type: 'restore' }>).cells.map((c) => [c.col, c.content, c.style > 0]);
    expect(restored).toEqual([
      [1, { kind: 'empty' }, true],
      [3, { kind: 'number', value: 7 }, true],
      [4, { kind: 'bool', value: true }, true],
      [5, { kind: 'error', value: '#N/A' }, true],
      [6, { kind: 'string', value: 'z' }, true],
    ]);
  });

  it('keeps a formula typed into a calculated field of a new record', async () => {
    const h = control(bytes, workbookSchema([logSheet()]), { applyFollowUps: false });
    h.apply({ type: 'input', sheet: LOG, row: 5, col: 7, text: '=C6*3' });
    await h.ctl.afterEdit([{ sheet: LOG, row: 5, col: 7, before: null }]);
    const [batch] = h.followed as Extract<EditOp, { type: 'batch' }>[];
    expect(batch.ops.map((o) => o.type)).toEqual(['restore']);
  });
});

describe('adding and removing records', () => {
  const table = (h: ReturnType<typeof control>, sheet: number) => h.ctl.boundSheet(sheet)!.tables[0];
  const reason = (r: EditOp | Text) => english.text(r as Text);

  it('starts a record after the last one with only its calculated fields', () => {
    const h = control(bytes, workbookSchema([logSheet()]));
    const t = table(h, LOG);
    const op = h.ctl.insertRecord(t, 99) as EditOp;
    expect(op).toEqual({ type: 'batch', ops: [{ type: 'copy', srcSheet: LOG, src: { r1: 4, c1: 7, r2: 4, c2: 7 }, sheet: LOG, dst: { r1: 5, c1: 7, r2: 5, c2: 7 } }] });
    h.apply(op);
    expect(t.recordCount()).toBe(5);
    expect(h.views[LOG].get(5, 7)?.f).toBe('C6*2');
  });

  it('shifts later records along to insert one in the middle, across columns too', () => {
    const h = control(bytes, plan());
    const op = h.ctl.insertRecord(table(h, PLAN), 1);
    expect(op).toEqual({
      type: 'batch',
      ops: [
        { type: 'copy', srcSheet: PLAN, src: { r1: 0, c1: 2, r2: 2, c2: 3 }, sheet: PLAN, dst: { r1: 0, c1: 3, r2: 2, c2: 4 } },
        { type: 'clear', sheet: PLAN, rects: [{ r1: 0, c1: 2, r2: 2, c2: 2 }] },
      ],
    });
  });

  it('refuses to insert when there is nothing to carry, the table is full, or its records are fixed', () => {
    const h = control(bytes, plan());
    expect(reason(h.ctl.insertRecord(table(h, PLAN), 3))).toBe('There is nothing to insert.');
    const empty = control(bytes, workbookSchema([planSheet({ at: 'F1', header: false })]));
    expect(reason(empty.ctl.insertRecord(table(empty, PLAN), 0))).toBe('There is nothing to insert.');
    const rows = control(bytes, workbookSchema([logSheet({}, { records: { max: 4 } })]));
    expect(reason(rows.ctl.insertRecord(table(rows, LOG), 0))).toBe('This table can have at most 4 rows.');
    const full = control(bytes, workbookSchema([planSheet({ records: { max: 3 } })]));
    expect(reason(full.ctl.insertRecord(table(full, PLAN), 0))).toBe('This table can have at most 3 columns.');
    const fixed = control(bytes, workbookSchema([fixedSheet()]));
    expect(reason(fixed.ctl.insertRecord(table(fixed, FIXED), 0))).toBe("This table's columns are fixed by the schema.");
    expect(reason(fixed.ctl.deleteRecords(table(fixed, FIXED), 0, 0))).toBe("This table's columns are fixed by the schema.");
    expect(fixed.ctl.recordAt(FIXED, ...rc('B2'))).toBeUndefined();
  });

  it('closes the gap when deleting records, and clears a whole table', () => {
    const h = control(bytes, workbookSchema([logSheet()]));
    const t = table(h, LOG);
    h.apply(h.ctl.deleteRecords(t, 1, 1) as EditOp);
    expect(t.recordCount()).toBe(3);
    expect([2, 3, 4].map((r) => h.views[LOG].get(r, 0)?.n)).toEqual([4, 3, undefined]);
    const p = control(bytes, plan());
    expect(p.ctl.deleteRecords(table(p, PLAN), 0, 2)).toEqual({ type: 'batch', ops: [{ type: 'clear', sheet: PLAN, rects: [{ r1: 0, c1: 1, r2: 2, c2: 3 }] }] });
  });

  it('refuses deletes below the minimum, or with no record selected', () => {
    const h = control(bytes, workbookSchema([logSheet({}, { records: { min: 4 } }), planSheet()]));
    expect(reason(h.ctl.deleteRecords(table(h, LOG), 0, 0))).toBe('This table needs at least 4 rows.');
    expect(reason(h.ctl.deleteRecords(table(h, LOG), 6, 9))).toBe('Select a row in the table to delete.');
    expect(reason(h.ctl.deleteRecords(table(h, PLAN), 4, 5))).toBe('Select a column in the table to delete.');
    const p = control(bytes, workbookSchema([planSheet({ records: { min: 3 } })]));
    expect(reason(p.ctl.deleteRecords(table(p, PLAN), 2, 2))).toBe('This table needs at least 3 columns.');
  });

  it('finds the record at a cell, including the row where a new one would start', () => {
    const h = control(bytes, workbookSchema([logSheet()]));
    expect(h.ctl.recordAt(LOG, ...rc('C3'))).toMatchObject({ record: 1, append: false });
    expect(h.ctl.recordAt(LOG, ...rc('C6'))).toMatchObject({ record: 4, append: true });
    expect(h.ctl.recordAt(LOG, ...rc('J2'))).toBeUndefined();
    expect(h.ctl.recordAt(LOG, ...rc('C7'))).toBeUndefined();
  });
});

describe('checking before saving', () => {
  it('stops for errors, and lets the person override only when the schema says so', async () => {
    const block = control(bytes, workbookSchema([logSheet()]));
    const r = await block.ctl.checkBeforeSave();
    expect(r.ok || { ...r, message: english.text(r.message) }).toEqual({ ok: false, message: '1 cell breaks the “Control” schema.', canOverride: false });
  });

  it('passes when the last check returns true or nothing', async () => {
    expect(await control(bytes, plan({ beforeSave: () => true })).ctl.checkBeforeSave()).toEqual({ ok: true });
    expect(await control(bytes, plan({ beforeSave: async () => undefined })).ctl.checkBeforeSave()).toEqual({ ok: true });
    expect(await control(bytes, plan()).ctl.checkBeforeSave()).toEqual({ ok: true });
  });

  it("stops with the host's own reason, in one language or several", async () => {
    expect(await control(bytes, plan({ beforeSave: () => 'Approve it first.' })).ctl.checkBeforeSave()).toEqual({ ok: false, message: 'Approve it first.', canOverride: false });
    const reason = { en: 'Approve it first.', de: 'Zuerst freigeben.' };
    expect(await control(bytes, plan({ beforeSave: async () => reason })).ctl.checkBeforeSave()).toEqual({ ok: false, message: reason, canOverride: false });
  });

  it('stops when the last check fails', async () => {
    const r = await control(bytes, plan({ beforeSave: () => { throw new Error('no connection'); } })).ctl.checkBeforeSave();
    expect(r.ok || english.text(r.message)).toBe('The save check failed: no connection');
  });
});
