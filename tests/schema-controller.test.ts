import { beforeAll, describe, expect, it } from 'vitest';
import { CellFlag, VT } from '../src/core/types';
import { english } from '../src/i18n';
import { toJson } from '../src/schema/controller';
import type { DeclarativeRule, FieldSchema, WorkbookSchema } from '../src/schema/types';
import { control, controllerWorkbook, FIXED, fixedSheet, FREE, LOG, logSheet, PLAN, planSheet, rc, workbookSchema } from './schema-fixture';

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await controllerWorkbook();
});

const all = (extra: Partial<WorkbookSchema> = {}) => workbookSchema([logSheet(), planSheet(), fixedSheet()], extra);
const withLog = (fields: Record<string, Partial<FieldSchema>>, checks?: DeclarativeRule[]) => workbookSchema([logSheet(fields, checks ? { checks } : {})]);

describe('field checks', () => {
  it('check formulas: required, forbidden, failing, and waiting for Excel to recalculate', () => {
    const h = control(bytes, withLog({ amount: { formula: 'require' }, total: { formula: 'forbid' } }));
    expect(h.at(LOG, 'C2')).toEqual(['Should be calculated by a formula.']);
    expect(h.at(LOG, 'H2')).toEqual(["Formulas aren't allowed here; type a value."]);
    expect(h.at(LOG, 'H4')).toEqual(["Formulas aren't allowed here; type a value.", 'The formula returns #DIV/0!.']);
    // A formula whose inputs changed while recalculation is off can't be judged yet.
    h.views[LOG].apply({ sheet: LOG, row: 4, col: 7, t: VT.Number, n: 14, style: 0, f: 'C5*2', flags: CellFlag.Formula | CellFlag.Stale });
    h.ctl.validate();
    expect(h.at(LOG, 'H5')).toEqual(["Formulas aren't allowed here; type a value.", "Can't be checked until Excel recalculates this formula."]);
    expect(h.ctl.cellIssues(LOG, 4, 7).map((i) => i.severity)).toEqual(['error', 'warning']);
  });

  it('check the type of each value', () => {
    const h = control(bytes, withLog({ seq: { type: 'percent' }, note: { type: 'number', required: true }, kind: { type: 'integer' }, amount: { type: 'integer' }, paid: { type: 'date' }, when: { type: 'boolean' } }));
    expect(h.at(LOG, 'A2')).toEqual([]);
    expect(h.at(LOG, 'G2')).toEqual(['Expected a number.']);
    expect(h.at(LOG, 'G3')).toEqual(['Required.']);
    expect(h.at(LOG, 'F2')).toEqual(['Expected a whole number.']);
    expect(h.at(LOG, 'C2')).toEqual(['Expected a whole number.']);
    expect(h.at(LOG, 'C3')).toEqual([]);
    expect(h.at(LOG, 'E2')).toEqual(['Expected a date.']);
    expect(h.at(LOG, 'B2')).toEqual(['Expected TRUE or FALSE.']);
    const p = control(bytes, withLog({ note: { type: 'percent' } }));
    expect(p.at(LOG, 'G2')).toEqual(['Expected a percentage.']);
  });

  it('check decimal places, of the percentage for percents', () => {
    const h = control(bytes, withLog({ amount: { decimals: 0 }, share: { decimals: 0 } }));
    expect(h.at(LOG, 'C2')).toEqual(['No decimal places allowed.']);
    expect(h.at(LOG, 'D2')).toEqual([]);
    expect(h.at(LOG, 'D4')).toEqual(['No decimal places allowed.']);
    const one = control(bytes, withLog({ amount: { decimals: 1 }, share: { decimals: 1 } }));
    expect(one.at(LOG, 'C2')).toEqual([]);
    expect(one.at(LOG, 'D4')).toEqual([]);
    one.type(LOG, 'C3', '20.25');
    expect(one.at(LOG, 'C3')).toEqual(['At most 1 decimal place.']);
  });

  it('check bounds, writing percent bounds as percentages and date bounds as dates', () => {
    const h = control(bytes, withLog({ amount: { min: 6, max: 15 }, share: { min: 0.2 }, when: { min: '2026-01-05', max: '2026-01-06' } }));
    expect(h.at(LOG, 'C4')).toEqual(['Must be at least 6.']);
    expect(h.at(LOG, 'C3')).toEqual(['Must be at most 15.']);
    expect(h.at(LOG, 'C2')).toEqual([]);
    expect(h.at(LOG, 'D5')).toEqual(['Must be at least 20%.']);
    expect(h.at(LOG, 'B4')).toEqual(['Must be on or after 2026-01-05.']);
    expect(h.at(LOG, 'B5')).toEqual(['Must be on or before 2026-01-06.']);
    expect(h.at(LOG, 'B2')).toEqual([]);
  });

  it('check enum values, text length and patterns, reading dates as ISO text', () => {
    const h = control(
      bytes,
      withLog({
        kind: { values: ['A', 'B'] },
        note: { minLength: 3 },
        when: { maxLength: 5, pattern: '^2026-01-0[56]$' },
        seq: { pattern: '^[123]$', patternHint: { en: 'one, two or three', de: 'eins, zwei oder drei' } },
      }),
    );
    expect(h.at(LOG, 'F5')).toEqual(['Must be one of: A, B.']);
    expect(h.at(LOG, 'G4')).toEqual(['At least 3 characters.']);
    expect(h.at(LOG, 'G2')).toEqual([]);
    expect(h.at(LOG, 'B2')).toEqual(['At most 5 characters.']);
    expect(h.at(LOG, 'B4')).toEqual(['At most 5 characters.', "Doesn't match the expected format."]);
    expect(h.at(LOG, 'A4')).toEqual(['Should be one, two or three.']);
    // The hint is the host's own text, in the editor's language.
    const de = control(bytes, withLog({ seq: { pattern: '^[123]$', patternHint: { en: 'one, two or three', de: 'eins, zwei oder drei' } } }), { language: 'de' });
    expect(de.at(LOG, 'A4')).toEqual(['Should be eins, zwei oder drei.']);
  });

  it('find duplicates of dates and logicals', () => {
    const h = control(bytes, withLog({ when: { unique: true }, paid: { unique: true } }));
    expect(h.at(LOG, 'E4')).toEqual(['Duplicate: E2 has the same Paid.']);
    expect(h.at(LOG, 'E5')).toEqual(['Duplicate: E3 has the same Paid.']);
    expect(h.at(LOG, 'B3')).toEqual([]);
    h.type(LOG, 'B3', '2026-01-05');
    expect(h.at(LOG, 'B3')).toEqual(['Duplicate: B2 has the same When.']);
  });

  it('check single cells as well as tables', () => {
    const h = control(bytes, workbookSchema([logSheet({}, {}, { cells: [{ key: 'code', ref: 'J2', type: 'integer' }, { key: 'missing', ref: 'J9', type: 'text', required: true }] })]));
    expect(h.at(LOG, 'J2')).toEqual(['Expected a whole number.']);
    expect(h.at(LOG, 'J9')).toEqual(['Required.']);
  });
});

describe('declarative table checks', () => {
  it('check that values rise or fall from record to record, numbers and dates alike', () => {
    const h = control(bytes, withLog({}, [
      { rule: 'monotonic', field: 'seq', direction: 'increasing' },
      { rule: 'monotonic', field: 'when', direction: 'nonDecreasing', severity: 'warning' },
      { rule: 'monotonic', field: 'nothing', direction: 'increasing' },
      { rule: 'monotonic', field: 'total', direction: 'increasing' },
    ]));
    // Values that aren't numbers or dates are passed over.
    expect(h.at(LOG, 'H4')).toEqual(['The formula returns #DIV/0!.']);
    expect(h.at(LOG, 'H5')).toEqual(['Should be more than the value before it.']);
    expect(h.at(LOG, 'A5')).toEqual(['Should be more than the value before it.']);
    expect(h.at(LOG, 'A4')).toEqual([]);
    expect(h.at(LOG, 'B4')).toEqual(['Should be no less than the value before it.']);
    expect(h.ctl.cellIssues(LOG, ...rc('B4'))[0].severity).toBe('warning');
    const plan = control(bytes, workbookSchema([planSheet({ checks: [
      { rule: 'monotonic', field: 'sales', direction: 'decreasing' },
      { rule: 'monotonic', field: 'cost', direction: 'nonIncreasing' },
    ] })]));
    expect(plan.at(PLAN, 'C2')).toEqual(['Should be less than the value before it.']);
    expect(plan.at(PLAN, 'D2')).toEqual([]);
    expect([plan.at(PLAN, 'C3'), plan.at(PLAN, 'D3')]).toEqual([['Should be no more than the value before it.'], ['Should be no more than the value before it.']]);
  });

  it('check the first value, as a number or as text ignoring case', () => {
    const h = control(bytes, withLog({}, [
      { rule: 'firstEquals', field: 'seq', value: 1 },
      { rule: 'firstEquals', field: 'kind', value: 'a' },
      { rule: 'firstEquals', field: 'note', value: 'Goodbye' },
      { rule: 'firstEquals', field: 'amount', value: 11 },
    ]));
    expect(h.at(LOG, 'A2')).toEqual([]);
    expect(h.at(LOG, 'F2')).toEqual([]);
    expect(h.at(LOG, 'G2')).toEqual(['The first value should be Goodbye.']);
    expect(h.at(LOG, 'C2')).toEqual(['The first value should be 11.']);
    const empty = control(bytes, workbookSchema([planSheet({ at: 'F1', header: false, checks: [{ rule: 'firstEquals', field: 'sales', value: 1 }] })]));
    expect(empty.ctl.summary().errors).toBe(0);
  });

  it('check totals against the header of the field, or of every field with "*"', () => {
    const h = control(bytes, withLog({}, [{ rule: 'sum', field: 'amount', equals: 40, max: 40, min: 50 }]));
    expect(h.at(LOG, 'C1')).toEqual([
      '“Amount” adds up to 42.5; it should be 40.',
      '“Amount” adds up to 42.5, over the limit of 40.',
      '“Amount” adds up to 42.5, under the minimum of 50.',
    ]);
    const plan = control(bytes, workbookSchema([planSheet({ checks: [{ rule: 'sum', field: '*', max: 200 }] })]));
    // Records run across, so each field's header is in the first column.
    expect(plan.at(PLAN, 'A2')).toEqual(['“Sales” adds up to 310, over the limit of 200.']);
    expect(plan.at(PLAN, 'A1')).toEqual([]);
    expect(plan.at(PLAN, 'A3')).toEqual([]);
  });

  it('compare fields in each record, with wording for numbers, text and dates', () => {
    const ops = (op: '<' | '<=' | '>' | '>=' | '=' | '<>', field = 'cost', other = 'sales') =>
      control(bytes, workbookSchema([planSheet({ checks: [{ rule: 'compare', field, op, other }] })])).at(PLAN, 'B3');
    expect(ops('<')).toEqual([]);
    expect(ops('<=')).toEqual([]);
    expect(ops('>')).toEqual(['Should be more than “Sales”.']);
    expect(ops('>=')).toEqual(['Should be at least “Sales”.']);
    expect(ops('=')).toEqual(['Should be equal to “Sales”.']);
    expect(ops('<>', 'cost', 'cost')).toEqual(['Should be different from “Cost”.']);
    const dates = (op: '<' | '<=' | '>' | '=') => control(bytes, withLog({}, [{ rule: 'compare', field: 'when', op, other: 'when' }])).at(LOG, 'B2');
    expect(dates('<')).toEqual(['Should be before “When”.']);
    expect(dates('>')).toEqual(['Should be after “When”.']);
    expect(dates('<=')).toEqual([]);
    expect(dates('=')).toEqual([]);
    const text = control(bytes, withLog({}, [{ rule: 'compare', field: 'kind', op: '>', other: 'note' }]));
    // "A" sorts before "hello", ignoring case; a blank on either side is not compared.
    expect(text.at(LOG, 'F2')).toEqual(['Should be more than “Note”.']);
    expect(text.at(LOG, 'F3')).toEqual([]);
  });

  it('does not compare with a formula error, which is reported on its own cell', () => {
    const h = control(bytes, withLog({}, [{ rule: 'compare', field: 'amount', op: '<', other: 'total' }]));
    expect(h.at(LOG, 'C4')).toEqual([]);
    expect(h.at(LOG, 'H4')).toEqual(['The formula returns #DIV/0!.']);
  });

  it('require a field when another is filled in', () => {
    const h = control(bytes, withLog({}, [{ rule: 'requiredIf', field: 'note', when: 'seq' }]));
    expect(h.at(LOG, 'G3')).toEqual(['Required when “Seq” is filled in.']);
    expect(h.at(LOG, 'G2')).toEqual([]);
  });
});

describe('record limits', () => {
  it('report too few or too many rows where the table ends', () => {
    const few = control(bytes, workbookSchema([logSheet({}, { records: { min: 6 } })]));
    expect(few.at(LOG, 'A6')).toEqual(['This table needs at least 6 rows.']);
    const many = control(bytes, workbookSchema([logSheet({}, { records: { max: 3 } })]));
    expect(many.at(LOG, 'A5')).toEqual(['This table can have at most 3 rows.']);
  });

  it('count columns for tables whose records run across', () => {
    const few = control(bytes, workbookSchema([planSheet({ records: { min: 4 } })]));
    expect(few.at(PLAN, 'E1')).toEqual(['This table needs at least 4 columns.']);
    const many = control(bytes, workbookSchema([planSheet({ records: { max: 2 } })]));
    expect(many.at(PLAN, 'D1')).toEqual(['This table can have at most 2 columns.']);
  });

  it('do not apply to tables with fixed labels', () => {
    const h = control(bytes, workbookSchema([fixedSheet({ records: { labels: ['North', 'South'], min: 5, max: 1 } })]));
    expect(h.ctl.summary()).toEqual({ errors: 0, warnings: 0, pending: 0 });
  });
});

describe('issues', () => {
  it('come errors first, then by sheet and position, and are counted', () => {
    const h = control(bytes, workbookSchema([logSheet({ note: { required: true } }, { checks: [{ rule: 'monotonic', field: 'seq', direction: 'increasing', severity: 'warning' }] }), planSheet({ records: { max: 2 } })]));
    const list = h.ctl.issues().map((i) => `${i.severity} ${i.sheetName}!${i.row},${i.col}`);
    expect(list).toEqual(['error Log!2,6', 'error Log!3,7', 'error Plan!0,3', 'warning Log!4,0']);
    expect(h.ctl.summary()).toEqual({ errors: 3, warnings: 1, pending: 0 });
  });
});

describe('values', () => {
  it('hand records and cells over as JSON: dates as ISO text, errors as their text', () => {
    const h = control(bytes, all());
    const v = h.ctl.values();
    expect(v.Log.tables.log).toEqual([
      { seq: 1, when: '2026-01-05', amount: 10.5, share: 0.25, paid: true, kind: 'A', note: 'hello', total: 21 },
      { seq: 2, when: '2026-01-06', amount: 20, share: 0.5, paid: false, kind: 'B', note: null, total: 40 },
      { seq: 4, when: '2026-01-04', amount: 5, share: 0.125, paid: true, kind: 'A', note: 'x', total: '#DIV/0!' },
      { seq: 3, when: '2026-01-08T12:00:00Z', amount: 7, share: 0.1, paid: false, kind: 'C', note: 'y', total: 14 },
    ]);
    // A text cell formatted as a date is read as the date it shows.
    expect(v.Log.cells).toEqual({ owner: 'Ann', code: 7.5, stamp: '2026-03-01' });
    expect(v.Plan.tables.plan).toEqual([
      { quarter: 'Q1', sales: 100, cost: 50 },
      { quarter: 'Q2', sales: 120, cost: 60 },
      { quarter: 'Q3', sales: 90, cost: 70 },
    ]);
    expect(v.Fixed.tables.fixed).toEqual([
      { units: 5, price: 1.5 },
      { units: 6, price: 2 },
    ]);
  });

  it('count dates from 1904 in workbooks that do', () => {
    const h = control(bytes, all(), { date1904: true });
    expect(h.ctl.values().Log.tables.log[0].when).toBe('2030-01-06');
  });

  it('read cells outside any field by their format, and nothing on sheets that are not there', () => {
    const h = control(bytes, all());
    expect(h.ctl.valueAt(LOG, ...rc('J3'))).toEqual(new Date(Date.UTC(2026, 2, 1)));
    expect(h.ctl.valueAt(LOG, ...rc('J2'))).toBe(7.5);
    expect(h.ctl.valueAt(FREE, ...rc('A1'))).toBe('anything');
    expect(h.ctl.valueAt(9, 0, 0)).toBeNull();
    expect(toJson(new Date(Date.UTC(2026, 0, 1, 9, 30)))).toBe('2026-01-01T09:30:00Z');
    expect(toJson(null)).toBeNull();
  });
});

describe('editing', () => {
  const h = () => control(bytes, workbookSchema([logSheet({ total: { readOnly: true } }), planSheet({ records: { max: 3 } }), fixedSheet()], { otherSheets: 'readOnly' }));
  const reason = (x: ReturnType<typeof h>, sheet: number, ref: string) => {
    const r = x.ctl.readOnlyReason(sheet, ...rc(ref));
    return r && english.text(r);
  };

  it('explains why labels, read-only fields and cells outside the schema are locked', () => {
    const x = h();
    expect(reason(x, LOG, 'A1')).toBe('Column names in this table are set by the “Control” schema.');
    expect(reason(x, LOG, 'H2')).toBe('“Total” is read-only.');
    expect(reason(x, LOG, 'C2')).toBeUndefined();
    expect(reason(x, LOG, 'I1')).toBe('This label is set by the “Control” schema.');
    expect(reason(x, LOG, 'J1')).toBe('“owner” is read-only.');
    expect(reason(x, LOG, 'J2')).toBeUndefined();
    expect(reason(x, LOG, 'L9')).toBe("Only the “Control” schema's tables and fields can be edited here. Add a row straight after the table, or use Insert.");
    expect(reason(x, PLAN, 'A2')).toBe('Row names in this table are set by the “Control” schema.');
    // Plan is full, so there is nowhere to add a column.
    expect(reason(x, PLAN, 'E1')).toBe("Only the “Control” schema's tables and fields can be edited here.");
    expect(reason(x, FIXED, 'B1')).toBe('Column labels in this table are set by the “Control” schema.');
    expect(reason(x, FIXED, 'A1')).toBe('Column labels in this table are set by the “Control” schema.');
    expect(reason(x, FREE, 'A1')).toBe('Only the sheets described by the “Control” schema can be edited.');
  });

  it('names the place to add a column when records run across, and the schema in the editor language', () => {
    const x = control(bytes, workbookSchema([planSheet()]), { language: 'de' });
    expect(reason(x, PLAN, 'A9')).toBe("Only the “Steuerung” schema's tables and fields can be edited here. Add a column straight after the table, or use Insert.");
    expect(reason(x, PLAN, 'E2')).toBeUndefined();
    // Sheets the schema doesn't describe stay editable unless it says otherwise.
    expect(reason(x, FREE, 'A1')).toBeUndefined();
    const open = control(bytes, workbookSchema([planSheet({}, { outside: 'editable' })]));
    expect(reason(open, PLAN, 'A9')).toBeUndefined();
  });

  it('tells labels from fields', () => {
    const x = h();
    expect([x.ctl.isLabel(LOG, ...rc('A1')), x.ctl.isLabel(LOG, ...rc('I1')), x.ctl.isLabel(FIXED, ...rc('A1')), x.ctl.isLabel(FIXED, ...rc('B1'))]).toEqual([true, true, true, true]);
    expect([x.ctl.isLabel(LOG, ...rc('A2')), x.ctl.isLabel(LOG, ...rc('J1')), x.ctl.isLabel(FREE, 0, 0)]).toEqual([false, false, false]);
  });
});

describe('describing fields', () => {
  const describe_ = (fields: Record<string, Partial<FieldSchema>>, ref: string, sheet = LOG, language?: string) => {
    const x = control(bytes, workbookSchema([logSheet(fields), planSheet()]), { language });
    const d = x.ctl.describe(sheet, ...rc(ref));
    return d && english.text(d);
  };

  it('lists type, precision, bounds and rules', () => {
    expect(describe_({ amount: { decimals: 2, min: 0, max: 100, required: true, unique: true, formula: 'forbid', description: 'What was paid.' } }, 'C2')).toBe(
      'Amount · number · up to 2 dp · ≥ 0 · ≤ 100 · required · unique · no formulas. What was paid.',
    );
    expect(describe_({ share: { decimals: 0, min: 0.05, max: 1 } }, 'D2')).toBe('Share · percent · no decimals · ≥ 5% · ≤ 100%');
    expect(describe_({ when: { min: '2026-01-01' } }, 'B2')).toBe('When · date · ≥ 2026-01-01');
    expect(describe_({}, 'F2')).toBe('Kind · one of A, B, C');
    expect(describe_({ kind: { values: undefined } }, 'F2')).toBe('Kind · one of ');
    expect(describe_({}, 'H2')).toBe('Total · number · formula');
    expect(describe_({ seq: { pattern: '^\\d$', patternHint: { en: 'one digit', de: 'eine Ziffer' } } }, 'A2', LOG, 'de')).toBe('Seq · integer · eine Ziffer');
    expect(describe_({}, 'E2')).toBe('Paid · boolean');
    expect(describe_({}, 'G2')).toBe('Note · text');
  });

  it('says where a new record would start, and describes single cells by label or key', () => {
    expect(describe_({}, 'A6')).toBe('Seq · integer · new row');
    expect(describe_({}, 'E1', PLAN)).toBe('Quarter · text · new column');
    expect(describe_({}, 'J1')).toBe('Owner · text');
    expect(describe_({}, 'J2')).toBe('code · number');
    expect(describe_({}, 'Z9')).toBeUndefined();
  });
});

describe('typed input', () => {
  const x = () => control(bytes, workbookSchema([logSheet({ amount: { formula: 'forbid' } })]));
  const prep = (ref: string, text: string) => {
    const r = x().ctl.prepareInput(LOG, ...rc(ref), text);
    return 'reject' in r ? { reject: english.text(r.reject) } : r.text;
  };

  it('stores text and enum values as text, matching enum values ignoring case', () => {
    expect(prep('G2', '0123')).toBe("'0123");
    expect(prep('G2', "'kept")).toBe("'kept");
    expect(prep('F2', ' b ')).toBe("'B");
    expect(prep('F2', 'other')).toBe("'other");
  });

  it('reads yes and no as logicals, and plain numbers in percent fields as percentages', () => {
    expect(['yes', 'Y', '1', 'true'].map((t) => prep('E2', t))).toEqual(['TRUE', 'TRUE', 'TRUE', 'TRUE']);
    expect(['no', 'N', '0', 'FALSE'].map((t) => prep('E2', t))).toEqual(['FALSE', 'FALSE', 'FALSE', 'FALSE']);
    expect(prep('E2', 'maybe')).toBe("'maybe");
    expect(prep('D2', '12')).toBe('12%');
    expect(prep('D2', '-1,000.5')).toBe('-1,000.5%');
    expect(prep('D2', '12%')).toBe('12%');
  });

  it('passes numbers, dates, formulas and empty text through, and refuses formulas where they are forbidden', () => {
    expect(prep('B2', '2026-02-01')).toBe('2026-02-01');
    expect(prep('A2', '12')).toBe('12');
    expect(prep('H2', '=C2*3')).toBe('=C2*3');
    expect(prep('G2', '=')).toBe("'=");
    expect(prep('G2', '')).toBe('');
    expect(prep('Z9', 'free')).toBe('free');
    expect(prep('C2', '=1+1')).toEqual({ reject: "Formulas aren't allowed in “Amount”. Type a value instead." });
  });
});
