import ExcelJS from 'exceljs';
import { parseA1 } from '@sitepen/formula-engine/refs';
import { WorkbookSession } from '../src/core/session';
import type { EditOp } from '../src/core/types';
import { english } from '../src/i18n';
import { bindSchema } from '../src/schema/binding';
import { SchemaController } from '../src/schema/controller';
import type { FieldSchema, SheetSchema, TableSchema, WorkbookSchema } from '../src/schema/types';
import { SheetView } from '../src/state/sheetView';

const day = (s: string) => new Date(`${s}T00:00:00Z`);

/**
 * Log (records as rows), with single cells beside the table:
 *   Seq | When       | Amount | Share | Paid  | Kind | Note  | Total      | Owner | Ann
 *   1   | 2026-01-05 | 10.5   | 25%   | TRUE  | A    | hello | =C2*2      |       | 7.5
 *   2   | 2026-01-06 | 20     | 50%   | FALSE | B    |       | =C3*2      |       | 2026-03-01 (date)
 *   4   | 2026-01-04 | 5      | 12.5% | TRUE  | A    | x     | =1/0       |
 *   3   | 2026-01-08 12:00 | 7 | 10%  | FALSE | C    | y     | =C5*2      |
 * with Paid, Kind and Note in bold.
 * Plan (records as columns): Quarter Q1 Q2 Q3 / Sales 100 120 90 / Cost 50 60 70.
 * Fixed (records as columns, with labels): Metric North South / Units 5 6 / Price 1.5 2.
 * Free: not described by any schema.
 */
export async function controllerWorkbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const log = wb.addWorksheet('Log');
  log.addRow(['Seq', 'When', 'Amount', 'Share', 'Paid', 'Kind', 'Note', 'Total', 'Owner', 'Ann']);
  log.addRow([1, day('2026-01-05'), 10.5, 0.25, true, 'A', 'hello', { formula: 'C2*2', result: 21 }, null, 7.5]);
  log.addRow([2, day('2026-01-06'), 20, 0.5, false, 'B', null, { formula: 'C3*2', result: 40 }, null, day('2026-03-01')]);
  log.addRow([4, day('2026-01-04'), 5, 0.125, true, 'A', 'x', { formula: '1/0', result: { error: '#DIV/0!' } }]);
  log.addRow([3, new Date('2026-01-08T12:00:00Z'), 7, 0.1, false, 'C', 'y', { formula: 'C5*2', result: 14 }]);
  for (let r = 2; r <= 5; r++) {
    log.getCell(`B${r}`).numFmt = 'yyyy-mm-dd';
    log.getCell(`D${r}`).numFmt = '0%';
    for (const c of 'EFG') log.getCell(`${c}${r}`).font = { bold: true };
  }
  log.getCell('J3').numFmt = 'yyyy-mm-dd';
  const plan = wb.addWorksheet('Plan');
  plan.addRow(['Quarter', 'Q1', 'Q2', 'Q3']);
  plan.addRow(['Sales', 100, 120, 90]);
  plan.addRow(['Cost', 50, 60, 70]);
  const fixed = wb.addWorksheet('Fixed');
  fixed.addRow(['Metric', 'North', 'South']);
  fixed.addRow(['Units', 5, 6]);
  fixed.addRow(['Price', 1.5, 2]);
  wb.addWorksheet('Free').getCell('A1').value = 'anything';
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

export const LOG = 0;
export const PLAN = 1;
export const FIXED = 2;
export const FREE = 3;

const LOG_FIELDS: FieldSchema[] = [
  { key: 'seq', label: 'Seq', type: 'integer' },
  { key: 'when', label: 'When', type: 'date' },
  { key: 'amount', label: 'Amount', type: 'number' },
  { key: 'share', label: 'Share', type: 'percent' },
  { key: 'paid', label: 'Paid', type: 'boolean' },
  { key: 'kind', label: 'Kind', type: 'enum', values: ['A', 'B', 'C'] },
  { key: 'note', label: 'Note', type: 'text' },
  { key: 'total', label: 'Total', type: 'number', formula: 'require' },
];

/** The Log sheet's schema, with fields changed by key and anything else on the table or sheet. */
export function logSheet(fields: Record<string, Partial<FieldSchema>> = {}, table: Partial<TableSchema> = {}, sheet: Partial<SheetSchema> = {}): SheetSchema {
  return {
    name: 'Log',
    tables: [{ key: 'log', at: 'A1', fields: LOG_FIELDS.map((f) => ({ ...f, ...fields[f.key] }) as FieldSchema), ...table }],
    cells: [
      { key: 'owner', ref: 'J1', label: { ref: 'I1', text: 'Owner' }, type: 'text', readOnly: true },
      { key: 'code', ref: 'J2', type: 'number' },
      { key: 'stamp', ref: 'J3', type: 'text' },
    ],
    ...sheet,
  };
}

export function planSheet(table: Partial<TableSchema> = {}, sheet: Partial<SheetSchema> = {}): SheetSchema {
  return {
    name: 'Plan',
    tables: [
      {
        key: 'plan',
        at: 'A1',
        orientation: 'columns',
        fields: [
          { key: 'quarter', label: 'Quarter', type: 'text' },
          { key: 'sales', label: 'Sales', type: 'number' },
          { key: 'cost', label: 'Cost', type: 'number' },
        ],
        ...table,
      },
    ],
    ...sheet,
  };
}

export function fixedSheet(table: Partial<TableSchema> = {}): SheetSchema {
  return {
    name: 'Fixed',
    tables: [
      {
        key: 'fixed',
        at: 'A1',
        orientation: 'columns',
        corner: 'Metric',
        records: { labels: ['North', 'South'] },
        fields: [
          { key: 'units', label: 'Units', type: 'integer' },
          { key: 'price', label: 'Price', type: 'number' },
        ],
        ...table,
      },
    ],
  };
}

export function workbookSchema(sheets: SheetSchema[], extra: Partial<WorkbookSchema> = {}): WorkbookSchema {
  return { id: 'ctl', name: { en: 'Control', de: 'Steuerung' }, sheets, ...extra };
}

export interface Harness {
  session: WorkbookSession;
  views: SheetView[];
  ctl: SchemaController;
  /** Edits the host asked for, in order. */
  followed: EditOp[];
  /** How often the host was told issues changed. */
  changes(): number;
  /** Messages on a cell, in English. */
  at(sheet: number, ref: string): string[];
  /** Applies an edit to the workbook and the views, without re-validating. */
  apply(op: EditOp): void;
  /** Types into a cell and re-validates, as the editor does. */
  type(sheet: number, ref: string, text: string): void;
}

/** Opens the workbook as the editor does and puts a schema controller on it. */
export function control(bytes: Uint8Array, schema: WorkbookSchema, host: { language?: string; date1904?: boolean; applyFollowUps?: boolean } = {}): Harness {
  const session = WorkbookSession.open(bytes, 'ctl.xlsx');
  const r = session.loadResult();
  const views = r.sheets.map((info, i) => new SheetView(info, r.grids[i]));
  for (const u of session.ensureEngine().updates) views[u.sheet].apply(u);
  const apply = (op: EditOp) => {
    for (const u of session.apply(op).updates) views[u.sheet].apply(u);
  };
  const bound = bindSchema(schema, views);
  if (!bound.binding) throw new Error(bound.problems.map((p) => english.text(p)).join('\n'));
  const followed: EditOp[] = [];
  let changed = 0;
  const ctl: SchemaController = new SchemaController(bound.binding, {
    views: () => views,
    styles: () => r.styles,
    date1904: () => host.date1904 ?? false,
    followUp: async (op) => {
      followed.push(op);
      if (host.applyFollowUps !== false) {
        apply(op);
        ctl.validate();
      }
    },
    changed: () => {
      changed++;
    },
    ...(host.language ? { language: () => host.language! } : {}),
  });
  ctl.validate();
  return {
    session,
    views,
    ctl,
    followed,
    changes: () => changed,
    at: (sheet, ref) => {
      const { row, col } = parseA1(ref)!;
      return ctl.cellIssues(sheet, row, col).map((i) => english.text(i.message));
    },
    apply,
    type: (sheet, ref, text) => {
      const { row, col } = parseA1(ref)!;
      apply({ type: 'input', sheet, row, col, text });
      ctl.validate();
    },
  };
}

/** Row and column of an address, for the controller's (sheet, row, col) methods. */
export function rc(ref: string): [number, number] {
  const { row, col } = parseA1(ref)!;
  return [row, col];
}
