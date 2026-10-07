import { msg, type MessageKey, type Text } from '../i18n/message';
import { parseA1, toA1 } from '../core/refs';
import { VT, type Rect } from '../core/types';
import type { SheetView } from '../state/sheetView';
import { describeMatch, textMatches } from './json';
import type { CellSchema, FieldSchema, SheetSchema, TableSchema, TextMatch, WorkbookSchema } from './types';

export interface Pos {
  row: number;
  col: number;
}

export type Location =
  | { kind: 'corner' }
  | { kind: 'header'; field: number }
  | { kind: 'label'; record: number }
  /** `append` marks the first empty record after the table, where a new record can start. */
  | { kind: 'field'; record: number; field: number; append: boolean };

/** Text a header or label cell shows, for matching. */
export function cellText(view: SheetView, row: number, col: number): string {
  const rec = view.get(row, col);
  if (!rec) return '';
  if (rec.t === VT.String || rec.t === VT.Error) return rec.s ?? '';
  if (rec.t === VT.Number) return String(rec.n);
  if (rec.t === VT.Bool) return rec.n === 1 ? 'TRUE' : 'FALSE';
  return '';
}

function isEmpty(view: SheetView, row: number, col: number): boolean {
  const rec = view.get(row, col);
  return !rec || (rec.t === VT.None && rec.f === undefined);
}

/**
 * A table schema placed on a sheet. Records and fields are mapped to cells
 * according to the orientation, so the rest of the code can work in
 * record/field terms either way round.
 */
export class BoundTable {
  readonly r0: number;
  readonly c0: number;
  readonly byRow: boolean;
  readonly header: boolean;
  readonly labels?: TextMatch[];
  private countVersion = -1;
  private count = 0;

  constructor(
    readonly schema: TableSchema,
    readonly sheet: number,
    readonly view: SheetView,
  ) {
    const at = parseA1(schema.at) ?? { row: 0, col: 0 };
    this.r0 = at.row;
    this.c0 = at.col;
    this.byRow = (schema.orientation ?? 'rows') === 'rows';
    this.header = schema.header !== false;
    this.labels = schema.records?.labels?.length ? schema.records.labels : undefined;
  }

  get fields(): FieldSchema[] {
    return this.schema.fields;
  }

  get fixed(): boolean {
    return !!this.labels;
  }

  /** Offset of the first record from the table origin, along the record axis. */
  private get recordStart(): number {
    return this.header ? 1 : 0;
  }

  /** Offset of the first field from the table origin, along the field axis. */
  private get fieldStart(): number {
    return this.labels ? 1 : 0;
  }

  private place(recordPos: number, fieldPos: number): Pos {
    return this.byRow ? { row: this.r0 + recordPos, col: this.c0 + fieldPos } : { row: this.r0 + fieldPos, col: this.c0 + recordPos };
  }

  cell(record: number, field: number): Pos {
    return this.place(this.recordStart + record, this.fieldStart + field);
  }

  headerCell(field: number): Pos {
    return this.place(0, this.fieldStart + field);
  }

  labelCell(record: number): Pos {
    return this.place(this.recordStart + record, 0);
  }

  get cornerCell(): Pos {
    return { row: this.r0, col: this.c0 };
  }

  /** Cells covered by records first..last, all fields. */
  recordsRect(first: number, last: number): Rect {
    const a = this.cell(first, 0);
    const b = this.cell(last, this.fields.length - 1);
    return { r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col), r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col) };
  }

  /** Fixed tables have one record per label; others end at the first empty record. */
  recordCount(): number {
    if (this.labels) return this.labels.length;
    if (this.countVersion === this.view.version) return this.count;
    let n = 0;
    const limit = (this.byRow ? this.view.rowCount + this.view.info.hiddenRows.length : this.view.colCount + this.view.info.hiddenCols.length) + 1;
    while (n < limit) {
      let empty = true;
      for (let f = 0; f < this.fields.length && empty; f++) {
        const p = this.cell(n, f);
        if (!isEmpty(this.view, p.row, p.col)) empty = false;
      }
      if (empty) break;
      n++;
    }
    this.count = n;
    this.countVersion = this.view.version;
    return n;
  }

  /** Whether a record can be added after the current ones. */
  get canAppend(): boolean {
    const max = this.schema.records?.max;
    return !this.fixed && (max === undefined || this.recordCount() < max);
  }

  locate(row: number, col: number): Location | null {
    const recordPos = this.byRow ? row - this.r0 : col - this.c0;
    const fieldPos = this.byRow ? col - this.c0 : row - this.r0;
    if (recordPos < 0 || fieldPos < 0) return null;
    if (this.header && recordPos === 0) {
      if (this.labels && fieldPos === 0) return { kind: 'corner' };
      const f = fieldPos - this.fieldStart;
      return f >= 0 && f < this.fields.length ? { kind: 'header', field: f } : null;
    }
    const record = recordPos - this.recordStart;
    if (this.labels && fieldPos === 0) return record < this.labels.length ? { kind: 'label', record } : null;
    const field = fieldPos - this.fieldStart;
    if (field >= this.fields.length) return null;
    const count = this.recordCount();
    if (record < count) return { kind: 'field', record, field, append: false };
    if (record === count && this.canAppend) return { kind: 'field', record, field, append: true };
    return null;
  }

  /** Problems that stop this table from matching the sheet. */
  problems(sheetName: string): Text[] {
    const out: Text[] = [];
    // The message when the cell holds other text, and when it is empty.
    const check = (p: Pos, match: TextMatch, keys: readonly [MessageKey, MessageKey]) => {
      if (match === '*') return;
      const actual = cellText(this.view, p.row, p.col);
      if (!textMatches(match, actual)) {
        out.push(msg(actual ? keys[0] : keys[1], { cell: `${sheetName}!${toA1(p.row, p.col)}`, expected: describeMatch(match), actual }));
      }
    };
    if (this.header) this.fields.forEach((f, i) => check(this.headerCell(i), f.label, ['schema.bind.fieldLabel', 'schema.bind.fieldLabelEmpty']));
    this.labels?.forEach((l, i) => check(this.labelCell(i), l, ['schema.bind.recordLabel', 'schema.bind.recordLabelEmpty']));
    if (this.schema.corner !== undefined && this.header && this.labels) {
      check(this.cornerCell, this.schema.corner, ['schema.bind.cornerLabel', 'schema.bind.cornerLabelEmpty']);
    }
    return out;
  }
}

export interface BoundCell {
  schema: CellSchema;
  pos: Pos;
  labelPos?: Pos;
}

export interface BoundSheet {
  schema: SheetSchema;
  sheet: number;
  name: string;
  view: SheetView;
  tables: BoundTable[];
  cells: BoundCell[];
}

export interface Binding {
  schema: WorkbookSchema;
  sheets: BoundSheet[];
}

export interface BindResult {
  binding?: Binding;
  problems: Text[];
  /** How many of the schema's sheets were found, to pick the closest schema when none match. */
  found: number;
}

/** Places a schema on a workbook's sheets and reports anything that doesn't line up. */
export function bindSchema(schema: WorkbookSchema, views: SheetView[]): BindResult {
  const problems: Text[] = [];
  const sheets: BoundSheet[] = [];
  let found = 0;
  const used = new Set<number>();
  for (const ss of schema.sheets) {
    const index = views.findIndex((v, i) => !used.has(i) && v.info.kind === 'worksheet' && textMatches(ss.name, v.info.name));
    if (index < 0) {
      if (!ss.optional) problems.push(msg('schema.bind.noSheet', { expected: describeMatch(ss.name) }));
      continue;
    }
    found++;
    used.add(index);
    const view = views[index];
    const name = view.info.name;
    const tables = (ss.tables ?? []).map((t) => new BoundTable(t, index, view));
    for (const t of tables) problems.push(...t.problems(name));
    const cells: BoundCell[] = (ss.cells ?? []).map((c) => {
      const pos = parseA1(c.ref) ?? { row: 0, col: 0 };
      const labelPos = c.label ? parseA1(c.label.ref) ?? undefined : undefined;
      if (c.label && labelPos) {
        const actual = cellText(view, labelPos.row, labelPos.col);
        if (!textMatches(c.label.text, actual)) {
          problems.push(msg(actual ? 'schema.bind.label' : 'schema.bind.labelEmpty', { cell: `${name}!${c.label.ref}`, expected: describeMatch(c.label.text), actual }));
        }
      }
      return { schema: c, pos, labelPos };
    });
    sheets.push({ schema: ss, sheet: index, name, view, tables, cells });
  }
  return { binding: problems.length ? undefined : { schema, sheets }, problems, found };
}
