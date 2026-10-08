import { msg, pickLanguage, type Localized, type LocalizedMap, type Message, type MessageKey, type Text } from '../i18n/message';
import { parseA1, toA1 } from '../core/refs';
import { CellFlag, VT, type CellContent, type DisplayStyle, type EditOp } from '../core/types';
import { isDateStyle, type CellRecord } from '../state/format';
import type { SheetView } from '../state/sheetView';
import { cellText, type Binding, type BoundCell, type BoundSheet, type BoundTable, type Pos } from './binding';
import { matchLabel, textMatches } from './json';
import type {
  CellSchema,
  ChangeContext,
  CheckResult,
  DeclarativeRule,
  FieldContext,
  FieldSchema,
  FieldValue,
  IssueInput,
  JsonValue,
  SchemaValues,
  SetTarget,
  Severity,
  SheetData,
  TableData,
  WorkbookData,
} from './types';

export interface Issue {
  severity: Severity;
  message: Text;
  /** Sheet index, or -1 for workbook-wide issues. */
  sheet: number;
  sheetName: string;
  row?: number;
  col?: number;
  source: 'field' | 'table' | 'sheet' | 'workbook';
}

export interface SchemaHost {
  views(): SheetView[];
  styles(): DisplayStyle[];
  date1904(): boolean;
  /** Applies an edit made by schema logic as part of the most recent undo step. */
  followUp(op: EditOp): Promise<void>;
  /** Tells the UI that issues changed. */
  changed(): void;
  /**
   * The editor's language (BCP 47, such as "en-US" or "de"). Hooks get it in
   * their context, and names the host gave in several languages are shown in
   * it. Default "en".
   */
  language?(): string;
}

export interface FieldRef {
  sheet: BoundSheet;
  field: FieldSchema | CellSchema;
  name: string;
  table?: BoundTable;
  cell?: BoundCell;
  record?: number;
  fieldIndex?: number;
  append?: boolean;
}

type Lite = { severity: Severity; message: Text };

/** An issue a check found, before it is placed on a cell: the host's, or the editor's own message. */
type Found = Omit<IssueInput, 'message'> & { message: Text };

const TYPE_NAMES: Record<FieldSchema['type'], MessageKey> = {
  text: 'schema.describe.text',
  number: 'schema.describe.number',
  integer: 'schema.describe.integer',
  percent: 'schema.describe.percent',
  date: 'schema.describe.date',
  boolean: 'schema.describe.boolean',
  enum: 'schema.describe.enum',
};

const TIMEOUTS: Record<Issue['source'], MessageKey> = {
  field: 'schema.hook.fieldTimeout',
  table: 'schema.hook.tableTimeout',
  sheet: 'schema.hook.sheetTimeout',
  workbook: 'schema.hook.workbookTimeout',
};

const MONOTONIC: Record<Extract<DeclarativeRule, { rule: 'monotonic' }>['direction'], MessageKey> = {
  increasing: 'schema.check.increasing',
  decreasing: 'schema.check.decreasing',
  nonIncreasing: 'schema.check.nonIncreasing',
  nonDecreasing: 'schema.check.nonDecreasing',
};

type CompareOp = Extract<DeclarativeRule, { rule: 'compare' }>['op'];

/** How one value should compare with another: numbers and text. */
const COMPARE: Record<CompareOp, MessageKey> = {
  '<': 'schema.compare.lessThan',
  '<=': 'schema.compare.atMost',
  '>': 'schema.compare.moreThan',
  '>=': 'schema.compare.atLeast',
  '=': 'schema.compare.equalTo',
  '<>': 'schema.compare.differentFrom',
};

/** How one value should compare with another: dates. */
const COMPARE_DATES: Record<CompareOp, MessageKey> = {
  '<': 'schema.compare.before',
  '<=': 'schema.compare.noLaterThan',
  '>': 'schema.compare.after',
  '>=': 'schema.compare.noEarlierThan',
  '=': 'schema.compare.sameAs',
  '<>': 'schema.compare.differentFrom',
};

const EPOCH = Date.UTC(1899, 11, 30);
const DAY = 86_400_000;

const DEFAULT_HOOK_TIMEOUT_MS = 10_000;

/** A field value as JSON: dates at midnight as "YYYY-MM-DD", other dates as full ISO text. */
export function toJson(v: FieldValue): JsonValue {
  if (v instanceof Date) {
    const iso = v.toISOString();
    return iso.endsWith('T00:00:00.000Z') ? iso.slice(0, 10) : iso.replace(/\.000Z$/, 'Z');
  }
  if (v !== null && typeof v === 'object') return v.error;
  return v;
}

/** A hook that ran past its time limit. */
class HookTimeout extends Error {
  constructor(readonly ms: number) {
    super(`didn't finish within ${ms} ms`);
  }

  /** The time limit as people read it: "30 ms", "10 seconds". */
  get duration(): Message {
    const ms = this.ms;
    return ms % 1000 ? msg('schema.duration.milliseconds', { n: ms }) : msg('schema.duration.seconds', { count: ms / 1000 });
  }
}

/** Why a hook failed, from whatever it threw. */
function failure(e: unknown): string {
  return (e as Error)?.message ?? String(e);
}

/** A field label inside a message: “Task”. */
function quoted(text: string): Message {
  return msg('schema.quoted', { text });
}

/** Several labels as one: “Task” and “Owner”. */
function and(items: Text[]): Text {
  return items.reduce((first, next) => msg('schema.and', { first, next }));
}

/** The hook's result, or a HookTimeout once it has run longer than `ms` (calling `onTimeout`, e.g. to abort it). */
function withTimeout<T>(p: Promise<T>, ms: number, onTimeout?: () => void): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const late = new Promise<never>((_, reject) => {
    timer = setTimeout(() => {
      onTimeout?.();
      reject(new HookTimeout(ms));
    }, ms);
  });
  return Promise.race([p, late]).finally(() => clearTimeout(timer));
}

function isPromise<T>(v: unknown): v is Promise<T> {
  return !!v && typeof (v as Promise<T>).then === 'function';
}

function decimalsOf(n: number): number {
  const s = String(parseFloat(n.toPrecision(15)));
  const m = /(?:\.(\d+))?(?:e([+-]\d+))?$/.exec(s);
  return Math.max(0, (m?.[1]?.length ?? 0) - Number(m?.[2] ?? 0));
}

function toLite(r: CheckResult): Lite | null {
  if (!r) return null;
  if (typeof r === 'string') return { severity: 'error', message: r };
  // { message, severity } from a hook, or the message itself in several languages.
  const found = r as { message?: Localized; severity?: Severity };
  if (found.message !== undefined) return { severity: found.severity ?? 'error', message: found.message };
  return { severity: 'error', message: r as LocalizedMap };
}

function contentOf(rec: CellRecord | null): CellContent {
  if (!rec) return { kind: 'empty' };
  if (rec.f !== undefined) return { kind: 'formula', formula: rec.f };
  switch (rec.t) {
    case VT.Number:
      return { kind: 'number', value: rec.n ?? 0 };
    case VT.String:
      return { kind: 'string', value: rec.s ?? '' };
    case VT.Bool:
      return { kind: 'bool', value: rec.n === 1 };
    case VT.Error:
      return { kind: 'error', value: rec.s ?? '#VALUE!' };
    default:
      return { kind: 'empty' };
  }
}

function isoDate(d: Date): string {
  return d.toISOString().slice(0, 10);
}

function valueKey(v: FieldValue): string {
  if (v instanceof Date) return `d:${v.getTime()}`;
  if (v && typeof v === 'object') return `e:${v.error}`;
  return `${typeof v}:${String(v)}`;
}

/**
 * Sets a property, even one called __proto__: keys come from schemas and
 * sheet names come from files, and plain assignment would set the object's
 * prototype instead of a value.
 */
function put<T>(obj: Record<string, T>, key: string, value: T): void {
  Object.defineProperty(obj, key, { value, enumerable: true, writable: true, configurable: true });
}

const patterns = new Map<string, RegExp>();
function regex(source: string): RegExp {
  let re = patterns.get(source);
  if (!re) patterns.set(source, (re = new RegExp(source)));
  return re;
}

/**
 * Applies one schema to an open workbook: validates it, decides which cells
 * can be edited, prepares typed input, runs the schema's hooks and plans
 * record inserts and deletes.
 */
export class SchemaController {
  private syncIssues: Issue[] = [];
  private readonly hookResults = new Map<string, Issue[]>();
  private readonly pendingHooks = new Map<string, Promise<void>>();
  private readonly fieldChecks = new Map<string, { key: string; issue: Lite | null; pending?: Promise<void>; abort: AbortController }>();
  private index = new Map<string, Issue[]>();
  private hookAbort = new AbortController();
  private readonly counts = new Map<BoundTable, number>();
  private generation = 0;

  /** How long async hooks may run. */
  private get timeoutMs(): number {
    return this.schema.hookTimeoutMs ?? DEFAULT_HOOK_TIMEOUT_MS;
  }

  constructor(
    readonly binding: Binding,
    private readonly host: SchemaHost,
  ) {}

  get schema() {
    return this.binding.schema;
  }

  // ---------------------------------------------------------------------------
  // Where things are

  boundSheet(sheet: number): BoundSheet | undefined {
    return this.binding.sheets.find((s) => s.sheet === sheet);
  }

  fieldAt(sheet: number, row: number, col: number): FieldRef | undefined {
    const bs = this.boundSheet(sheet);
    if (!bs) return undefined;
    for (const table of bs.tables) {
      const loc = table.locate(row, col);
      if (loc?.kind === 'field') {
        const field = table.fields[loc.field];
        return { sheet: bs, table, field, name: matchLabel(field.label), record: loc.record, fieldIndex: loc.field, append: loc.append };
      }
    }
    const cell = bs.cells.find((c) => c.pos.row === row && c.pos.col === col);
    if (cell) return { sheet: bs, cell, field: cell.schema, name: cell.schema.label ? matchLabel(cell.schema.label.text) : cell.schema.key };
    return undefined;
  }

  /** Header, record label or cell label positions fixed by the schema. */
  isLabel(sheet: number, row: number, col: number): boolean {
    const bs = this.boundSheet(sheet);
    if (!bs) return false;
    for (const t of bs.tables) {
      const loc = t.locate(row, col);
      if (loc && loc.kind !== 'field') return true;
    }
    return bs.cells.some((c) => c.labelPos?.row === row && c.labelPos.col === col);
  }

  /** The editor's language. */
  private get language(): string {
    return this.host.language?.() ?? 'en';
  }

  /** Text the host gave in one language or several, in the editor's language. */
  private localized(text: Localized): string {
    return pickLanguage(text, this.language);
  }

  /** Why a cell can't be edited, if it can't. */
  readOnlyReason(sheet: number, row: number, col: number): Message | undefined {
    const bs = this.boundSheet(sheet);
    const schema = this.localized(this.schema.name);
    if (!bs) return this.schema.otherSheets === 'readOnly' ? msg('schema.readOnly.otherSheet', { schema }) : undefined;
    for (const t of bs.tables) {
      const loc = t.locate(row, col);
      if (!loc) continue;
      if (loc.kind === 'header') return msg(t.byRow ? 'schema.readOnly.columnNames' : 'schema.readOnly.rowNames', { schema });
      if (loc.kind === 'label' || loc.kind === 'corner') return msg(t.byRow ? 'schema.readOnly.rowLabels' : 'schema.readOnly.columnLabels', { schema });
      const f = t.fields[loc.field];
      return f.readOnly ? msg('schema.readOnly.field', { field: matchLabel(f.label) }) : undefined;
    }
    for (const c of bs.cells) {
      if (c.labelPos?.row === row && c.labelPos.col === col) return msg('schema.readOnly.label', { schema });
      if (c.pos.row === row && c.pos.col === col) return c.schema.readOnly ? msg('schema.readOnly.field', { field: c.schema.key }) : undefined;
    }
    if (bs.schema.outside === 'editable') return undefined;
    const append = bs.tables.find((t) => t.canAppend);
    if (!append) return msg('schema.readOnly.outside', { schema });
    return msg(append.byRow ? 'schema.readOnly.outsideAddRow' : 'schema.readOnly.outsideAddColumn', { schema });
  }

  /** One-line description of the field at a cell, for the formula bar: "Price · number · up to 2 dp · ≥ 0 · required". */
  describe(sheet: number, row: number, col: number): Text | undefined {
    const ref = this.fieldAt(sheet, row, col);
    if (!ref) return undefined;
    const f = ref.field;
    const parts: Text[] = [ref.name, f.type === 'enum' ? msg('schema.describe.enum', { values: f.values?.join(', ') ?? '' }) : msg(TYPE_NAMES[f.type])];
    if (f.decimals !== undefined) parts.push(f.decimals === 0 ? msg('schema.describe.noDecimals') : msg('schema.describe.decimals', { count: f.decimals }));
    if (f.min !== undefined) parts.push(msg('schema.describe.min', { bound: this.bound(f, f.min) }));
    if (f.max !== undefined) parts.push(msg('schema.describe.max', { bound: this.bound(f, f.max) }));
    if (f.patternHint) parts.push(this.localized(f.patternHint));
    if (f.required) parts.push(msg('schema.describe.required'));
    if (f.unique) parts.push(msg('schema.describe.unique'));
    if (f.formula === 'require') parts.push(msg('schema.describe.formula'));
    if (f.formula === 'forbid') parts.push(msg('schema.describe.noFormulas'));
    if (ref.append) parts.push(msg(ref.table?.byRow ? 'schema.describe.newRow' : 'schema.describe.newColumn'));
    const summary = parts.reduce((first, next) => msg('schema.describe.next', { first, next }));
    return f.description ? msg('schema.describe.withDescription', { summary, description: this.localized(f.description) }) : summary;
  }

  /** A bound as people read it: dates as written, percents as percentages. */
  private bound(f: FieldSchema | CellSchema, b: number | string): Text {
    if (typeof b === 'string') return b;
    if (f.type === 'percent') return msg('schema.bound.percent', { n: +(b * 100).toPrecision(12) });
    return msg('schema.bound.number', { n: b });
  }

  // ---------------------------------------------------------------------------
  // Typed input

  /** Rewrites typed or pasted text so it is stored as the field's type, or refuses it. */
  prepareInput(sheet: number, row: number, col: number, text: string): { text: string } | { reject: Message } {
    const ref = this.fieldAt(sheet, row, col);
    if (!ref || text === '') return { text };
    const f = ref.field;
    if (text.startsWith('=') && text.length > 1) {
      return f.formula === 'forbid' ? { reject: msg('schema.input.formulaForbidden', { field: ref.name }) } : { text };
    }
    const t = text.trim();
    switch (f.type) {
      case 'text':
        return { text: text.startsWith("'") ? text : "'" + text };
      case 'enum': {
        const match = f.values?.find((v) => v.toLowerCase() === t.toLowerCase());
        return { text: "'" + (match ?? text) };
      }
      case 'boolean':
        if (/^(true|yes|y|1)$/i.test(t)) return { text: 'TRUE' };
        if (/^(false|no|n|0)$/i.test(t)) return { text: 'FALSE' };
        return { text: "'" + text };
      case 'percent':
        // Like Excel in a percentage cell: typing 12 means 12%.
        return { text: /^-?[\d,]*\.?\d+$/.test(t) ? `${t}%` : text };
      default:
        return { text };
    }
  }

  // ---------------------------------------------------------------------------
  // Values

  private toValue(rec: CellRecord | null, field?: FieldSchema | CellSchema): FieldValue {
    if (!rec) return null;
    switch (rec.t) {
      case VT.Number: {
        const n = rec.n ?? 0;
        const style = this.host.styles()[rec.style];
        if (field?.type === 'date' || (field === undefined || field.type === 'text' ? style && isDateStyle(style) : false)) {
          return new Date(EPOCH + Math.round((n + (this.host.date1904() ? 1462 : 0)) * DAY));
        }
        return n;
      }
      case VT.String:
        return rec.s ?? '';
      case VT.Bool:
        return rec.n === 1;
      case VT.Error:
        return { error: rec.s ?? '#VALUE!' };
      default:
        return null;
    }
  }

  /** A cell's value as schema code sees it, typed by its field when it has one. */
  valueAt(sheet: number, row: number, col: number): FieldValue {
    const view = this.host.views()[sheet];
    return view ? this.toValue(view.get(row, col), this.fieldAt(sheet, row, col)?.field) : null;
  }

  private recordValues(t: BoundTable, record: number): Record<string, FieldValue> {
    const out: Record<string, FieldValue> = {};
    t.fields.forEach((f, j) => {
      const p = t.cell(record, j);
      put(out, f.key, this.toValue(t.view.get(p.row, p.col), f));
    });
    return out;
  }

  private tableData(bs: BoundSheet, t: BoundTable): TableData {
    const self = this;
    return {
      key: t.schema.key,
      sheet: bs.name,
      fields: t.fields,
      get records() {
        return Array.from({ length: t.recordCount() }, (_, i) => ({
          index: i,
          label: t.labels ? cellText(t.view, t.labelCell(i).row, t.labelCell(i).col) : undefined,
          values: self.recordValues(t, i),
          ref: (field: string) => {
            const j = Math.max(0, t.fields.findIndex((f) => f.key === field));
            const p = t.cell(i, j);
            return toA1(p.row, p.col);
          },
        }));
      },
    };
  }

  private sheetData(bs: BoundSheet): SheetData {
    return {
      name: bs.name,
      table: (key) => {
        const t = bs.tables.find((x) => x.schema.key === key);
        return t ? this.tableData(bs, t) : undefined;
      },
      cell: (key) => {
        const c = bs.cells.find((x) => x.schema.key === key);
        return c ? this.toValue(bs.view.get(c.pos.row, c.pos.col), c.schema) : null;
      },
    };
  }

  /** Every table's records and every single cell, as plain JSON, for handing to a host on save. */
  values(): SchemaValues {
    const out: SchemaValues = {};
    for (const bs of this.binding.sheets) {
      const sheet: SchemaValues[string] = { tables: {}, cells: {} };
      for (const t of bs.tables) {
        put(
          sheet.tables,
          t.schema.key,
          Array.from({ length: t.recordCount() }, (_, r) => {
            const values = this.recordValues(t, r);
            return Object.fromEntries(Object.entries(values).map(([k, v]) => [k, toJson(v)]));
          }),
        );
      }
      for (const c of bs.cells) put(sheet.cells, c.schema.key, toJson(this.toValue(bs.view.get(c.pos.row, c.pos.col), c.schema)));
      put(out, bs.name, sheet);
    }
    return out;
  }

  data(signal: AbortSignal): WorkbookData {
    return {
      signal,
      language: this.language,
      sheet: (name) => {
        const bs = this.binding.sheets.find((b) => b.name.toLowerCase() === name.toLowerCase() || textMatches(b.schema.name, name));
        return bs ? this.sheetData(bs) : undefined;
      },
    };
  }

  // ---------------------------------------------------------------------------
  // Validation

  private checkValue(f: FieldSchema | CellSchema, rec: CellRecord | null, value: FieldValue): Lite[] {
    const out: Lite[] = [];
    const err = (key: MessageKey, params?: Record<string, unknown>) => out.push({ severity: 'error', message: msg(key, params) });
    const isFormula = rec?.f !== undefined;
    if (f.formula === 'require' && !isFormula) {
      err('schema.check.formulaRequired');
      return out;
    }
    if (f.formula === 'forbid' && isFormula) err('schema.check.formulaForbidden');
    if (rec && rec.flags & CellFlag.Stale) {
      out.push({ severity: 'warning', message: msg('schema.check.stale') });
      return out;
    }
    if (value === null || value === '') {
      if (f.required) err('schema.check.required');
      return out;
    }
    if (typeof value === 'object' && !(value instanceof Date)) {
      err('schema.check.formulaError', { error: value.error });
      return out;
    }
    switch (f.type) {
      case 'number':
      case 'integer':
      case 'percent': {
        if (typeof value !== 'number') {
          err(f.type === 'percent' ? 'schema.check.expectedPercent' : f.type === 'integer' ? 'schema.check.expectedInteger' : 'schema.check.expectedNumber');
          break;
        }
        if (f.type === 'integer' && !Number.isInteger(value)) err('schema.check.expectedInteger');
        if (f.decimals !== undefined && decimalsOf(f.type === 'percent' ? value * 100 : value) > f.decimals) {
          if (f.decimals === 0) err('schema.check.noDecimals');
          else err('schema.check.maxDecimals', { count: f.decimals });
        }
        if (typeof f.min === 'number' && value < f.min) err('schema.check.min', { bound: this.bound(f, f.min) });
        if (typeof f.max === 'number' && value > f.max) err('schema.check.max', { bound: this.bound(f, f.max) });
        break;
      }
      case 'date':
        if (!(value instanceof Date)) {
          err('schema.check.expectedDate');
          break;
        }
        if (typeof f.min === 'string' && isoDate(value) < f.min) err('schema.check.onOrAfter', { date: f.min });
        if (typeof f.max === 'string' && isoDate(value) > f.max) err('schema.check.onOrBefore', { date: f.max });
        break;
      case 'boolean':
        if (typeof value !== 'boolean') err('schema.check.expectedBoolean');
        break;
      case 'enum':
        // Joined plainly rather than as a sentence list: these are the values to type.
        if (!f.values?.includes(String(value))) err('schema.check.oneOf', { values: f.values?.join(', ') ?? '' });
        break;
      case 'text':
        break;
    }
    if (f.minLength !== undefined || f.maxLength !== undefined) {
      const length = (value instanceof Date ? isoDate(value) : String(value)).length;
      if (f.minLength !== undefined && length < f.minLength) err('schema.check.minLength', { count: f.minLength });
      if (f.maxLength !== undefined && length > f.maxLength) err('schema.check.maxLength', { count: f.maxLength });
    }
    if (f.pattern && !regex(f.pattern).test(value instanceof Date ? isoDate(value) : String(value))) {
      if (f.patternHint) err('schema.check.patternHint', { hint: this.localized(f.patternHint) });
      else err('schema.check.pattern');
    }
    return out;
  }

  /** Runs a field's own check hook, reusing async results until the value changes. */
  private fieldHook(id: string, f: FieldSchema | CellSchema, value: FieldValue, ctx: Omit<FieldContext, 'signal' | 'language'>): Lite | null {
    const key = valueKey(value);
    const existing = this.fieldChecks.get(id);
    if (existing && existing.key === key && (existing.pending || existing.issue !== undefined)) return null;
    existing?.abort.abort();
    const abort = new AbortController();
    let result: CheckResult | Promise<CheckResult>;
    try {
      result = f.validate!(value, { ...ctx, language: this.language, signal: abort.signal });
    } catch (e) {
      this.fieldChecks.delete(id);
      return { severity: 'error', message: msg('schema.hook.failed', { error: failure(e) }) };
    }
    if (!isPromise<CheckResult>(result)) {
      this.fieldChecks.delete(id);
      return toLite(result);
    }
    const entry: { key: string; issue: Lite | null; pending?: Promise<void>; abort: AbortController } = { key, issue: null, abort };
    entry.pending = withTimeout(result, this.timeoutMs, () => abort.abort()).then(
      (r) => {
        entry.issue = toLite(r);
      },
      (e: Error) => {
        if (e instanceof HookTimeout) entry.issue = { severity: 'warning', message: msg('schema.hook.fieldTimeout', { duration: e.duration }) };
        else if (!abort.signal.aborted) entry.issue = { severity: 'error', message: msg('schema.hook.failed', { error: failure(e) }) };
      },
    ).finally(() => {
      entry.pending = undefined;
      if (this.fieldChecks.get(id) === entry) {
        this.reindex();
        this.host.changed();
      }
    });
    this.fieldChecks.set(id, entry);
    return null;
  }

  private resolveIssues(inputs: Found[], source: Issue['source'], bs?: BoundSheet, t?: BoundTable): Issue[] {
    return inputs.map((i) => {
      let sheet = bs;
      let pos: Pos | undefined;
      if (i.ref) {
        const [sheetName, cell] = i.ref.includes('!') ? i.ref.split('!') : [undefined, i.ref];
        if (sheetName) sheet = this.binding.sheets.find((b) => b.name.toLowerCase() === sheetName.replace(/^'|'$/g, '').toLowerCase()) ?? sheet;
        pos = parseA1(cell.split(':')[0]) ?? undefined;
      } else if (t && (i.record !== undefined || i.field !== undefined)) {
        const j = i.field !== undefined ? Math.max(0, t.fields.findIndex((f) => f.key === i.field)) : 0;
        pos = i.record !== undefined ? t.cell(i.record, j) : t.headerCell(j);
      }
      return {
        severity: i.severity ?? 'error',
        message: i.message,
        sheet: sheet?.sheet ?? -1,
        sheetName: sheet?.name ?? '',
        row: pos?.row,
        col: pos?.col,
        source,
      };
    });
  }

  /** Runs a hook that may be async; results from superseded runs are dropped. */
  private runHook(id: string, generation: number, source: Issue['source'], run: () => IssueInput[] | void | Promise<IssueInput[] | void>, bs?: BoundSheet, t?: BoundTable): void {
    const failed = (e: unknown): Issue[] => this.resolveIssues([{ message: msg('schema.hook.failed', { error: failure(e) }) }], source, bs, t);
    let result: IssueInput[] | void | Promise<IssueInput[] | void>;
    try {
      result = run();
    } catch (e) {
      this.hookResults.set(id, failed(e));
      return;
    }
    if (!isPromise<IssueInput[] | void>(result)) {
      this.hookResults.set(id, this.resolveIssues(result ?? [], source, bs, t));
      return;
    }
    const signal = this.hookAbort.signal;
    const p: Promise<void> = withTimeout(result, this.timeoutMs)
      .then(
        (r) => {
          if (generation === this.generation) this.hookResults.set(id, this.resolveIssues(r ?? [], source, bs, t));
        },
        (e) => {
          if (generation !== this.generation) return;
          if (e instanceof HookTimeout) this.hookResults.set(id, this.resolveIssues([{ severity: 'warning', message: msg(TIMEOUTS[source], { duration: e.duration }) }], source, bs, t));
          else if (!signal.aborted) this.hookResults.set(id, failed(e));
        },
      )
      .finally(() => {
        if (this.pendingHooks.get(id) === p) this.pendingHooks.delete(id);
        this.reindex();
        this.host.changed();
      });
    this.pendingHooks.set(id, p);
  }

  private runCheck(t: BoundTable, bs: BoundSheet, c: DeclarativeRule, count: number): Issue[] {
    if (c.rule === 'compare' || c.rule === 'requiredIf' || c.rule === 'uniqueTogether') return this.resolveIssues(this.recordCheck(t, c, count), 'table', bs, t);
    const fields = c.field === '*' ? t.fields.map((_, j) => j) : [t.fields.findIndex((f) => f.key === c.field)].filter((j) => j >= 0);
    const inputs: Found[] = [];
    const num = (r: number, j: number) => {
      const v = this.recordValues(t, r)[t.fields[j].key];
      return typeof v === 'number' ? v : v instanceof Date ? v.getTime() : null;
    };
    for (const j of fields) {
      const key = t.fields[j].key;
      if (c.rule === 'monotonic') {
        const ok = { increasing: (a: number, b: number) => b > a, decreasing: (a: number, b: number) => b < a, nonIncreasing: (a: number, b: number) => b <= a, nonDecreasing: (a: number, b: number) => b >= a }[c.direction];
        const message = msg(MONOTONIC[c.direction]);
        let prev: number | null = null;
        for (let r = 0; r < count; r++) {
          const v = num(r, j);
          if (v === null) continue;
          if (prev !== null && !ok(prev, v)) inputs.push({ record: r, field: key, severity: c.severity, message });
          prev = v;
        }
      } else if (c.rule === 'firstEquals') {
        if (count === 0) continue;
        const v = this.recordValues(t, 0)[key];
        const same = typeof c.value === 'number' ? v === c.value : String(v ?? '').toLowerCase() === c.value.toLowerCase();
        if (!same) inputs.push({ record: 0, field: key, severity: c.severity, message: msg('schema.check.firstEquals', { value: c.value }) });
      } else if (c.rule === 'sum') {
        let total = 0;
        for (let r = 0; r < count; r++) total += num(r, j) ?? 0;
        total = parseFloat(total.toPrecision(15));
        const field = matchLabel(t.fields[j].label);
        if (c.equals !== undefined && total !== c.equals) inputs.push({ field: key, severity: c.severity, message: msg('schema.check.sumEquals', { field, total, expected: c.equals }) });
        if (c.max !== undefined && total > c.max) inputs.push({ field: key, severity: c.severity, message: msg('schema.check.sumMax', { field, total, max: c.max }) });
        if (c.min !== undefined && total < c.min) inputs.push({ field: key, severity: c.severity, message: msg('schema.check.sumMin', { field, total, min: c.min }) });
      }
    }
    return this.resolveIssues(inputs, 'table', bs, t);
  }

  /** Checks that compare fields within each record, or records with each other. */
  private recordCheck(t: BoundTable, c: Extract<DeclarativeRule, { rule: 'compare' | 'requiredIf' | 'uniqueTogether' }>, count: number): Found[] {
    const out: Found[] = [];
    const label = (key: string): Text => {
      const f = t.fields.find((x) => x.key === key);
      return f ? quoted(matchLabel(f.label)) : key;
    };
    const blank = (v: FieldValue) => v === null || v === '';
    const comparable = (v: FieldValue) => (v instanceof Date ? v.getTime() : typeof v === 'string' ? v.toLowerCase() : v);
    const seen = new Map<string, number>();
    for (let r = 0; r < count; r++) {
      const values = this.recordValues(t, r);
      if (c.rule === 'compare') {
        const a = values[c.field];
        const b = values[c.other];
        // Like blanks, a formula error has nothing to compare; its own cell already reports it.
        const isError = (v: unknown) => typeof v === 'object' && v !== null && !(v instanceof Date) && 'error' in v;
        if (blank(a) || blank(b) || isError(a) || isError(b)) continue;
        const x = comparable(a);
        const y = comparable(b);
        const ok = { '<': x! < y!, '<=': x! <= y!, '>': x! > y!, '>=': x! >= y!, '=': x === y, '<>': x !== y }[c.op];
        if (!ok) out.push({ record: r, field: c.field, severity: c.severity, message: msg((a instanceof Date ? COMPARE_DATES : COMPARE)[c.op], { other: label(c.other) }) });
      } else if (c.rule === 'requiredIf') {
        const w = values[c.when];
        const applies = c.equals === undefined ? !blank(w) : comparable(w) === comparable(c.equals as FieldValue);
        if (applies && blank(values[c.field])) {
          const message =
            c.equals === undefined
              ? msg('schema.check.requiredWhenFilled', { field: label(c.when) })
              : msg('schema.check.requiredWhenEquals', { field: label(c.when), value: String(c.equals) });
          out.push({ record: r, field: c.field, severity: c.severity, message });
        }
      } else {
        const parts = c.fields.map((k) => values[k]);
        if (parts.every(blank)) continue;
        const key = parts.map((v) => valueKey(v).toLowerCase()).join('\u0000');
        const first = seen.get(key);
        if (first === undefined) seen.set(key, r);
        else out.push({ record: r, field: c.field, severity: c.severity, message: msg('schema.check.sameAsRecord', { fields: and(c.fields.map(label)), record: first + 1 }) });
      }
    }
    return out;
  }

  /** Checks the whole workbook against the schema. Async hooks report back when they finish. */
  validate(): void {
    const generation = ++this.generation;
    this.hookAbort.abort();
    this.hookAbort = new AbortController();
    const data = this.data(this.hookAbort.signal);
    const issues: Issue[] = [];
    const seenChecks = new Set<string>();
    const push = (bs: BoundSheet, p: Pos, lite: Lite | null) => {
      if (lite) issues.push({ ...lite, sheet: bs.sheet, sheetName: bs.name, row: p.row, col: p.col, source: 'field' });
    };

    for (const bs of this.binding.sheets) {
      for (const t of bs.tables) {
        const count = t.recordCount();
        this.counts.set(t, count);
        const { min, max } = t.schema.records ?? {};
        if (!t.fixed && min !== undefined && count < min) {
          issues.push(...this.resolveIssues([{ record: count, message: msg(t.byRow ? 'schema.table.minRows' : 'schema.table.minColumns', { count: min }) }], 'table', bs, t));
        }
        if (!t.fixed && max !== undefined && count > max) {
          issues.push(...this.resolveIssues([{ record: max, message: msg(t.byRow ? 'schema.table.maxRows' : 'schema.table.maxColumns', { count: max }) }], 'table', bs, t));
        }
        const seen = t.fields.map((f) => (f.unique ? new Map<string, number>() : null));
        for (let r = 0; r < count; r++) {
          const values = this.recordValues(t, r);
          t.fields.forEach((f, j) => {
            const p = t.cell(r, j);
            const rec = t.view.get(p.row, p.col);
            const value = values[f.key];
            for (const lite of this.checkValue(f, rec, value)) push(bs, p, lite);
            const dupes = seen[j];
            if (dupes && value !== null && value !== '') {
              const k = valueKey(value).toLowerCase();
              const first = dupes.get(k);
              if (first !== undefined) {
                const fp = t.cell(first, j);
                push(bs, p, { severity: 'error', message: msg('schema.check.duplicate', { ref: toA1(fp.row, fp.col), field: matchLabel(f.label) }) });
              } else {
                dupes.set(k, r);
              }
            }
            if (f.validate && rec?.flags !== undefined && !(rec.flags & CellFlag.Stale) && value !== null) {
              const id = `${bs.sheet}:${p.row}:${p.col}`;
              seenChecks.add(id);
              push(bs, p, this.fieldHook(id, f, value, { sheet: bs.name, table: t.schema.key, record: r, values }));
            }
          });
        }
        for (const c of t.schema.checks ?? []) issues.push(...this.runCheck(t, bs, c, count));
        t.schema.rules?.forEach((rule, i) => {
          if (typeof rule === 'function') this.runHook(`table:${bs.sheet}:${t.schema.key}:${i}`, generation, 'table', () => rule(this.tableData(bs, t), data), bs, t);
        });
      }
      for (const c of bs.cells) {
        const rec = bs.view.get(c.pos.row, c.pos.col);
        const value = this.toValue(rec, c.schema);
        for (const lite of this.checkValue(c.schema, rec, value)) push(bs, c.pos, lite);
        if (c.schema.validate && value !== null) {
          const id = `${bs.sheet}:${c.pos.row}:${c.pos.col}`;
          seenChecks.add(id);
          push(bs, c.pos, this.fieldHook(id, c.schema, value, { sheet: bs.name }));
        }
      }
      if (bs.schema.validate) {
        this.runHook(`sheet:${bs.sheet}`, generation, 'sheet', () => bs.schema.validate!(this.sheetData(bs), data), bs);
      }
    }
    if (this.schema.validate) this.runHook('workbook', generation, 'workbook', () => this.schema.validate!(data));

    for (const [id, entry] of this.fieldChecks) {
      if (!seenChecks.has(id)) {
        entry.abort.abort();
        this.fieldChecks.delete(id);
      }
    }
    this.syncIssues = issues;
    this.reindex();
    this.host.changed();
  }

  private reindex(): void {
    const index = new Map<string, Issue[]>();
    for (const i of this.issues()) {
      if (i.row === undefined || i.col === undefined) continue;
      const k = `${i.sheet}:${i.row}:${i.col}`;
      const list = index.get(k);
      if (list) list.push(i);
      else index.set(k, [i]);
    }
    this.index = index;
  }

  /** Every current issue, errors first. */
  issues(): Issue[] {
    const out = [...this.syncIssues];
    for (const list of this.hookResults.values()) out.push(...list);
    for (const [id, e] of this.fieldChecks) {
      if (!e.issue || e.pending) continue;
      const [sheet, row, col] = id.split(':').map(Number);
      out.push({ ...e.issue, sheet, sheetName: this.boundSheet(sheet)?.name ?? '', row, col, source: 'field' });
    }
    return out.sort((a, b) => (a.severity === b.severity ? a.sheet - b.sheet || (a.row ?? -1) - (b.row ?? -1) || (a.col ?? -1) - (b.col ?? -1) : a.severity === 'error' ? -1 : 1));
  }

  cellIssues(sheet: number, row: number, col: number): Issue[] {
    return this.index.get(`${sheet}:${row}:${col}`) ?? [];
  }

  isPending(sheet: number, row: number, col: number): boolean {
    return !!this.fieldChecks.get(`${sheet}:${row}:${col}`)?.pending;
  }

  summary(): { errors: number; warnings: number; pending: number } {
    const all = this.issues();
    let pending = this.pendingHooks.size;
    for (const e of this.fieldChecks.values()) if (e.pending) pending++;
    return { errors: all.filter((i) => i.severity === 'error').length, warnings: all.filter((i) => i.severity === 'warning').length, pending };
  }

  /** Resolves once no checks are running, or after the timeout. */
  async settled(timeoutMs = 15_000): Promise<void> {
    const deadline = Date.now() + timeoutMs;
    for (;;) {
      const waiting: Promise<void>[] = [...this.pendingHooks.values()];
      for (const e of this.fieldChecks.values()) if (e.pending) waiting.push(e.pending);
      if (!waiting.length || Date.now() > deadline) return;
      await Promise.race([Promise.allSettled(waiting), new Promise((r) => setTimeout(r, Math.max(0, deadline - Date.now())))]);
    }
  }

  // ---------------------------------------------------------------------------
  // Reacting to edits

  /**
   * After the user edits cells: re-validate, carry formulas into newly
   * started records, then run the schema's onChange hook. `edited` holds
   * each cell's value from before the edit.
   */
  async afterEdit(edited: { sheet: number; row: number; col: number; before: FieldValue }[]): Promise<void> {
    // Record counts from the last validation are the "before" state; read
    // them before validating again.
    const fills: EditOp[] = [];
    for (const bs of this.binding.sheets) {
      for (const t of bs.tables) {
        if (t.fixed) continue;
        const before = this.counts.get(t) ?? 0;
        const now = t.recordCount();
        if (now <= before || before === 0) continue;
        // Like an Excel table: new records take the calculated fields and
        // number formats of the record before them.
        for (let r = before; r < now; r++) {
          const restyle: { sheet: number; row: number; col: number; content: CellContent; style: number }[] = [];
          t.fields.forEach((f, j) => {
            const p = t.cell(r, j);
            const s = t.cell(r - 1, j);
            const rec = t.view.get(p.row, p.col);
            if (f.formula === 'require' && rec?.f === undefined) {
              fills.push({ type: 'copy', srcSheet: t.sheet, src: { r1: s.row, c1: s.col, r2: s.row, c2: s.col }, sheet: t.sheet, dst: { r1: p.row, c1: p.col, r2: p.row, c2: p.col } });
              return;
            }
            const prevStyle = t.view.get(s.row, s.col)?.style ?? 0;
            if ((rec?.style ?? 0) === 0 && prevStyle !== 0) restyle.push({ sheet: t.sheet, row: p.row, col: p.col, content: contentOf(rec), style: prevStyle });
          });
          if (restyle.length) fills.push({ type: 'restore', cells: restyle });
        }
      }
    }
    this.validate();
    if (fills.length) await this.host.followUp({ type: 'batch', ops: fills });

    const onChange = this.schema.onChange;
    if (!onChange) return;
    const changes = edited
      .filter((e) => this.boundSheet(e.sheet))
      .map((e) => {
        const ref = this.fieldAt(e.sheet, e.row, e.col);
        return {
          sheet: this.boundSheet(e.sheet)!.name,
          table: ref?.table?.schema.key,
          record: ref?.record,
          field: ref?.field.key,
          ref: toA1(e.row, e.col),
          before: e.before,
          after: this.valueAt(e.sheet, e.row, e.col),
        };
      });
    if (!changes.length) return;
    const writes = new Map<number, { row: number; col: number; text: string }[]>();
    const abort = new AbortController();
    const ctx: ChangeContext = {
      ...this.data(abort.signal),
      changes,
      set: (target: SetTarget, value: FieldValue) => {
        const bs = this.binding.sheets.find((b) => b.name.toLowerCase() === target.sheet.toLowerCase());
        if (!bs) return;
        let pos: Pos | null = null;
        if ('ref' in target) pos = parseA1(target.ref);
        else {
          const t = bs.tables.find((x) => x.schema.key === target.table);
          const j = t ? t.fields.findIndex((f) => f.key === target.field) : -1;
          if (t && j >= 0) pos = t.cell(target.record, j);
        }
        if (!pos) return;
        let text: string;
        if (value === null) text = '';
        else if (value instanceof Date) text = isoDate(value);
        else if (typeof value === 'boolean') text = value ? 'TRUE' : 'FALSE';
        else if (typeof value === 'object') text = value.error;
        else text = String(value);
        const prepared = typeof value === 'string' ? this.prepareInput(bs.sheet, pos.row, pos.col, text) : { text };
        if ('reject' in prepared) return;
        const list = writes.get(bs.sheet) ?? [];
        list.push({ row: pos.row, col: pos.col, text: prepared.text });
        writes.set(bs.sheet, list);
      },
    };
    try {
      await withTimeout(Promise.resolve(onChange(ctx)), this.timeoutMs, () => abort.abort());
    } catch (e) {
      const message = e instanceof HookTimeout ? msg('schema.hook.onChangeTimeout', { duration: e.duration }) : msg('schema.hook.onChangeFailed', { error: failure(e) });
      this.hookResults.set('onChange', [{ severity: 'error', message, sheet: -1, sheetName: '', source: 'workbook' }]);
      this.reindex();
      this.host.changed();
      return;
    }
    this.hookResults.delete('onChange');
    if (writes.size) {
      await this.host.followUp({ type: 'batch', ops: [...writes].map(([sheet, cells]) => ({ type: 'inputs' as const, sheet, cells })) });
    }
  }

  // ---------------------------------------------------------------------------
  // Adding and removing records

  /** The table and record at a cell, when records there can be inserted or deleted. */
  recordAt(sheet: number, row: number, col: number): { table: BoundTable; record: number; append: boolean } | undefined {
    const ref = this.fieldAt(sheet, row, col);
    if (!ref?.table || ref.record === undefined || ref.table.fixed) return undefined;
    return { table: ref.table, record: ref.record, append: !!ref.append };
  }

  /** Plans inserting an empty record before `at`; returns a reason when it can't. */
  insertRecord(t: BoundTable, at: number): EditOp | Message {
    if (t.fixed) return msg(t.byRow ? 'schema.table.rowsFixed' : 'schema.table.columnsFixed');
    const count = t.recordCount();
    const max = t.schema.records?.max;
    if (max !== undefined && count >= max) return msg(t.byRow ? 'schema.table.maxRows' : 'schema.table.maxColumns', { count: max });
    const pos = Math.min(at, count);
    const ops: EditOp[] = [];
    if (pos < count) {
      ops.push({ type: 'copy', srcSheet: t.sheet, src: t.recordsRect(pos, count - 1), sheet: t.sheet, dst: t.recordsRect(pos + 1, count) });
      ops.push({ type: 'clear', sheet: t.sheet, rects: [t.recordsRect(pos, pos)] });
    }
    // Calculated fields carry over from the neighbouring record, as in an Excel table.
    const neighbour = pos < count ? pos + 1 : pos - 1;
    if (neighbour >= 0) {
      t.fields.forEach((f, j) => {
        if (f.formula !== 'require') return;
        const s = t.cell(neighbour, j);
        const d = t.cell(pos, j);
        ops.push({ type: 'copy', srcSheet: t.sheet, src: { r1: s.row, c1: s.col, r2: s.row, c2: s.col }, sheet: t.sheet, dst: { r1: d.row, c1: d.col, r2: d.row, c2: d.col } });
      });
    }
    return ops.length ? { type: 'batch', ops } : msg('schema.table.nothingToInsert');
  }

  /** Plans deleting records first..last, closing the gap; returns a reason when it can't. */
  deleteRecords(t: BoundTable, first: number, last: number): EditOp | Message {
    if (t.fixed) return msg(t.byRow ? 'schema.table.rowsFixed' : 'schema.table.columnsFixed');
    const count = t.recordCount();
    const a = Math.max(0, first);
    const b = Math.min(last, count - 1);
    if (b < a) return msg(t.byRow ? 'schema.table.selectRowToDelete' : 'schema.table.selectColumnToDelete');
    const n = b - a + 1;
    const min = t.schema.records?.min;
    if (min !== undefined && count - n < min) return msg(t.byRow ? 'schema.table.minRows' : 'schema.table.minColumns', { count: min });
    const ops: EditOp[] = [];
    if (b + 1 <= count - 1) {
      ops.push({ type: 'copy', srcSheet: t.sheet, src: t.recordsRect(b + 1, count - 1), sheet: t.sheet, dst: t.recordsRect(a, a + (count - 1 - (b + 1))) });
    }
    ops.push({ type: 'clear', sheet: t.sheet, rects: [t.recordsRect(count - n, count - 1)] });
    return { type: 'batch', ops };
  }

  // ---------------------------------------------------------------------------
  // Saving

  async checkBeforeSave(): Promise<{ ok: true } | { ok: false; message: Text; canOverride: boolean }> {
    await this.settled();
    const { errors } = this.summary();
    if (errors) {
      return {
        ok: false,
        message: msg('schema.save.errors', { count: errors, schema: this.localized(this.schema.name) }),
        canOverride: this.schema.saveWithErrors === 'confirm',
      };
    }
    if (this.schema.beforeSave) {
      try {
        const abort = new AbortController();
        const r = await withTimeout(Promise.resolve(this.schema.beforeSave(this.data(abort.signal))), this.timeoutMs, () => abort.abort());
        // The host's own reason, in one language or several.
        if (typeof r === 'string' || (typeof r === 'object' && r !== null)) return { ok: false, message: r, canOverride: false };
      } catch (e) {
        if (e instanceof HookTimeout) return { ok: false, message: msg('schema.save.timeout', { duration: e.duration }), canOverride: false };
        return { ok: false, message: msg('schema.save.failed', { error: failure(e) }), canOverride: false };
      }
    }
    return { ok: true };
  }
}
