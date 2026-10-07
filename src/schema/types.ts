/**
 * Schemas let a consumer of the editor describe the workbooks it expects:
 * which sheets exist, where tables sit, what each field holds, and the
 * business rules that apply. A workbook that matches a schema is validated
 * and constrained by it; one that matches none opens as a plain spreadsheet.
 *
 * Everything except the hook functions is plain JSON, so schemas can be
 * generated from a workbook, edited in the designer, exported and imported.
 * Hooks (custom checks, business logic, async calls) are added in code.
 */

/** A cell's value as schema code sees it. Dates arrive as Date (UTC midnight). */
export type FieldValue = string | number | boolean | Date | null | { error: string };

export type FieldType = 'text' | 'number' | 'integer' | 'percent' | 'date' | 'boolean' | 'enum';

export type Severity = 'error' | 'warning';

/**
 * Expected text. A plain string matches ignoring case and surrounding
 * spaces; a pattern is a regular expression.
 */
export type TextMatch = string | RegExp | { pattern: string; flags?: string };

/** What a check returns: nothing when the value is fine, otherwise a message. */
export type CheckResult = string | { message: string; severity?: Severity } | null | undefined | void;

export interface FieldContext {
  sheet: string;
  table?: string;
  /** Record index within the table (0-based), when the field belongs to a table. */
  record?: number;
  /** The record's values by field key, when the field belongs to a table. */
  values?: Record<string, FieldValue>;
  /** Aborted when the value changes again before this check finishes. */
  signal: AbortSignal;
}

export interface FieldSchema {
  /** Stable name used in code and rules. */
  key: string;
  /** Header text in the workbook. */
  label: TextMatch;
  type: FieldType;
  required?: boolean;
  /** Inclusive bounds: numbers (percents as fractions) or ISO dates for date fields. */
  min?: number | string;
  max?: number | string;
  /** Most decimal places allowed; for percents, of the percentage (12.5% has 1). */
  decimals?: number;
  /** Regular expression the text form of the value must match. */
  pattern?: string;
  /** How to describe the pattern to people, e.g. "like SKU-1234". */
  patternHint?: string;
  /** Allowed values for "enum" fields; typed input is matched ignoring case. */
  values?: string[];
  /** No two records may share a value. */
  unique?: boolean;
  readOnly?: boolean;
  /** Whether cells may, must or must not hold formulas. Default "allow". */
  formula?: 'allow' | 'forbid' | 'require';
  /** Shown in the formula bar when a cell of this field is selected. */
  description?: string;
  /** Code hook: extra check, sync or async. Async checks re-run only when the value changes. */
  validate?: (value: FieldValue, ctx: FieldContext) => CheckResult | Promise<CheckResult>;
}

export interface IssueInput {
  message: string;
  severity?: Severity;
  /** Pin the issue to a record and/or field of the table it came from. */
  record?: number;
  field?: string;
  /** Or to a cell on the sheet, e.g. "B4". */
  ref?: string;
}

export interface RecordData {
  index: number;
  /** The fixed record label, for tables with record labels. */
  label?: string;
  values: Record<string, FieldValue>;
  /** Address of one of this record's cells, e.g. "D7". */
  ref(field: string): string;
}

export interface TableData {
  key: string;
  sheet: string;
  records: RecordData[];
  fields: FieldSchema[];
}

export interface SheetData {
  name: string;
  table(key: string): TableData | undefined;
  /** Value of a single-cell field defined under `cells`. */
  cell(key: string): FieldValue;
}

export interface WorkbookData {
  sheet(name: string): SheetData | undefined;
  /** Aborted when the workbook changes again before this check finishes. */
  signal: AbortSignal;
}

export type TableRule = (table: TableData, ctx: WorkbookData) => IssueInput[] | void | Promise<IssueInput[] | void>;

/** Built-in table checks that can be expressed as data. */
export type DeclarativeRule =
  | { rule: 'monotonic'; field: string; direction: 'increasing' | 'decreasing' | 'nonIncreasing' | 'nonDecreasing'; severity?: Severity }
  | { rule: 'firstEquals'; field: string; value: number | string; severity?: Severity }
  | { rule: 'sum'; field: string; equals?: number; max?: number; min?: number; severity?: Severity };

export interface TableSchema {
  key: string;
  /** Top-left cell of the table, including any label row or column. */
  at: string;
  /**
   * "rows": one record per row, fields across the columns (the usual table).
   * "columns": one record per column, fields down the rows.
   */
  orientation?: 'rows' | 'columns';
  /** The first row ("rows") or column ("columns") holds the field labels. Default true. */
  header?: boolean;
  fields: FieldSchema[];
  records?: {
    /**
     * Fixed labels in the first column ("rows") or row ("columns"). When
     * given, the number of records is fixed and the labels can't be edited.
     */
    labels?: TextMatch[];
    /** Bounds on the number of records when they can be added and removed. */
    min?: number;
    max?: number;
  };
  /** Text expected in the top-left cell when the table has both kinds of labels. */
  corner?: TextMatch;
  /** Declarative checks; "*" as the field applies a check to every field. */
  checks?: DeclarativeRule[];
  /** Code hooks: table-wide checks, sync or async. */
  rules?: TableRule[];
}

export interface CellSchema extends Omit<FieldSchema, 'label'> {
  /** Where the value lives, e.g. "B8". */
  ref: string;
  /** A fixed label that must sit in another cell, e.g. { ref: "A8", text: "Approved by" }. */
  label?: { ref: string; text: TextMatch };
}

export interface SheetSchema {
  name: TextMatch;
  optional?: boolean;
  tables?: TableSchema[];
  cells?: CellSchema[];
  /** Whether cells outside the tables and cells above can be edited. Default "readOnly". */
  outside?: 'readOnly' | 'editable';
  /** Code hook: sheet-wide checks, sync or async. */
  validate?: (sheet: SheetData, ctx: WorkbookData) => IssueInput[] | void | Promise<IssueInput[] | void>;
}

/** How numbers are shown and copied. */
export interface NumberOptions {
  /** "full" shows every stored digit instead of the cell's rounded format. */
  display?: 'formatted' | 'full';
  /** "raw" puts exact values on the clipboard instead of the displayed text. */
  copy?: 'formatted' | 'raw';
}

export interface Change {
  sheet: string;
  table?: string;
  record?: number;
  field?: string;
  ref: string;
  before: FieldValue;
  after: FieldValue;
}

export type SetTarget = { sheet: string; table: string; record: number; field: string } | { sheet: string; ref: string };

export interface ChangeContext extends WorkbookData {
  changes: Change[];
  /** Write a value into a cell; applied, as part of the same undo step, when the hook finishes. */
  set(target: SetTarget, value: FieldValue): void;
}

export interface WorkbookSchema {
  id: string;
  name: string;
  description?: string;
  sheets: SheetSchema[];
  /** Whether sheets the schema doesn't describe can be edited. Default "editable". */
  otherSheets?: 'editable' | 'readOnly';
  numbers?: NumberOptions;
  /** "block" refuses to save while errors remain; "confirm" asks first. Default "block". */
  saveWithErrors?: 'block' | 'confirm';
  /** Code hook: workbook-wide business rules; may call services. */
  validate?: (ctx: WorkbookData) => IssueInput[] | void | Promise<IssueInput[] | void>;
  /** Code hook: business logic run after each edit; may fill other cells via ctx.set. */
  onChange?: (ctx: ChangeContext) => void | Promise<void>;
  /** Code hook: last check before saving. Return a message to stop the save. */
  beforeSave?: (ctx: WorkbookData) => true | string | void | Promise<true | string | void>;
}

/** Identity helper so schema files get type checking and completion. */
export function defineSchema(schema: WorkbookSchema): WorkbookSchema {
  return schema;
}
