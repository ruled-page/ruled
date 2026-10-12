/**
 * The editor's core without a browser, for tests: open a workbook, check it
 * against schemas (async hooks included), read its values, change cells and
 * save it, with the same rules the editor applies.
 *
 *   const wb = HeadlessWorkbook.open(bytes, { fileName: 'prices.xlsx', schemas });
 *   const issues = await wb.issues();
 *   await wb.set('Prices', 'D4', '12.50');
 *   const out = wb.save();
 */
import { MAX_COLS, MAX_ROWS, parseA1, toA1 } from '@sitepen/formula-engine/refs';
import { english, errorText } from '../src/i18n';
import { WorkbookSession, type OpenOptions } from '../src/core/session';
import type { CellChange, CellUpdate, DisplayStyle, EditOp, EditResult, Finding } from '../src/core/types';
import { bindSchema } from '../src/schema/binding';
import { SchemaController } from '../src/schema/controller';
import { stampedFirst, stampFor, versionNote } from '../src/schema/stamp';
import type { FieldValue, SchemaValues, WorkbookSchema } from '../src/schema/types';
import { workbookReaders, type WorkbookReaders } from '../src/state/readers';
import { SheetView } from '../src/state/sheetView';

export type { CellChange, EditOp, EditResult, Finding } from '../src/core/types';
export type { JsonValue, SchemaValues, WorkbookSchema } from '../src/schema/types';
export { defineSchema } from '../src/schema/types';
export { schemaFromJSON, schemaToJSON } from '../src/schema/json';

export interface OpenWorkbookOptions extends OpenOptions {
  fileName?: string;
  /** For reading typed input such as dates and decimal commas. Default "en-US". */
  locale?: string;
  /** Schemas to try, in order; a file stamped with one tries that one first. */
  schemas?: WorkbookSchema[];
}

/** A schema problem, with where it is. */
export interface WorkbookIssue {
  severity: 'error' | 'warning';
  message: string;
  sheet: string;
  /** The cell, when the issue is about one. */
  ref?: string;
}

/** Results in English, as plain text. */
function inEnglish(r: EditResult): EditResult {
  return {
    ...r,
    skippedReasons: r.skippedReasons.map((x) => english.text(x)),
    warnings: r.warnings.map((x) => english.text(x)),
    ...(r.rejected !== undefined ? { rejected: english.text(r.rejected) } : {}),
  };
}

export class HeadlessWorkbook implements WorkbookReaders {
  readonly fileName: string;
  private readonly session: WorkbookSession;
  private readonly views: SheetView[];
  private readonly styles: DisplayStyle[];
  private readonly date1904: boolean;
  private readonly controller?: SchemaController;
  /** The schema the workbook matched, if any. */
  readonly schema?: WorkbookSchema;
  /** When schemas were given but none matched: the closest one, and why it didn't. */
  readonly mismatch?: { schema?: string; problems: string[] };
  /** Set when the file was last saved under another version of the matched schema. */
  readonly note?: string;

  private constructor(bytes: Uint8Array, options: OpenWorkbookOptions) {
    this.fileName = options.fileName ?? 'workbook.xlsx';
    try {
      this.session = WorkbookSession.open(bytes, this.fileName, options.locale ?? 'en-US', options);
    } catch (e) {
      // Files the editor turns away say why in words people read, here in English.
      throw new Error(english.text(errorText(e)));
    }
    const r = this.session.loadResult();
    this.styles = r.styles;
    this.date1904 = r.date1904;
    this.views = r.sheets.map((info, i) => new SheetView(info, r.grids[i]));
    this.apply(this.session.ensureEngine().updates);

    const schemas = options.schemas ?? [];
    let closest: { schema: string; problems: string[]; found: number } | undefined;
    for (const s of stampedFirst(schemas, r.properties)) {
      const bound = bindSchema(s, this.views);
      if (bound.binding) {
        this.schema = s;
        const note = versionNote(s, r.properties);
        this.note = note === undefined ? undefined : english.text(note);
        this.controller = new SchemaController(bound.binding, {
          views: () => this.views,
          styles: () => this.styles,
          date1904: () => this.date1904,
          // Business logic writes go in as further edits.
          followUp: async (op) => {
            this.apply(this.session.apply(op).updates);
            this.controller?.validate();
          },
          changed: () => {},
        });
        this.controller.validate();
        return;
      }
      if (!closest || bound.found > closest.found) closest = { schema: english.localized(s.name), problems: bound.problems.map((p) => english.text(p)), found: bound.found };
    }
    if (schemas.length) this.mismatch = { schema: closest?.found ? closest.schema : undefined, problems: closest?.problems ?? [] };
  }

  /** Opens a workbook from its bytes. Throws for files that aren't workbooks or exceed the limits. */
  static open(bytes: Uint8Array | ArrayBuffer, options: OpenWorkbookOptions = {}): HeadlessWorkbook {
    return new HeadlessWorkbook(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes), options);
  }

  private apply(updates: CellUpdate[]): void {
    for (const u of updates) this.views[u.sheet]?.apply(u);
  }

  private sheetIndex(name: string): number {
    const i = this.views.findIndex((v) => v.info.name.toLowerCase() === name.toLowerCase());
    if (i < 0) throw new Error(`There is no sheet called "${name}".`);
    return i;
  }

  get sheetNames(): string[] {
    return this.views.map((v) => v.info.name);
  }

  /** What the file holds that the editor keeps but doesn't show, and how its formulas were calculated (in English). */
  get findings(): Finding[] {
    return [...this.session.loadResult().findings, ...this.session.ensureEngine().findings].map((f) => ({
      ...f,
      title: english.text(f.title),
      detail: english.text(f.detail),
    }));
  }

  /** Every schema problem, once all async checks have finished. Empty without a schema. */
  async issues(): Promise<WorkbookIssue[]> {
    const ctl = this.controller;
    if (!ctl) return [];
    ctl.validate();
    await ctl.settled();
    return ctl.issues().map((i) => ({
      severity: i.severity,
      message: english.text(i.message),
      sheet: i.sheetName,
      ...(i.row !== undefined && i.col !== undefined ? { ref: toA1(i.row, i.col) } : {}),
    }));
  }

  /** The values in the shape of the matched schema; undefined without one. */
  values(): SchemaValues | undefined {
    return this.controller?.values();
  }

  cell(sheet: string, ref: string) {
    return workbookReaders(this.views, this.styles, this.date1904).cell(sheet, ref);
  }

  range(sheet: string, ref?: string) {
    return workbookReaders(this.views, this.styles, this.date1904).range(sheet, ref);
  }

  /**
   * Types text into a cell, as a user would: parsed by the cell's format, the
   * schema's field type and locale, then the schema's business logic runs.
   * Resolves to what changed; `rejected` says why input wasn't accepted.
   */
  async set(sheet: string, ref: string, text: string): Promise<EditResult> {
    const s = this.sheetIndex(sheet);
    const at = parseA1(ref);
    // The A1 parser takes row 0 and columns past XFD; a column past XFD would land on the next row.
    if (!at || at.row < 0 || at.row >= MAX_ROWS || at.col >= MAX_COLS) throw new Error(`"${ref}" isn't a cell reference.`);
    const { row, col } = at;
    const ctl = this.controller;
    const prepared = ctl?.prepareInput(s, row, col, text) ?? { text };
    if ('reject' in prepared) return inEnglish({ updates: [], undo: [], redo: [], skipped: 0, skippedReasons: [], warnings: [], newStyles: [], editedCount: this.session.editedCount, rejected: prepared.reject });
    const reason = ctl?.readOnlyReason(s, row, col);
    if (reason) return inEnglish({ updates: [], undo: [], redo: [], skipped: 1, skippedReasons: [reason], warnings: [], newStyles: [], editedCount: this.session.editedCount, rejected: reason });
    return this.edit({ type: 'input', sheet: s, row, col, text: prepared.text });
  }

  /** Applies any edit the editor supports (paste, fill, sort, clear…), then the schema's business logic. */
  async edit(op: EditOp): Promise<EditResult> {
    const ctl = this.controller;
    const result = this.session.apply(op);
    // The views still hold the old values here: what business logic sees as "before".
    const before: { sheet: number; row: number; col: number; before: FieldValue }[] = ctl
      ? result.redo.map((c) => ({ sheet: c.sheet, row: c.row, col: c.col, before: ctl.valueAt(c.sheet, c.row, c.col) }))
      : [];
    this.apply(result.updates);
    if (ctl && before.length) await ctl.afterEdit(before);
    return inEnglish(result);
  }

  /** Every change since the workbook was opened. */
  changes(): CellChange[] {
    return this.session.changes().map((c) => (c.note === undefined ? c : { ...c, note: english.text(c.note) }));
  }

  get editedCount(): number {
    return this.session.editedCount;
  }

  /**
   * The schema's last word before saving: its errors (unless the schema saves
   * with confirmation) and its beforeSave hook.
   */
  async checkBeforeSave(): Promise<{ ok: true } | { ok: false; message: string; canOverride: boolean }> {
    const check = await (this.controller?.checkBeforeSave() ?? Promise.resolve({ ok: true as const }));
    return check.ok ? check : { ...check, message: english.text(check.message) };
  }

  /**
   * The workbook as .xlsx bytes, stamped with the matched schema's version
   * (unless `stamp: false`). An unchanged workbook comes back as it was opened.
   */
  save(options: { stamp?: boolean } = {}): Uint8Array {
    return this.session.save(options.stamp === false ? undefined : stampFor(this.schema));
  }
}

/** Checks a workbook against schemas and sums up whether it passes. */
export async function validateWorkbook(
  bytes: Uint8Array | ArrayBuffer,
  options: OpenWorkbookOptions = {},
): Promise<{
  ok: boolean;
  schema?: { id: string; name: string; version?: string };
  mismatch?: { schema?: string; problems: string[] };
  note?: string;
  errors: number;
  warnings: number;
  issues: WorkbookIssue[];
  values?: SchemaValues;
}> {
  const wb = HeadlessWorkbook.open(bytes, options);
  const issues = await wb.issues();
  const errors = issues.filter((i) => i.severity === 'error').length;
  return {
    ok: !wb.mismatch && errors === 0,
    ...(wb.schema ? { schema: { id: wb.schema.id, name: english.localized(wb.schema.name), version: wb.schema.version } } : {}),
    ...(wb.mismatch ? { mismatch: wb.mismatch } : {}),
    ...(wb.note ? { note: wb.note } : {}),
    errors,
    warnings: issues.length - errors,
    issues,
    values: wb.values(),
  };
}
