import { mergeEditResults } from '../core/editResult';
import { inRect } from '../core/refs';
import { bindSchema, type Binding } from '../schema/binding';
import { SchemaController } from '../schema/controller';
import type { FieldValue, NumberOptions, WorkbookSchema } from '../schema/types';
import {
  CellFlag,
  type CalcReport,
  type DisplayStyle,
  type EditOp,
  type EditResult,
  type Finding,
  type ListValidation,
  type Rect,
} from '../core/types';
import type { WorkerRequest, WorkerResponse } from '../worker/protocol';
import { displayText, exactText } from './format';
import { DEFAULT_LIST_ERROR, listAllows, listOptions, listValidationAt } from './validation';
import { SheetView } from './sheetView';

export interface Toast {
  id: number;
  kind: 'info' | 'warn' | 'error';
  text: string;
}

export interface ConfirmRequest {
  title: string;
  body: string;
  confirmLabel: string;
  resolve: (ok: boolean) => void;
}

/** The workbook as saved, handed to the host's onSave callback. */
export interface SavedWorkbook {
  bytes: Uint8Array;
  fileName: string;
  /** Cells changed since the workbook was opened. */
  editedCount: number;
}

export interface EditorFeatures {
  /** Open button, file picker and drag-and-drop. Default true. */
  open: boolean;
  /** Save button and Cmd/Ctrl+S. Default true. */
  save: boolean;
  /** Schema designer button. Default true. */
  schemas: boolean;
}

export interface StoreOptions {
  /** Schemas to try, in order, against each workbook that is opened. */
  schemas?: WorkbookSchema[];
  /** Default number display; a matching schema can override it. */
  numbers?: NumberOptions;
  /** Creates the worker that holds the workbook; defaults to the bundled worker file. */
  createWorker?: () => Worker;
  /**
   * Receives the workbook when the user saves, e.g. to upload it. Throwing
   * shows the error and keeps the changes. Without it, saving downloads the file.
   */
  onSave?: (file: SavedWorkbook) => void | Promise<void>;
  /** Ask before leaving the page with unsaved changes. Default true. */
  warnOnUnload?: boolean;
  /** Parts of the UI to show; hosts that open and save files themselves can hide them. */
  features?: Partial<EditorFeatures>;
  /** Locale for parsing and showing dates and numbers; defaults to the browser's. */
  locale?: string;
}

function defaultWorker(): Worker {
  return new Worker(new URL('../worker/workbook.worker.ts', import.meta.url), { type: 'module' });
}

export interface SchemaMismatch {
  /** The schema that came closest, if any of its sheets were found. */
  name?: string;
  problems: string[];
  tried: number;
}

export type Panel = 'issues' | 'report';

export interface Selection {
  /** Active cell, in sheet coordinates. */
  row: number;
  col: number;
  /** Selected ranges, in sheet coordinates. */
  rects: Rect[];
}

interface Clip {
  sheet: number;
  rect: Rect;
  first: string;
}

type Pending = { resolve: (r: WorkerResponse) => void; reject: (e: Error) => void };

const rectCells = (r: Rect) => (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);

/**
 * Application state. The workbook itself lives in a Web Worker; this keeps
 * a view of each sheet for rendering, the undo history, and the UI state
 * around it (dialogs, toasts, panels).
 */
export class WorkbookStore {
  status: 'empty' | 'loading' | 'ready' | 'error' = 'empty';
  stage = '';
  error?: string;
  fileName = '';
  sheets: SheetView[] = [];
  styles: DisplayStyle[] = [];
  findings: Finding[] = [];
  calc?: CalcReport;
  date1904 = false;
  timings: Record<string, number> = {};
  active = 0;
  selections = new Map<number, Selection>();
  editedCount = 0;
  busy = false;
  undoStack: EditResult[] = [];
  redoStack: EditResult[] = [];
  toasts: Toast[] = [];
  confirmRequest?: ConfirmRequest;
  panel: Panel | null = null;
  designerOpen = false;
  bannerDismissed = false;
  highlightEdits = true;
  /** The user's override for exact numbers; undefined follows the schema or defaults. */
  exactNumbers?: boolean;
  schemas: WorkbookSchema[];
  schema?: SchemaController;
  schemaMismatch?: SchemaMismatch;
  /** Changes whenever schema issues change, so the grid redraws its markers. */
  schemaVersion = 0;
  version = 0;
  private readonly options: StoreOptions;

  private worker?: Worker;
  private seq = 0;
  private pending = new Map<number, Pending>();
  private listeners = new Set<() => void>();
  private clip?: Clip;
  private queue: Promise<unknown> = Promise.resolve();
  private toastSeq = 0;
  private readonly listCache = new WeakMap<ListValidation, { version: number; options: string[] | undefined }>();

  readonly features: EditorFeatures;

  constructor(options: StoreOptions = {}) {
    this.options = options;
    this.schemas = [...(options.schemas ?? [])];
    this.features = { open: true, save: true, schemas: true, ...options.features };
  }

  /** Whether the host handles saving instead of the browser downloading the file. */
  get hasSaveHandler(): boolean {
    return !!this.options.onSave;
  }

  get warnOnUnload(): boolean {
    return this.options.warnOnUnload !== false;
  }

  get locale(): string {
    return this.options.locale ?? (typeof navigator !== 'undefined' ? navigator.language : '') ?? 'en-US';
  }

  /** Stops the worker and releases the workbook. */
  dispose(): void {
    this.worker?.terminate();
    this.worker = undefined;
    for (const p of this.pending.values()) p.reject(new Error('Superseded'));
    this.pending.clear();
    this.listeners.clear();
  }

  get reportOpen(): boolean {
    return this.panel === 'report';
  }

  /** Number display in effect: defaults, then the schema's, then the user's toggle. */
  get numbers(): Required<NumberOptions> {
    const base: Required<NumberOptions> = { display: 'formatted', copy: 'formatted', ...this.options.numbers, ...this.schema?.schema.numbers };
    if (this.exactNumbers === undefined) return base;
    return this.exactNumbers ? { display: 'full', copy: 'raw' } : { display: 'formatted', copy: 'formatted' };
  }

  subscribe = (fn: () => void) => {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  };

  getVersion = () => this.version;

  notify(): void {
    this.version++;
    for (const fn of this.listeners) fn();
  }

  get view(): SheetView | undefined {
    return this.sheets[this.active];
  }

  get selection(): Selection {
    return this.selections.get(this.active) ?? { row: 0, col: 0, rects: [{ r1: 0, c1: 0, r2: 0, c2: 0 }] };
  }

  setSelection(sel: Selection): void {
    this.selections.set(this.active, sel);
    this.notify();
  }

  setActive(i: number): void {
    this.active = i;
    this.notify();
  }

  // ---------------------------------------------------------------------------
  // Worker plumbing

  private request(msg: Omit<WorkerRequest, 'id'> & { type: WorkerRequest['type'] }, transfer: Transferable[] = []): Promise<WorkerResponse> {
    const worker = this.worker;
    if (!worker) return Promise.reject(new Error('No workbook is open.'));
    const id = ++this.seq;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      worker.postMessage({ ...msg, id } as WorkerRequest, transfer);
    });
  }

  private onMessage = (e: MessageEvent<WorkerResponse>) => {
    const msg = e.data;
    if (msg.type === 'calc') {
      this.calc = msg.report;
      this.findings = [...this.findings.filter((f) => f.kind !== 'formulaEngine'), ...msg.report.findings];
      for (const u of msg.report.updates) this.sheets[u.sheet]?.apply(u);
      this.schema?.validate();
      this.notify();
      return;
    }
    const p = this.pending.get(msg.id);
    if (!p) return;
    this.pending.delete(msg.id);
    if (msg.type === 'error') p.reject(new Error(msg.message));
    else p.resolve(msg);
  };

  // ---------------------------------------------------------------------------
  // Opening and saving

  async open(fileName: string, bytes: ArrayBuffer): Promise<void> {
    if (this.editedCount > 0 && !(await this.confirm({
      title: 'Discard your changes?',
      body: `You have unsaved changes in ${this.fileName}. Opening another file discards them.`,
      confirmLabel: 'Discard and open',
    }))) {
      return;
    }
    if (!/\.xls[xm]$/i.test(fileName)) {
      this.toast('error', 'Only Excel workbooks (.xlsx or .xlsm) can be opened.');
      return;
    }
    this.worker?.terminate();
    for (const p of this.pending.values()) p.reject(new Error('Superseded'));
    this.pending.clear();
    // A fresh worker per file releases the previous workbook's memory.
    this.worker = (this.options.createWorker ?? defaultWorker)();
    this.worker.onmessage = this.onMessage;
    Object.assign(this, {
      status: 'loading',
      stage: `Reading ${fileName}…`,
      error: undefined,
      fileName,
      calc: undefined,
      undoStack: [],
      redoStack: [],
      editedCount: 0,
      bannerDismissed: false,
      panel: null,
      clip: undefined,
      schema: undefined,
      schemaMismatch: undefined,
    });
    this.selections.clear();
    this.notify();
    try {
      const res = await this.request(
        { type: 'open', bytes, fileName, locale: this.locale || 'en-US' } as Omit<WorkerRequest, 'id'>,
        [bytes],
      );
      if (res.type !== 'opened') throw new Error('Unexpected response');
      const r = res.result;
      this.sheets = r.sheets.map((info, i) => new SheetView(info, r.grids[i]));
      this.styles = r.styles;
      this.findings = r.findings;
      this.date1904 = r.date1904;
      this.timings = r.timings;
      this.active = Math.max(0, r.sheets.findIndex((s) => s.state === 'visible'));
      this.status = 'ready';
      this.stage = '';
      this.bindSchemas();
    } catch (err) {
      if ((err as Error).message === 'Superseded') return;
      this.status = 'error';
      this.error = (err as Error).message;
    }
    this.notify();
  }

  async save(): Promise<{ bytes: Uint8Array; fileName: string } | undefined> {
    await this.queue;
    this.busy = true;
    this.notify();
    try {
      const res = await this.request({ type: 'save' } as Omit<WorkerRequest, 'id'>);
      if (res.type !== 'saved') return undefined;
      return { bytes: res.bytes, fileName: res.fileName };
    } catch (err) {
      this.toast('error', `Could not save: ${(err as Error).message}`);
      return undefined;
    } finally {
      this.busy = false;
      this.notify();
    }
  }

  async download(): Promise<void> {
    if (this.schema) {
      this.busy = true;
      this.notify();
      const check = await this.schema.checkBeforeSave();
      this.busy = false;
      this.notify();
      if (!check.ok) {
        if (!check.canOverride) {
          this.set('panel', 'issues');
          this.toast('error', `Not saved. ${check.message}`);
          return;
        }
        const go = await this.confirm({ title: 'Save with errors?', body: `${check.message} Save the file anyway?`, confirmLabel: 'Save anyway' });
        if (!go) return;
      }
    }
    const saved = await this.save();
    if (!saved) return;
    if (this.options.onSave) {
      this.busy = true;
      this.notify();
      try {
        await this.options.onSave({ ...saved, editedCount: this.editedCount });
        this.toast('info', `Saved ${saved.fileName}.`);
      } catch (err) {
        this.toast('error', `Not saved: ${(err as Error).message}`);
      } finally {
        this.busy = false;
        this.notify();
      }
      return;
    }
    const type = /\.xlsm$/i.test(saved.fileName)
      ? 'application/vnd.ms-excel.sheet.macroEnabled.12'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const url = URL.createObjectURL(new Blob([saved.bytes as BlobPart], { type }));
    const a = document.createElement('a');
    a.href = url;
    a.download = saved.fileName;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);
    this.toast('info', this.editedCount ? `Saved ${saved.fileName} with ${this.editedCount.toLocaleString()} changed cells.` : `Saved ${saved.fileName} (unchanged).`);
  }

  // ---------------------------------------------------------------------------
  // Editing

  /**
   * Runs an edit after any in-flight one, applies its results and records
   * undo. "merge" folds the result into the previous undo step, for edits
   * made by schema logic in response to the user's edit.
   */
  private edit(op: EditOp, record: 'undo' | 'none' | 'merge' = 'undo'): Promise<EditResult | undefined> {
    const run = async () => {
      this.busy = true;
      this.notify();
      try {
        const res = await this.request({ type: 'edit', op } as Omit<WorkerRequest, 'id'>);
        if (res.type !== 'edited') return undefined;
        const r = res.result;
        const ctl = this.schema;
        // Schema logic sees each edited cell's value from before the edit.
        const before = record === 'undo' && ctl ? r.redo.map((c) => ({ sheet: c.sheet, row: c.row, col: c.col, before: ctl.valueAt(c.sheet, c.row, c.col) as FieldValue })) : [];
        this.applyResult(r);
        if (record === 'undo' && r.undo.length) {
          this.undoStack.push(r);
          this.redoStack = [];
        } else if (record === 'merge' && r.undo.length) {
          const last = this.undoStack.pop();
          this.undoStack.push(last ? mergeEditResults([last, r]) : r);
        }
        if (ctl) {
          if (record === 'undo' && before.length) void ctl.afterEdit(before);
          else ctl.validate();
        }
        return r;
      } catch (err) {
        if ((err as Error).message !== 'Superseded') this.toast('error', (err as Error).message);
        return undefined;
      } finally {
        this.busy = false;
        this.notify();
      }
    };
    const p = this.queue.then(run, run);
    this.queue = p;
    return p;
  }

  private applyResult(r: EditResult): void {
    for (const s of r.newStyles) this.styles[s.index] = s.style;
    for (const u of r.updates) this.sheets[u.sheet]?.apply(u);
    this.editedCount = r.editedCount;
    if (r.rejected) this.toast('error', r.rejected);
    if (r.skipped) {
      const reason = r.skippedReasons[0] ?? '';
      this.toast('warn', r.skipped === 1 ? reason : `${r.skipped.toLocaleString()} cells were left unchanged. ${reason}`);
    }
    for (const w of r.warnings) this.toast('warn', w);
  }

  async undo(): Promise<void> {
    await this.queue;
    const entry = this.undoStack.pop();
    if (!entry) return;
    this.goTo(entry.undo[0]?.sheet);
    await this.edit({ type: 'restore', cells: entry.undo }, 'none');
    this.redoStack.push(entry);
    this.notify();
  }

  async redo(): Promise<void> {
    await this.queue;
    const entry = this.redoStack.pop();
    if (!entry) return;
    this.goTo(entry.redo[0]?.sheet);
    await this.edit({ type: 'restore', cells: entry.redo }, 'none');
    this.undoStack.push(entry);
    this.notify();
  }

  private goTo(sheet: number | undefined): void {
    if (sheet !== undefined && sheet !== this.active) this.setActive(sheet);
  }

  /** How many formulas in `rects` would be replaced by something that is not a formula. */
  private formulasLost(view: SheetView, rects: Rect[], keeps?: (row: number, col: number) => boolean): number {
    let n = 0;
    for (const rect of rects) {
      view.forEach(rect, (row, col, rec) => {
        if (!(rec.flags & CellFlag.Formula)) return;
        if (view.readOnlyReason(row, col)) return;
        if (keeps?.(row, col)) return;
        n++;
      });
    }
    return n;
  }

  private async okToReplace(n: number): Promise<boolean> {
    if (n === 0) return true;
    return this.confirm({
      title: n === 1 ? 'Replace a formula?' : `Replace ${n.toLocaleString()} formulas?`,
      body:
        n === 1
          ? 'This cell holds a formula. It will be replaced with a plain value, and the formula will not be in the saved file. You can undo this.'
          : 'Some of these cells hold formulas. They will be replaced with plain values, and those formulas will not be in the saved file. You can undo this.',
      confirmLabel: n === 1 ? 'Replace formula' : 'Replace formulas',
    });
  }

  async input(row: number, col: number, text: string): Promise<void> {
    const view = this.view;
    if (!view) return;
    const reason = view.readOnlyReason(row, col);
    if (reason) return this.toast('warn', reason);
    const prepared = this.schema?.prepareInput(this.active, row, col, text) ?? { text };
    if ('reject' in prepared) return this.toast('warn', prepared.reject);
    if (!(await this.passesList(row, col, text))) return;
    const rec = view.get(row, col);
    if (rec?.flags && rec.flags & CellFlag.Formula && !text.startsWith('=') && !(await this.okToReplace(1))) return;
    await this.edit({ type: 'input', sheet: this.active, row, col, text: prepared.text });
  }

  async clear(rects: Rect[]): Promise<void> {
    const view = this.view;
    if (!view) return;
    if (!(await this.okToReplace(this.formulasLost(view, rects)))) return;
    await this.edit({ type: 'clear', sheet: this.active, rects });
  }

  /** Records what was copied so a paste back into the workbook can carry formulas. */
  noteCopy(rect: Rect): void {
    const view = this.view;
    if (!view) return;
    const rec = view.get(rect.r1, rect.c1);
    const style = rec ? (this.styles[rec.style] ?? this.styles[0]) : undefined;
    const raw = rec && style && this.numbers.copy === 'raw' ? exactText(rec, style, false) : undefined;
    const first = rec && style ? (raw ?? displayText(rec, style, this.date1904).text) : '';
    this.clip = { sheet: this.active, rect, first };
  }

  async paste(target: { row: number; col: number }, values: readonly (readonly string[])[], selection: Rect): Promise<void> {
    const view = this.view;
    if (!view || !values.length) return;
    const clip = this.clip;
    const h = values.length;
    const w = Math.max(...values.map((v) => v.length));
    const selH = selection.r2 - selection.r1 + 1;
    const selW = selection.c2 - selection.c1 + 1;
    // Like Excel, a selection that is a whole multiple of the copied block is filled with it.
    const tile = selH >= h && selW >= w && selH % h === 0 && selW % w === 0 && (selH > h || selW > w);
    const dst: Rect = tile ? selection : { r1: target.row, c1: target.col, r2: target.row + h - 1, c2: target.col + w - 1 };

    const internal =
      clip &&
      clip.rect.r2 - clip.rect.r1 + 1 === h &&
      clip.rect.c2 - clip.rect.c1 + 1 === w &&
      (values[0][0] ?? '') === clip.first;
    if (internal) {
      const src = this.sheets[clip.sheet];
      const keeps = (row: number, col: number) => {
        const sr = clip.rect.r1 + ((row - dst.r1) % h);
        const sc = clip.rect.c1 + ((col - dst.c1) % w);
        return !!((src.get(sr, sc)?.flags ?? 0) & CellFlag.Formula);
      };
      if (!(await this.okToReplace(this.formulasLost(view, [dst], keeps)))) return;
      await this.edit({ type: 'copy', srcSheet: clip.sheet, src: clip.rect, sheet: this.active, dst });
      return;
    }

    let grid = values.map((line) => [...line]);
    if (tile) {
      grid = [];
      for (let r = 0; r < selH; r++) grid.push(Array.from({ length: selW }, (_, c) => values[r % h][c % w] ?? ''));
    }
    if (this.schema) {
      for (let i = 0; i < grid.length; i++) {
        for (let j = 0; j < grid[i].length; j++) {
          const prepared = this.schema.prepareInput(this.active, dst.r1 + i, dst.c1 + j, grid[i][j]);
          if ('reject' in prepared) return this.toast('warn', `Nothing was pasted. ${prepared.reject}`);
          grid[i][j] = prepared.text;
        }
      }
    }
    const keeps = (row: number, col: number) => (grid[row - dst.r1]?.[col - dst.c1] ?? '').startsWith('=');
    if (!(await this.okToReplace(this.formulasLost(view, [dst], keeps)))) return;
    if (rectCells(dst) > 200_000) this.toast('info', `Pasting ${rectCells(dst).toLocaleString()} cells…`);
    await this.edit({ type: 'pasteText', sheet: this.active, row: dst.r1, col: dst.c1, values: grid });
  }

  async fill(src: Rect, dst: Rect): Promise<void> {
    const view = this.view;
    if (!view) return;
    const h = src.r2 - src.r1 + 1;
    const w = src.c2 - src.c1 + 1;
    const mod = (a: number, n: number) => ((a % n) + n) % n;
    const keeps = (row: number, col: number) => {
      if (inRect(src, row, col)) return true;
      const sr = src.r1 + mod(row - src.r1, h);
      const sc = src.c1 + mod(col - src.c1, w);
      return !!((view.get(sr, sc)?.flags ?? 0) & CellFlag.Formula);
    };
    if (!(await this.okToReplace(this.formulasLost(view, [dst], keeps)))) return;
    await this.edit({ type: 'fill', sheet: this.active, src, dst });
  }

  // ---------------------------------------------------------------------------
  // Excel list validations

  /** The list validation at a cell, unless a schema field there defines its own values. */
  listAt(sheet: number, row: number, col: number): ListValidation | undefined {
    const view = this.sheets[sheet];
    if (!view?.info.listValidations.length) return undefined;
    if (this.schema?.fieldAt(sheet, row, col)?.field.type === 'enum') return undefined;
    return listValidationAt(view.info, row, col);
  }

  /** A list's allowed values, re-read when its source sheet changes. */
  listOptions(dv: ListValidation): string[] | undefined {
    const version = dv.source.kind === 'range' ? (this.sheets[dv.source.sheet]?.version ?? 0) : -1;
    const hit = this.listCache.get(dv);
    if (hit && hit.version === version) return hit.options;
    const options = listOptions(dv, this.sheets, this.styles, this.date1904);
    this.listCache.set(dv, { version, options });
    return options;
  }

  /** Applies a list validation's alert to typed text, as Excel does. Resolves false to reject. */
  private async passesList(row: number, col: number, text: string): Promise<boolean> {
    const dv = this.listAt(this.active, row, col);
    if (!dv || !dv.showError || text === '' || text.startsWith('=')) return true;
    const options = this.listOptions(dv);
    if (!options || listAllows(options, text.replace(/^'/, ''))) return true;
    const message = dv.error ?? DEFAULT_LIST_ERROR;
    if (dv.errorStyle === 'stop') {
      this.toast('error', dv.errorTitle ? `${dv.errorTitle}: ${message}` : message);
      return false;
    }
    if (dv.errorStyle === 'warning') {
      return this.confirm({ title: dv.errorTitle ?? 'Not in the list', body: `${message} Keep this value anyway?`, confirmLabel: 'Keep value' });
    }
    this.toast('info', message);
    return true;
  }

  // ---------------------------------------------------------------------------
  // Schemas

  /** Tries each registered schema against the open workbook; the first that fits applies. */
  bindSchemas(): void {
    for (const v of this.sheets) v.constraint = undefined;
    this.schema = undefined;
    this.schemaMismatch = undefined;
    if (this.status !== 'ready' || !this.schemas.length) return this.notify();
    let closest: { name: string; problems: string[]; found: number } | undefined;
    for (const s of this.schemas) {
      const r = bindSchema(s, this.sheets);
      if (r.binding) {
        this.attachSchema(r.binding);
        return;
      }
      if (r.found > 0 && (!closest || r.found > closest.found || (r.found === closest.found && r.problems.length < closest.problems.length))) {
        closest = { name: s.name, problems: r.problems, found: r.found };
      }
    }
    this.schemaMismatch = { name: closest?.name, problems: closest?.problems ?? [], tried: this.schemas.length };
    this.notify();
  }

  private attachSchema(binding: Binding): void {
    const ctl = new SchemaController(binding, {
      views: () => this.sheets,
      styles: () => this.styles,
      date1904: () => this.date1904,
      followUp: async (op) => {
        await this.edit(op, 'merge');
      },
      changed: () => {
        this.schemaVersion++;
        this.notify();
      },
    });
    this.schema = ctl;
    this.sheets.forEach((v, i) => {
      v.constraint = (row, col) => ctl.readOnlyReason(i, row, col);
    });
    ctl.validate();
    if (ctl.summary().errors) this.panel = 'issues';
    this.notify();
  }

  /** Adds a schema, replacing any with the same id, and re-checks the open workbook. */
  registerSchema(schema: WorkbookSchema): void {
    // Newest first, so a schema just applied takes priority over older ones that also fit.
    this.schemas = [schema, ...this.schemas.filter((s) => s.id !== schema.id)];
    this.bindSchemas();
  }

  removeSchema(id: string): void {
    this.schemas = this.schemas.filter((s) => s.id !== id);
    this.bindSchemas();
  }

  /** The schema table record under the active cell, when records there can be added or removed. */
  get recordTarget() {
    const sel = this.selection;
    return this.schema?.recordAt(this.active, sel.row, sel.col);
  }

  async insertRecord(): Promise<void> {
    const target = this.recordTarget;
    if (!target || !this.schema) return;
    const op = this.schema.insertRecord(target.table, target.record);
    if (typeof op === 'string') return this.toast('warn', op);
    await this.edit(op);
  }

  async deleteRecords(): Promise<void> {
    const target = this.recordTarget;
    if (!target || !this.schema) return;
    const t = target.table;
    const rect = this.selection.rects[0];
    const span = (pos: { row: number; col: number }) => t.locate(pos.row, pos.col);
    const a = rect ? span({ row: rect.r1, col: rect.c1 }) : undefined;
    const b = rect ? span({ row: rect.r2, col: rect.c2 }) : undefined;
    const first = a?.kind === 'field' ? a.record : target.record;
    const last = b?.kind === 'field' ? b.record : target.record;
    const op = this.schema.deleteRecords(t, Math.min(first, last), Math.max(first, last));
    if (typeof op === 'string') return this.toast('warn', op);
    await this.edit(op);
  }

  // ---------------------------------------------------------------------------
  // Dialogs and messages

  confirm(req: Omit<ConfirmRequest, 'resolve'>): Promise<boolean> {
    return new Promise((resolve) => {
      this.confirmRequest = {
        ...req,
        resolve: (ok) => {
          this.confirmRequest = undefined;
          this.notify();
          resolve(ok);
        },
      };
      this.notify();
    });
  }

  toast(kind: Toast['kind'], text: string): void {
    const id = ++this.toastSeq;
    this.toasts = [...this.toasts.slice(-3), { id, kind, text }];
    this.notify();
    setTimeout(() => this.dismissToast(id), kind === 'error' ? 8000 : 5000);
  }

  dismissToast(id: number): void {
    this.toasts = this.toasts.filter((t) => t.id !== id);
    this.notify();
  }

  set<K extends 'panel' | 'bannerDismissed' | 'highlightEdits' | 'exactNumbers' | 'designerOpen'>(key: K, value: WorkbookStore[K]): void {
    (this as unknown as Record<string, unknown>)[key] = value;
    this.notify();
  }
}
