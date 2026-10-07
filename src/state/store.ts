import { mergeEditResults } from '../core/editResult';
import type { EditorTheme } from '../editor/theme';
import type { IconOverrides } from '../ui/icons';
import type { StructuralOp } from '../core/refTransform';
import { shiftFormula } from '../core/formula';
import { inRect, MAX_COLS, MAX_ROWS, parseRange, toA1 } from '../core/refs';
import { bindSchema, type Binding } from '../schema/binding';
import { SchemaController } from '../schema/controller';
import type { FieldValue, NumberOptions, SchemaValues, WorkbookSchema } from '../schema/types';
import {
  CellFlag,
  VT,
  type CalcReport,
  type DisplayStyle,
  type CellChange,
  type EditOp,
  type EditResult,
  type Finding,
  type ListValidation,
  type Rect,
  type RuleValidation,
  type ValidationAlert,
} from '../core/types';
import type { OpenOptions } from '../core/session';
import type { WorkerRequest, WorkerResponse } from '../worker/protocol';
import { deleteDraft, fileKey, loadDraft, saveDraft, type DraftOp } from './drafts';
import { pasteRisks } from './pasteRisks';
import { parseClipboardText } from './clipboard';
import { displayText, editText, exactText } from './format';
import { workbookReaders, type WorkbookReaders } from './readers';
import { stampedFirst, stampFor, versionNote } from '../schema/stamp';
import { DEFAULT_LIST_ERROR, listAllows, listOptions, listValidationAt } from './validation';
import { SheetView } from './sheetView';

export interface Toast {
  id: number;
  kind: 'info' | 'warn' | 'error';
  text: string;
}

export type Answer = 'confirm' | 'alternate' | 'cancel';

export interface ConfirmRequest {
  title: string;
  body: string;
  confirmLabel: string;
  /** A second action besides confirming, e.g. "Keep as text". */
  alternateLabel?: string;
  cancelLabel?: string;
  /** A table shown under the message. */
  details?: { header: string[]; rows: string[][]; more?: number };
  /** Style the confirm button as destructive. Default true. */
  danger?: boolean;
  resolve: (answer: Answer) => void;
}

/** One action in the change history. */
export interface HistoryEntry {
  at: number;
  user?: string;
  action: string;
  cells: number;
  /** Up to five of the cells it changed, e.g. "Prices!B4". */
  refs: string[];
}

/**
 * Throw from onSave when the stored file changed since it was opened. The
 * editor keeps the user's changes and offers to download them as a copy.
 */
export class SaveConflictError extends Error {
  constructor(message = 'Someone else saved a newer version of this workbook.') {
    super(message);
    this.name = 'SaveConflictError';
  }
}

/**
 * The workbook as saved, handed to the host's onSave callback (and the
 * element's `save` event): the file itself, what changed, and its values, so
 * the host can store the file, the values, or both.
 */
export interface SavedWorkbook extends WorkbookReaders {
  bytes: Uint8Array;
  fileName: string;
  /** Changes since the workbook was opened: cells, plus any rows or columns inserted or deleted, or cells moved. */
  editedCount: number;
  /** Each change with its value when opened and now. */
  changes: CellChange[];
  /** What was done, by whom and when, oldest first. */
  history: HistoryEntry[];
  /** The schema the workbook was checked against, and what it found, when one applied. */
  schema?: { id: string; name: string; version?: string; errors: number; warnings: number };
  /** The values in the shape of that schema: each table's records by field, each single cell. */
  values?: SchemaValues;
}

export interface EditorFeatures {
  /** Open button, file picker and drag-and-drop. Default true. */
  open: boolean;
  /** Save button and Cmd/Ctrl+S. Default true. */
  save: boolean;
  /** With a host handling saves, a separate button to download a copy (an export). Default true. */
  download: boolean;
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
  /** The Save button's label when the host handles saving, e.g. "Save to ITAS". Default "Save". */
  saveLabel?: string;
  /** Overrides for the theme tokens (colours, fonts, sizes, corners, shadows); see EditorTheme. */
  theme?: EditorTheme;
  /** Replacement icons by name: SVG path data on a 24×24 grid, or (for React hosts) an element. */
  icons?: IconOverrides;
  /** Ask before leaving the page with unsaved changes. Default true. */
  warnOnUnload?: boolean;
  /** Parts of the UI to show; hosts that open and save files themselves can hide them. */
  features?: Partial<EditorFeatures>;
  /** Locale for parsing and showing dates and numbers; defaults to the browser's. */
  locale?: string;
  /** Name recorded against each action in the change history. */
  user?: string;
  /** Keep unsaved edits in the browser and offer them back when the same file is reopened. Default true. */
  drafts?: boolean;
  /** Size and complexity limits for opened files; see OpenOptions. */
  limits?: OpenOptions['limits'];
  /** VBA macros: keep them (default), remove them when saving, or refuse such files. */
  macros?: OpenOptions['macros'];
  /**
   * Live recalculation: 'auto' (default) turns it off for workbooks too large
   * to recalculate quickly, where the user can turn it on; 'on' and 'off' force it.
   */
  recalc?: OpenOptions['recalc'];
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

export type Panel = 'issues' | 'changes' | 'report';

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
  /** Cut rather than copied: pasting moves the cells, once. */
  cut?: boolean;
}

/** An undoable step: cell edits, undone by restoring cells, or a structural change, undone in the worker. */
export type UndoEntry = EditResult | { structure: string };
const isStructure = (e: UndoEntry): e is { structure: string } => 'structure' in e;

type Pending = { resolve: (r: WorkerResponse) => void; reject: (e: Error) => void };

const rectCells = (r: Rect) => (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);

/**
 * Application state. The workbook itself lives in a Web Worker; this keeps
 * a view of each sheet for rendering, the undo history, and the UI state
 * around it (dialogs, toasts, panels).
 */
interface FormulaList {
  version: number;
  options: string[] | undefined;
  ready: Promise<string[] | undefined>;
}

export interface FindQuery {
  text: string;
  matchCase: boolean;
  /** The whole cell must match, not just part of it. */
  wholeCell: boolean;
  /** Search what cells show, or what they contain (formulas as typed). */
  lookIn: 'values' | 'formulas';
  scope: 'sheet' | 'workbook';
}

/** A structural change in words, for the history and undo. */
function describeStructure(op: StructuralOp, sheets: SheetView[]): string {
  const n = (count: number, one: string) => `${count} ${one}${count === 1 ? '' : 's'}`;
  switch (op.type) {
    case 'insertRows':
      return `Inserted ${n(op.count, 'row')}`;
    case 'deleteRows':
      return `Deleted ${n(op.count, 'row')}`;
    case 'insertCols':
      return `Inserted ${n(op.count, 'column')}`;
    case 'deleteCols':
      return `Deleted ${n(op.count, 'column')}`;
    case 'move':
      return op.toSheet === op.sheet ? `Moved ${rangeLabel(op.src)}` : `Moved ${rangeLabel(op.src)} to ${sheets[op.toSheet]?.info.name ?? 'another sheet'}`;
  }
}

const rangeLabel = (r: Rect) => (r.r1 === r.r2 && r.c1 === r.c2 ? toA1(r.r1, r.c1) : `${toA1(r.r1, r.c1)}:${toA1(r.r2, r.c2)}`);

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
  undoStack: UndoEntry[] = [];
  redoStack: UndoEntry[] = [];
  toasts: Toast[] = [];
  confirmRequest?: ConfirmRequest;
  panel: Panel | null = null;
  /** The grid's right-click menu, while open: where, and whether it was opened on a column or row heading. */
  menu?: { x: number; y: number; kind: 'cell' | 'column' | 'row' };
  /** The value picker for filtering a column, while open. */
  filterPicker?: { col: number; x: number; y: number };
  findOpen = false;
  /** The list of keyboard shortcuts. */
  shortcutsOpen = false;
  /** Asking for a column's width, from the keyboard. */
  columnWidthFor?: number;
  designerOpen = false;
  bannerDismissed = false;
  highlightEdits = true;
  /** The user's override for exact numbers; undefined follows the schema or defaults. */
  exactNumbers?: boolean;
  schemas: WorkbookSchema[];
  schema?: SchemaController;
  schemaMismatch?: SchemaMismatch;
  /** Set when the file was last saved with another version of the schema now applied to it. */
  schemaNote?: string;
  /** The schema being edited in the designer, kept while the designer is closed or picking cells. */
  schemaDraft?: WorkbookSchema;
  /** The designer is waiting for cells to be selected on the sheet, for a new table or a single cell. */
  designerPick?: 'table' | 'cell';
  /** The workbook's custom document properties, as opened. */
  properties: Record<string, string> = {};
  /** Changes whenever schema issues change, so the grid redraws its markers. */
  schemaVersion = 0;
  history: HistoryEntry[] = [];
  /** The latest list of changed cells, refreshed while the Changes panel is open. */
  changeList: CellChange[] = [];
  private draftKey?: string;
  private draftOps: DraftOp[] = [];
  private draftTimer?: ReturnType<typeof setTimeout>;
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
  private readonly formulaLists = new Map<string, FormulaList>();
  /** Bumped whenever cell values change, so formula-based lists are re-evaluated. */
  private dataVersion = 0;

  readonly features: EditorFeatures;

  constructor(options: StoreOptions = {}) {
    this.options = options;
    this.schemas = [...(options.schemas ?? [])];
    this.features = { open: true, save: true, download: true, schemas: true, ...options.features };
  }

  /** The theme tokens set through options or setTheme. */
  get theme(): EditorTheme | undefined {
    return this.options.theme;
  }

  /** Bumped when the theme changes, so the canvas grid reads its colours again. */
  themeVersion = 0;

  /** Changes the theme while the editor is open. */
  setTheme(theme: EditorTheme | undefined): void {
    this.options.theme = theme;
    this.themeVersion++;
    this.notify();
  }

  get icons(): IconOverrides | undefined {
    return this.options.icons;
  }

  /** The Save button's label. */
  get saveLabel(): string {
    return this.hasSaveHandler ? (this.options.saveLabel ?? 'Save') : 'Save as Excel';
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

  /** Number handling in effect: defaults, then the schema's, then the user's toggle. */
  get numbers(): Required<NumberOptions> {
    const base: Required<NumberOptions> = {
      display: 'formatted',
      copy: 'formatted',
      paste: 'lenient',
      reviewBeforeSave: false,
      ...this.options.numbers,
      ...this.schema?.schema.numbers,
    };
    if (this.exactNumbers === undefined) return base;
    return this.exactNumbers ? { ...base, display: 'full', copy: 'raw' } : { ...base, display: 'formatted', copy: 'formatted' };
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
    this.refreshConditionalFormats(i);
    this.notify();
  }

  private readonly cfRuns = new Map<number, { version: number; running: boolean; warned?: boolean }>();

  /**
   * Re-evaluates a sheet's conditional formatting in the worker if cell values
   * have changed since it last was. Runs one at a time per sheet; a change
   * made while one runs triggers another when it finishes.
   */
  refreshConditionalFormats(sheet: number): void {
    const view = this.sheets[sheet];
    if (!view?.info.conditionalFormats.length) return;
    let run = this.cfRuns.get(sheet);
    if (!run) this.cfRuns.set(sheet, (run = { version: -1, running: false }));
    if (run.running || run.version === this.dataVersion) return;
    const state = run;
    const version = this.dataVersion;
    state.running = true;
    void this.request({ type: 'conditionalFormats', sheet } as Omit<WorkerRequest, 'id'>)
      .then((res) => {
        if (res.type !== 'conditionalFormats' || this.sheets[sheet] !== view) return;
        view.setConditionalFormats(res.result);
        if (res.result.skipped && !state.warned) {
          state.warned = true;
          this.toast('info', `${res.result.skipped.toLocaleString()} conditional format rule(s) on ${view.info.name} cover too many cells to show here.`);
        }
        this.notify();
      })
      .catch(() => undefined)
      .finally(() => {
        // Recorded even on failure, so a rule that can't be evaluated isn't retried in a loop.
        state.version = version;
        state.running = false;
        if (this.sheets[sheet] === view && sheet === this.active) this.refreshConditionalFormats(sheet);
      });
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
      this.applyCalc(msg.report);
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
    if (this.draftKey && this.editedCount > 0) void deleteDraft(this.draftKey);
    // Hash before the bytes move to the worker; the hash identifies this file's drafts.
    const key = this.options.drafts === false ? undefined : await fileKey(bytes).catch(() => undefined);
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
      history: [],
      changeList: [],
      draftKey: key,
      draftOps: [],
    });
    this.selections.clear();
    this.notify();
    try {
      const res = await this.request(
        { type: 'open', bytes, fileName, locale: this.locale || 'en-US', options: { limits: this.options.limits, macros: this.options.macros, recalc: this.options.recalc } } as Omit<WorkerRequest, 'id'>,
        [bytes],
      );
      if (res.type !== 'opened') throw new Error('Unexpected response');
      const r = res.result;
      this.sheets = r.sheets.map((info, i) => new SheetView(info, r.grids[i]));
      this.styles = r.styles;
      this.findings = r.findings;
      this.date1904 = r.date1904;
      this.properties = r.properties ?? {};
      this.timings = r.timings;
      this.active = Math.max(0, r.sheets.findIndex((s) => s.state === 'visible'));
      this.status = 'ready';
      this.stage = '';
      this.cfRuns.clear();
      this.bindSchemas();
      this.refreshConditionalFormats(this.active);
    } catch (err) {
      if ((err as Error).message === 'Superseded') return;
      this.status = 'error';
      this.error = (err as Error).message;
    }
    this.notify();
    if (this.status === 'ready' && key) await this.offerDraft(key);
  }

  /** Offers back edits left unsaved the last time this exact file was open. */
  private async offerDraft(key: string): Promise<void> {
    const draft = await loadDraft(key);
    if (!draft?.ops.length || this.draftKey !== key) return;
    const when = new Date(draft.updatedAt).toLocaleString();
    const answer = await this.ask({
      title: 'Restore unsaved changes?',
      body: `This file was edited on ${when} and not saved (${draft.editedCount.toLocaleString()} changed cell${draft.editedCount === 1 ? '' : 's'}). Restore those changes?`,
      confirmLabel: 'Restore changes',
      cancelLabel: 'Discard them',
      danger: false,
    });
    if (answer !== 'confirm') {
      void deleteDraft(key);
      return;
    }
    // Cell edits replay as batches between structural changes, which replay on their own.
    let batch: EditOp[] = [];
    const flush = async () => {
      if (batch.length) await this.edit({ type: 'batch', ops: batch }, 'undo', { label: 'Restored unsaved changes', replay: true });
      batch = [];
    };
    for (const op of draft.ops) {
      if (op.type === 'structure') {
        await flush();
        await this.restructure(op.op, { replay: true });
      } else {
        batch.push(op);
      }
    }
    await flush();
  }

  private scheduleDraft(): void {
    if (!this.draftKey) return;
    clearTimeout(this.draftTimer);
    this.draftTimer = setTimeout(() => {
      const key = this.draftKey;
      if (!key) return;
      if (this.editedCount === 0) void deleteDraft(key);
      else void saveDraft({ key, fileName: this.fileName, updatedAt: Date.now(), editedCount: this.editedCount, ops: this.draftOps });
    }, 400);
  }

  /** Re-reads the list of changed cells from the worker. */
  async refreshChanges(): Promise<void> {
    try {
      const res = await this.request({ type: 'changes' } as Omit<WorkerRequest, 'id'>);
      if (res.type === 'changes') {
        this.changeList = res.changes;
        this.notify();
      }
    } catch {
      // The workbook was closed meanwhile.
    }
  }

  /** The edited workbook, without running save checks or the host's onSave. */
  async save(): Promise<SavedWorkbook | undefined> {
    await this.queue;
    this.busy = true;
    this.notify();
    try {
      const res = await this.request({ type: 'save', properties: this.schemaStamp() } as Omit<WorkerRequest, 'id'>);
      if (res.type !== 'saved') return undefined;
      const ctl = this.schema;
      const summary = ctl?.summary();
      return {
        bytes: res.bytes,
        fileName: res.fileName,
        changes: res.changes,
        history: [...this.history],
        editedCount: this.editedCount,
        ...(ctl && summary
          ? {
              schema: { id: ctl.schema.id, name: ctl.schema.name, version: ctl.schema.version, errors: summary.errors, warnings: summary.warnings },
              values: ctl.values(),
            }
          : {}),
        ...this.readers(),
      };
    } catch (err) {
      this.toast('error', `Could not save: ${(err as Error).message}`);
      return undefined;
    } finally {
      this.busy = false;
      this.notify();
    }
  }

  /**
   * The Save button: schema checks, an optional review of every change,
   * then the host's onSave or a download.
   */
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
    if (this.numbers.reviewBeforeSave && saved.changes.length) {
      const LIMIT = 300;
      const answer = await this.ask({
        title: `Save ${saved.changes.length.toLocaleString()} changed cell${saved.changes.length === 1 ? '' : 's'}?`,
        body: 'Check each value before it goes into the file.',
        confirmLabel: 'Save',
        danger: false,
        details: {
          header: ['Cell', 'Was', 'Now'],
          rows: saved.changes.slice(0, LIMIT).map((c) => [`${c.sheet}!${c.ref}`, c.before, c.formatOnly ? `${c.after} (format changed)` : c.after]),
          more: Math.max(0, saved.changes.length - LIMIT),
        },
      });
      if (answer !== 'confirm') return;
    }
    if (this.options.onSave) {
      this.busy = true;
      this.notify();
      try {
        await this.options.onSave(saved);
        this.afterSave();
        this.toast('info', `Saved ${saved.fileName}.`);
      } catch (err) {
        if (err instanceof SaveConflictError || (err as Error)?.name === 'SaveConflictError') {
          const answer = await this.ask({
            title: 'This workbook was changed elsewhere',
            body: `${(err as Error).message} Your changes are still here. Download them as a copy so nothing is lost?`,
            confirmLabel: 'Download my copy',
            cancelLabel: 'Keep editing',
            danger: false,
          });
          if (answer === 'confirm') this.downloadFile({ ...saved, fileName: saved.fileName.replace(/(\.xls[xm])$/i, ' (my changes)$1') });
        } else {
          this.toast('error', `Not saved: ${(err as Error).message}`);
        }
      } finally {
        this.busy = false;
        this.notify();
      }
      return;
    }
    this.downloadFile(saved);
    this.afterSave();
    this.toast('info', saved.editedCount ? `Saved ${saved.fileName} with ${saved.editedCount.toLocaleString()} changed cells.` : `Saved ${saved.fileName} (unchanged).`);
  }

  private downloadFile(saved: SavedWorkbook): void {
    const type = /\.xlsm$/i.test(saved.fileName)
      ? 'application/vnd.ms-excel.sheet.macroEnabled.12'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const url = URL.createObjectURL(new Blob([saved.bytes as BlobPart], { type }));
    const a = document.createElement('a');
    a.href = url;
    a.download = saved.fileName;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);
  }

  /** Once saved, the draft for this file is no longer needed. */
  private afterSave(): void {
    clearTimeout(this.draftTimer);
    if (this.draftKey) void deleteDraft(this.draftKey);
    this.savedSeq = this.changeSeq;
  }

  /** Bumped by every change; compared with savedSeq to know whether there is anything unsaved. */
  private changeSeq = 0;
  private savedSeq = 0;

  /** There are changes the host (or a download) hasn't received yet. */
  get unsaved(): boolean {
    return this.editedCount > 0 && this.changeSeq !== this.savedSeq;
  }

  /** Readers over the workbook as it is now, for the host to pull values from on save. */
  private readers(): WorkbookReaders {
    return workbookReaders(this.sheets, this.styles, this.date1904);
  }

  /**
   * Downloads a copy of the workbook as it is now, without schema checks.
   * With a host handling saves this is an export: the host still has unsaved changes.
   */
  async exportCopy(): Promise<void> {
    const saved = await this.save();
    if (!saved) return;
    this.downloadFile(saved);
    if (!this.hasSaveHandler) this.afterSave();
    this.toast('info', `Downloaded ${saved.fileName}.`);
  }

  // ---------------------------------------------------------------------------
  // Editing

  /**
   * Runs an edit after any in-flight one, applies its results and records
   * undo. "merge" folds the result into the previous undo step, for edits
   * made by schema logic in response to the user's edit.
   */
  private edit(
    op: EditOp,
    record: 'undo' | 'none' | 'merge' = 'undo',
    opts: { label?: string; replay?: boolean } = {},
  ): Promise<EditResult | undefined> {
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
          if (last && isStructure(last)) this.undoStack.push(last, r);
          else this.undoStack.push(last ? mergeEditResults([last, r]) : r);
        }
        if (r.undo.length) {
          this.recordHistory(op, record, r, opts.label);
          this.draftOps.push(op);
          this.scheduleDraft();
          if (this.panel === 'changes') void this.refreshChanges();
        }
        if (ctl) {
          if (record === 'undo' && before.length && !opts.replay) void ctl.afterEdit(before);
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

  private recordHistory(op: EditOp, record: 'undo' | 'none' | 'merge', r: EditResult, label?: string): void {
    const action =
      label ??
      (record === 'merge'
        ? 'Filled in by business logic'
        : ({ input: 'Edit', pasteText: 'Paste', copy: 'Paste', fill: 'Fill', clear: 'Clear', restore: 'Undo or redo', inputs: 'Edit', sort: 'Sort', batch: 'Edit' } as const)[op.type]);
    const refs = r.redo.slice(0, 5).map((c) => `${this.sheets[c.sheet]?.info.name ?? '?'}!${toA1(c.row, c.col)}`);
    this.history.push({ at: Date.now(), user: this.options.user, action, cells: r.redo.length, refs });
  }

  private applyResult(r: EditResult): void {
    this.dataVersion++;
    if (r.undo.length) this.changeSeq++;
    for (const s of r.newStyles) this.styles[s.index] = s.style;
    for (const u of r.updates) this.sheets[u.sheet]?.apply(u);
    this.editedCount = r.editedCount;
    if (r.rejected) this.toast('error', r.rejected);
    if (r.skipped) {
      const reason = r.skippedReasons[0] ?? '';
      this.toast('warn', r.skipped === 1 ? reason : `${r.skipped.toLocaleString()} cells were left unchanged. ${reason}`);
    }
    for (const w of r.warnings) this.toast('warn', w);
    this.refreshConditionalFormats(this.active);
  }

  async undo(): Promise<void> {
    await this.queue;
    const entry = this.undoStack.pop();
    if (!entry) return;
    if (isStructure(entry)) {
      if (await this.swapStructure('undoStructure', `Undo: ${entry.structure}`)) this.redoStack.push(entry);
      else this.undoStack = [];
      this.notify();
      return;
    }
    this.showSheet(entry.undo[0]?.sheet);
    await this.edit({ type: 'restore', cells: entry.undo }, 'none', { label: 'Undo' });
    this.redoStack.push(entry);
    this.notify();
  }

  async redo(): Promise<void> {
    await this.queue;
    const entry = this.redoStack.pop();
    if (!entry) return;
    if (isStructure(entry)) {
      if (await this.swapStructure('redoStructure', `Redo: ${entry.structure}`)) this.undoStack.push(entry);
      this.notify();
      return;
    }
    this.showSheet(entry.redo[0]?.sheet);
    await this.edit({ type: 'restore', cells: entry.redo }, 'none', { label: 'Redo' });
    this.undoStack.push(entry);
    this.notify();
  }

  // ---------------------------------------------------------------------------
  // Structural changes: inserting and deleting rows and columns, moving cells

  /**
   * Inserts or deletes rows or columns, or moves cells, with every reference
   * to the cells that move following them. The worker rebuilds the workbook,
   * so all the sheets are reloaded here. Resolves false if it was refused.
   */
  restructure(op: StructuralOp, opts: { replay?: boolean } = {}): Promise<boolean> {
    const run = async () => {
      if (this.status !== 'ready') return false;
      this.busy = true;
      this.notify();
      try {
        const res = await this.request({ type: 'structure', op } as Omit<WorkerRequest, 'id'>);
        if (res.type !== 'restructured') return false;
        this.adoptWorkbook(res);
        const label = describeStructure(op, this.sheets);
        this.undoStack.push({ structure: label });
        this.redoStack = [];
        this.history.push({ at: Date.now(), user: this.options.user, action: label, cells: 0, refs: [] });
        this.draftOps.push({ type: 'structure', op });
        this.scheduleDraft();
        if (!opts.replay) this.selectAfter(op);
        for (const w of res.warnings) this.toast('warn', w);
        if (this.panel === 'changes') void this.refreshChanges();
        return true;
      } catch (err) {
        if ((err as Error).message !== 'Superseded') this.toast('error', (err as Error).message);
        return false;
      } finally {
        this.busy = false;
        this.notify();
      }
    };
    const p = this.queue.then(run, run);
    this.queue = p;
    return p;
  }

  private async swapStructure(type: 'undoStructure' | 'redoStructure', label: string): Promise<boolean> {
    try {
      this.busy = true;
      this.notify();
      const res = await this.request({ type } as Omit<WorkerRequest, 'id'>);
      if (res.type !== 'restructured') return false;
      this.adoptWorkbook(res);
      this.history.push({ at: Date.now(), user: this.options.user, action: label, cells: 0, refs: [] });
      if (this.panel === 'changes') void this.refreshChanges();
      return true;
    } catch (err) {
      this.toast('error', (err as Error).message);
      return false;
    } finally {
      this.busy = false;
    }
  }

  /** Replaces every sheet's view with the worker's rebuilt workbook, keeping the active sheet and selections. */
  private adoptWorkbook(res: Extract<WorkerResponse, { type: 'restructured' }>): void {
    const r = res.result;
    this.sheets = r.sheets.map((info, i) => new SheetView(info, r.grids[i]));
    this.styles = r.styles;
    this.date1904 = r.date1904;
    this.findings = res.report ? [...r.findings.filter((f) => f.kind !== 'formulaEngine'), ...res.report.findings] : r.findings;
    this.calc = res.report;
    for (const u of res.updates) this.sheets[u.sheet]?.apply(u);
    this.editedCount = res.editedCount;
    this.changeSeq++;
    this.active = Math.min(this.active, this.sheets.length - 1);
    this.dataVersion++;
    this.formulaLists.clear();
    this.cfRuns.clear();
    this.clip = undefined;
    this.bindSchemas();
    this.refreshConditionalFormats(this.active);
  }

  /** Turns live recalculation on for a workbook opened with it off. Building the engine can take a while. */
  async enableRecalc(): Promise<void> {
    await this.queue;
    this.recalcStarting = true;
    this.busy = true;
    this.notify();
    try {
      const res = await this.request({ type: 'enableRecalc' } as Omit<WorkerRequest, 'id'>);
      if (res.type === 'recalcOn') this.applyCalc(res.report);
    } catch (err) {
      this.toast('error', (err as Error).message);
    } finally {
      this.recalcStarting = false;
      this.busy = false;
      this.notify();
    }
  }

  /** Set while the engine is being built after the user turned recalculation on. */
  recalcStarting = false;

  private applyCalc(report: CalcReport): void {
    this.calc = report;
    this.findings = [...this.findings.filter((f) => f.kind !== 'formulaEngine'), ...report.findings];
    for (const u of report.updates) this.sheets[u.sheet]?.apply(u);
    this.dataVersion++;
    this.schema?.validate();
    this.refreshConditionalFormats(this.active);
  }

  /** Selects what a structural change made or left, as Excel does. */
  private selectAfter(op: StructuralOp): void {
    const view = this.sheets[op.type === 'move' ? op.toSheet : op.sheet];
    if (!view) return;
    const lastCol = Math.max(view.info.maxCol, 0);
    const lastRow = Math.max(view.info.maxRow, 0);
    const sel = this.selection;
    let rect: Rect;
    if (op.type === 'move') {
      rect = { r1: op.row, c1: op.col, r2: op.row + op.src.r2 - op.src.r1, c2: op.col + op.src.c2 - op.src.c1 };
      this.active = op.toSheet;
    } else if (op.type === 'insertRows') {
      rect = { r1: op.at, r2: op.at + op.count - 1, c1: 0, c2: lastCol };
    } else if (op.type === 'insertCols') {
      rect = { r1: 0, r2: lastRow, c1: op.at, c2: op.at + op.count - 1 };
    } else if (op.type === 'deleteRows') {
      rect = { r1: op.at, r2: op.at, c1: sel.col, c2: sel.col };
    } else {
      rect = { r1: sel.row, r2: sel.row, c1: op.at, c2: op.at };
    }
    this.setSelection({ row: rect.r1, col: op.type === 'insertRows' ? sel.col : rect.c1, rects: [rect] });
  }

  /** Inserts as many rows as the selection spans, above or below it. */
  insertRows(where: 'above' | 'below'): Promise<boolean> {
    const r = this.selection.rects[0] ?? { r1: this.selection.row, r2: this.selection.row, c1: 0, c2: 0 };
    return this.restructure({ type: 'insertRows', sheet: this.active, at: where === 'above' ? r.r1 : r.r2 + 1, count: r.r2 - r.r1 + 1 });
  }

  deleteRows(): Promise<boolean> {
    const r = this.selection.rects[0] ?? { r1: this.selection.row, r2: this.selection.row, c1: 0, c2: 0 };
    return this.restructure({ type: 'deleteRows', sheet: this.active, at: r.r1, count: r.r2 - r.r1 + 1 });
  }

  /** Inserts as many columns as the selection spans, left or right of it. */
  insertCols(where: 'left' | 'right'): Promise<boolean> {
    const r = this.selection.rects[0] ?? { r1: 0, r2: 0, c1: this.selection.col, c2: this.selection.col };
    return this.restructure({ type: 'insertCols', sheet: this.active, at: where === 'left' ? r.c1 : r.c2 + 1, count: r.c2 - r.c1 + 1 });
  }

  deleteCols(): Promise<boolean> {
    const r = this.selection.rects[0] ?? { r1: 0, r2: 0, c1: this.selection.col, c2: this.selection.col };
    return this.restructure({ type: 'deleteCols', sheet: this.active, at: r.c1, count: r.c2 - r.c1 + 1 });
  }

  // ---------------------------------------------------------------------------
  // Sorting and filtering

  /**
   * The block a sort or filter applies to: the table or filter range around
   * the active cell, else the contiguous block of non-empty cells around it,
   * Excel's "current region". Its first row is taken as headings.
   */
  dataRegion(sheet = this.active, row = this.selection.row, col = this.selection.col): Rect | undefined {
    const view = this.sheets[sheet];
    if (!view) return undefined;
    const sel = this.selection.rects[0];
    if (sel && (sel.r2 > sel.r1 || sel.c2 > sel.c1) && sel.r2 < MAX_ROWS - 1 && sel.c2 < MAX_COLS - 1) return sel;
    const table = view.tableAt(row, col);
    if (table) return { ...table.rect, r2: table.rect.r2 - table.totalsRows };
    const af = view.filter?.rect ?? view.info.autoFilter;
    if (af && inRect(af, row, col)) return af;
    const filled = (r: number, c: number) => r >= 0 && c >= 0 && !!view.get(r, c) && view.get(r, c)!.t !== VT.None;
    if (!filled(row, col)) return undefined;
    const rect = { r1: row, r2: row, c1: col, c2: col };
    // Grow while any cell just outside an edge (diagonals included) is filled.
    for (let grew = true; grew; ) {
      grew = false;
      const rowHas = (r: number) => {
        for (let c = rect.c1 - 1; c <= rect.c2 + 1; c++) if (filled(r, c)) return true;
        return false;
      };
      const colHas = (c: number) => {
        for (let r = rect.r1 - 1; r <= rect.r2 + 1; r++) if (filled(r, c)) return true;
        return false;
      };
      if (rowHas(rect.r1 - 1)) (rect.r1--, (grew = true));
      if (rowHas(rect.r2 + 1)) (rect.r2++, (grew = true));
      if (colHas(rect.c1 - 1)) (rect.c1--, (grew = true));
      if (colHas(rect.c2 + 1)) (rect.c2++, (grew = true));
    }
    return rect;
  }

  /** Sorts the data region by the active cell's column. Its first row stays put as headings. */
  sort(descending: boolean, col = this.selection.col): Promise<EditResult | undefined> {
    const rect = this.dataRegion();
    if (!rect || rect.r2 <= rect.r1) {
      this.toast('info', 'Select a cell in a block of data to sort it.');
      return Promise.resolve(undefined);
    }
    return this.edit({ type: 'sort', sheet: this.active, rect, keys: [{ col, descending }], header: true }, 'undo', { label: descending ? 'Sort Z to A' : 'Sort A to Z' });
  }

  /** The text a cell shows, as filters and find compare it. */
  shownText(view: SheetView, row: number, col: number): string {
    const rec = view.get(row, col);
    if (!rec) return '';
    return displayText(rec, this.styles[rec.style] ?? this.styles[0], this.date1904).text;
  }

  /** The values in a column of the filter range, with how many rows hold each, for picking what to show. */
  filterChoices(col: number): { rect: Rect; values: { text: string; count: number }[]; allowed?: Set<string> } | undefined {
    const view = this.view;
    if (!view) return undefined;
    const rect = view.filter?.rect ?? this.dataRegion();
    if (!rect || col < rect.c1 || col > rect.c2) return undefined;
    const counts = new Map<string, number>();
    for (let row = rect.r1 + 1; row <= rect.r2; row++) {
      const text = this.shownText(view, row, col);
      counts.set(text, (counts.get(text) ?? 0) + 1);
    }
    const values = [...counts].map(([text, count]) => ({ text, count }));
    values.sort((a, b) => (a.text === '' ? 1 : b.text === '' ? -1 : a.text.localeCompare(b.text, undefined, { numeric: true })));
    return { rect, values, allowed: view.filter?.columns.get(col) };
  }

  /** Shows only rows whose value in `col` is one of `allowed`; undefined shows them all again. */
  setFilter(rect: Rect, col: number, allowed: Set<string> | undefined): void {
    const view = this.view;
    if (!view) return;
    const columns = new Map(view.filter?.columns ?? []);
    if (allowed) columns.set(col, allowed);
    else columns.delete(col);
    view.setFilter(view.filter?.rect ?? rect, columns, (row, c) => this.shownText(view, row, c));
    this.notify();
  }

  clearFilter(): void {
    const view = this.view;
    if (!view?.filter) return;
    view.setFilter(view.filter.rect, new Map(), () => '');
    this.notify();
  }

  // ---------------------------------------------------------------------------
  // Find and replace

  /** Cells matching a search, in reading order: by sheet, then row, then column. */
  find(q: FindQuery): { sheet: number; row: number; col: number }[] {
    if (!q.text) return [];
    const out: { sheet: number; row: number; col: number }[] = [];
    const needle = q.matchCase ? q.text : q.text.toLowerCase();
    const sheets = q.scope === 'workbook' ? this.sheets.map((_, i) => i) : [this.active];
    for (const s of sheets) {
      const view = this.sheets[s];
      if (view.info.kind !== 'worksheet') continue;
      const hits: { row: number; col: number }[] = [];
      view.forEach({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }, (row, col, rec) => {
        if (rec.t === VT.None && rec.f === undefined) return;
        let text = q.lookIn === 'formulas' ? editText(rec, this.styles[rec.style] ?? this.styles[0], this.date1904) : this.shownText(view, row, col);
        if (!q.matchCase) text = text.toLowerCase();
        if (q.wholeCell ? text === needle : text.includes(needle)) hits.push({ row, col });
      });
      hits.sort((a, b) => a.row - b.row || a.col - b.col);
      out.push(...hits.map((h) => ({ sheet: s, ...h })));
    }
    return out;
  }

  /** Replaces the search text in matching cells' contents (formulas included). Resolves to how many cells changed. */
  async replace(q: FindQuery, replacement: string, only?: { sheet: number; row: number; col: number }): Promise<number> {
    const hits = only ? [only] : this.find({ ...q, lookIn: 'formulas' });
    const flags = q.matchCase ? 'g' : 'gi';
    const pattern = new RegExp(q.text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags);
    const bySheet = new Map<number, { row: number; col: number; text: string }[]>();
    let skipped = 0;
    for (const h of hits) {
      const view = this.sheets[h.sheet];
      const rec = view.get(h.row, h.col);
      if (!rec || view.readOnlyReason(h.row, h.col)) {
        skipped++;
        continue;
      }
      const before = editText(rec, this.styles[rec.style] ?? this.styles[0], this.date1904);
      const after = q.wholeCell ? replacement : before.replace(pattern, () => replacement);
      if (after === before) continue;
      const list = bySheet.get(h.sheet) ?? [];
      list.push({ row: h.row, col: h.col, text: after });
      bySheet.set(h.sheet, list);
    }
    const ops: EditOp[] = [...bySheet].map(([sheet, cells]) => ({ type: 'inputs', sheet, cells }));
    const n = ops.reduce((k, o) => k + (o.type === 'inputs' ? o.cells.length : 0), 0);
    if (n) await this.edit(ops.length === 1 ? ops[0] : { type: 'batch', ops }, 'undo', { label: 'Replace' });
    if (skipped) this.toast('info', `${skipped.toLocaleString()} matching cell${skipped === 1 ? " is" : 's are'} read-only here and ${skipped === 1 ? 'was' : 'were'} left as ${skipped === 1 ? 'it was' : 'they were'}.`);
    return n;
  }

  /** Bumped to ask the grid to scroll the active cell into view. */
  revealSeq = 0;

  /** Set with a reveal when the grid should also take the keyboard focus. */
  revealFocus = false;

  /** Moves to a cell, switching sheet if needed, and scrolls it into view. */
  reveal(sheet: number, row: number, col: number): void {
    if (sheet !== this.active) this.setActive(sheet);
    this.revealSeq++;
    this.setSelection({ row, col, rects: [{ r1: row, c1: col, r2: row, c2: col }] });
  }

  /** Moves the keyboard focus to the grid, at the active cell. */
  focusGrid(): void {
    this.revealFocus = true;
    this.revealSeq++;
    this.notify();
  }

  /** A column's width as shown here, in Excel's units (characters of the default font). */
  columnWidth(col: number): number {
    return Math.round(((this.view?.colWidth(col) ?? 64) / 7) * 100) / 100;
  }

  /** Sets how wide a column is shown. The file keeps its own widths, as with resizing by dragging. */
  setColumnWidth(col: number, chars: number): void {
    const view = this.view;
    if (!view || !(chars >= 0)) return;
    view.colWidths.set(col, Math.min(2000, Math.round(chars * 7)));
    this.colVersion++;
    this.notify();
  }

  /** Bumped when column widths change outside the grid. */
  colVersion = 0;

  /**
   * Goes to a reference typed in the name box, as Excel does: a cell or range,
   * on another sheet with "Prices!B4". The grid takes the focus, so typing
   * goes into the cell. False when the text isn't a reference to a sheet here.
   */
  goTo(text: string): boolean {
    const m = /^(?:(?:'((?:[^']|'')+)'|([^!']+))!)?([^!]+)$/.exec(text.trim());
    if (!m) return false;
    const name = m[1]?.replace(/''/g, "'") ?? m[2];
    const sheet = name === undefined ? this.active : this.sheets.findIndex((s) => s.info.name.toLowerCase() === name.toLowerCase());
    const rect = parseRange(m[3].trim());
    if (sheet < 0 || !rect || rect.r1 < 0 || rect.r2 >= MAX_ROWS || rect.c2 >= MAX_COLS) return false;
    if (sheet !== this.active) this.setActive(sheet);
    this.revealFocus = true;
    this.revealSeq++;
    this.setSelection({ row: rect.r1, col: rect.c1, rects: [rect] });
    return true;
  }

  /** The selected cells as tab-separated text, as they are copied to the clipboard. */
  selectionText(rect: Rect): string {
    const view = this.view;
    if (!view) return '';
    const lines: string[] = [];
    for (let row = rect.r1; row <= Math.min(rect.r2, Math.max(view.info.maxRow, rect.r1)); row++) {
      if (view.gridRow(row) < 0) continue;
      const cells: string[] = [];
      for (let col = rect.c1; col <= Math.min(rect.c2, Math.max(view.info.maxCol, rect.c1)); col++) {
        const rec = view.get(row, col);
        const style = rec ? (this.styles[rec.style] ?? this.styles[0]) : undefined;
        const raw = rec && style && this.numbers.copy === 'raw' ? exactText(rec, style, false) : undefined;
        const text = rec && style ? (raw ?? displayText(rec, style, this.date1904).text) : '';
        cells.push(/[\t\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text);
      }
      lines.push(cells.join('\t'));
    }
    return lines.join('\n');
  }

  /** Copies or cuts the selection from the menu, where the keyboard's clipboard events aren't available. */
  async copySelection(cut: boolean): Promise<void> {
    const rect = this.selection.rects[0];
    if (!rect) return;
    this.noteCopy(rect, cut);
    try {
      await navigator.clipboard.writeText(this.selectionText(rect));
    } catch {
      // Pasting within this workbook still works; other apps won't see it.
      this.toast('info', `${cut ? 'Cut' : 'Copied'} for pasting here. The browser didn't allow putting it on the clipboard for other apps; use ${cut ? '⌘X or Ctrl+X' : '⌘C or Ctrl+C'} for that.`);
    }
  }

  /** Pastes from the menu. Browsers ask permission to read the clipboard this way; the keyboard doesn't need it. */
  async pasteFromClipboard(): Promise<void> {
    const sel = this.selection;
    const rect = sel.rects[0] ?? { r1: sel.row, c1: sel.col, r2: sel.row, c2: sel.col };
    let text: string | undefined;
    try {
      text = await navigator.clipboard.readText();
    } catch {
      text = undefined;
    }
    if (text === undefined) {
      // Without the clipboard, paste what was cut or copied here, if anything.
      if (this.clip) return this.paste({ row: rect.r1, col: rect.c1 }, [], rect, true);
      this.toast('warn', "The browser didn't allow reading the clipboard. Use ⌘V or Ctrl+V instead.");
      return;
    }
    await this.paste({ row: rect.r1, col: rect.c1 }, parseClipboardText(text), rect);
  }

  /** The range marked by the last cut, shown with a dashed outline until it is pasted. */
  get cutRange(): { sheet: number; rect: Rect } | undefined {
    return this.clip?.cut ? { sheet: this.clip.sheet, rect: this.clip.rect } : undefined;
  }

  cancelCut(): void {
    if (!this.clip?.cut) return;
    this.clip = undefined;
    this.notify();
  }

  private showSheet(sheet: number | undefined): void {
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
    if (!(await this.passesValidation(row, col, text))) return;
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

  /** Records what was copied (or cut) so a paste back into the workbook can carry formulas, or move the cells. */
  noteCopy(rect: Rect, cut = false): void {
    const view = this.view;
    if (!view) return;
    const rec = view.get(rect.r1, rect.c1);
    const style = rec ? (this.styles[rec.style] ?? this.styles[0]) : undefined;
    const raw = rec && style && this.numbers.copy === 'raw' ? exactText(rec, style, false) : undefined;
    const first = rec && style ? (raw ?? displayText(rec, style, this.date1904).text) : '';
    this.clip = { sheet: this.active, rect, first, cut };
    this.notify();
  }

  /**
   * Pastes clipboard text, or what was cut or copied in this workbook when
   * the text is the same (so formulas come along, and a cut moves cells).
   * With `ownClipboard`, pastes what was cut or copied here without the text,
   * for when the browser won't let the page read the clipboard.
   */
  async paste(target: { row: number; col: number }, values: readonly (readonly string[])[], selection: Rect, ownClipboard = false): Promise<void> {
    const view = this.view;
    const clip = this.clip;
    if (!view || (!values.length && !(ownClipboard && clip))) return;
    const h = ownClipboard && clip ? clip.rect.r2 - clip.rect.r1 + 1 : values.length;
    const w = ownClipboard && clip ? clip.rect.c2 - clip.rect.c1 + 1 : Math.max(...values.map((v) => v.length));
    const selH = selection.r2 - selection.r1 + 1;
    const selW = selection.c2 - selection.c1 + 1;
    // Like Excel, a selection that is a whole multiple of the copied block is filled with it.
    const tile = selH >= h && selW >= w && selH % h === 0 && selW % w === 0 && (selH > h || selW > w);
    const dst: Rect = tile ? selection : { r1: target.row, c1: target.col, r2: target.row + h - 1, c2: target.col + w - 1 };

    const internal =
      clip &&
      (ownClipboard || (clip.rect.r2 - clip.rect.r1 + 1 === h && clip.rect.c2 - clip.rect.c1 + 1 === w && (values[0][0] ?? '') === clip.first));
    if (internal && clip.cut && !tile) {
      // Pasting a cut moves the cells, and references to them follow.
      await this.restructure({ type: 'move', sheet: clip.sheet, src: clip.rect, toSheet: this.active, row: dst.r1, col: dst.c1 });
      return;
    }
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
    if (this.numbers.paste === 'strict') {
      // Ask before storing values that wouldn't read back the way they were pasted.
      const risks = pasteRisks(values, this.locale);
      if (risks.length) {
        const answer = await this.ask({
          title: risks.length === 1 ? 'Check a pasted value' : `Check ${risks.length} pasted values`,
          body: 'These would be stored differently from how they read. Paste them as shown, or keep them exactly as text?',
          confirmLabel: 'Paste as shown',
          alternateLabel: 'Keep as text',
          danger: false,
          details: {
            header: ['Pasted', 'Would become', 'Because'],
            rows: risks.slice(0, 20).map((r) => [r.text, r.becomes, r.reason]),
            more: Math.max(0, risks.length - 20),
          },
        });
        if (answer === 'cancel') return;
        if (answer === 'alternate') {
          const risky = new Set(risks.map((r) => r.text));
          grid = grid.map((line) => line.map((v) => (risky.has(v.trim()) ? "'" + v : v)));
        }
      }
    }
    if (tile) {
      const block = grid;
      grid = [];
      for (let r = 0; r < selH; r++) grid.push(Array.from({ length: selW }, (_, c) => block[r % h]?.[c % w] ?? ''));
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

  /**
   * Excel's Ctrl+D and Ctrl+R: copies the selection's top row down through it
   * (or its left column across), or with one row (column) selected, the cells
   * above (to the left). Formulas adjust as if copied; series don't extend.
   */
  async fillDown(across = false): Promise<void> {
    const view = this.view;
    const rect = this.selection.rects[0];
    if (!view || !rect) return;
    const single = across ? rect.c1 === rect.c2 : rect.r1 === rect.r2;
    if (single && (across ? rect.c1 : rect.r1) === 0) return;
    const src: Rect = across
      ? { ...rect, c1: single ? rect.c1 - 1 : rect.c1, c2: single ? rect.c1 - 1 : rect.c1 }
      : { ...rect, r1: single ? rect.r1 - 1 : rect.r1, r2: single ? rect.r1 - 1 : rect.r1 };
    const dst: Rect = single ? rect : across ? { ...rect, c1: rect.c1 + 1 } : { ...rect, r1: rect.r1 + 1 };
    const keeps = (row: number, col: number) => !!((view.get(across ? row : src.r1, across ? src.c1 : col)?.flags ?? 0) & CellFlag.Formula);
    if (!(await this.okToReplace(this.formulasLost(view, [dst], keeps)))) return;
    await this.edit({ type: 'copy', srcSheet: this.active, src, sheet: this.active, dst }, 'undo', { label: across ? 'Fill right' : 'Fill down' });
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

  /**
   * A list's allowed values at a cell. Range lists are re-read when their
   * sheet changes; formula lists are evaluated by the worker, so the first
   * call returns undefined and the editor updates when the values arrive.
   */
  listOptions(dv: ListValidation, row = dv.rects[0].r1, col = dv.rects[0].c1): string[] | undefined {
    if (dv.source.kind !== 'formula') {
      const version = dv.source.kind === 'range' ? (this.sheets[dv.source.sheet]?.version ?? 0) : -1;
      const hit = this.listCache.get(dv);
      if (hit && hit.version === version) return hit.options;
      const options = listOptions(dv, this.sheets, this.styles, this.date1904);
      this.listCache.set(dv, { version, options });
      return options;
    }
    return this.formulaList(dv, row, col).options;
  }

  /**
   * A formula list at a cell: the values last evaluated (shown while newer
   * ones load) and a promise of its values for the current data.
   */
  private formulaList(dv: ListValidation, row: number, col: number): FormulaList {
    const source = dv.source.kind === 'formula' ? dv.source.formula : '';
    // The formula is written for the rule's first cell, e.g. INDIRECT(C2) for a whole column.
    const formula = shiftFormula(source, dv.rects[0].r1, dv.rects[0].c1, row, col) ?? source;
    const sheet = this.sheets.findIndex((v) => v.info.listValidations.includes(dv));
    const key = `${sheet}|${formula}`;
    const hit = this.formulaLists.get(key);
    if (hit && hit.version === this.dataVersion) return hit;
    const entry: FormulaList = {
      version: this.dataVersion,
      options: hit?.options,
      ready: this.request({ type: 'listSource', sheet, formula } as Omit<WorkerRequest, 'id'>)
        .then((res) => {
          const src = res.type === 'listSource' ? res.source : null;
          entry.options = src ? listOptions({ ...dv, source: src }, this.sheets, this.styles, this.date1904) : undefined;
          this.notify();
          return entry.options;
        })
        .catch(() => undefined),
    };
    this.formulaLists.set(key, entry);
    return entry;
  }

  /** The non-list validation rule at a cell, with its index on the sheet. */
  ruleAt(sheet: number, row: number, col: number): { rule: RuleValidation; index: number } | undefined {
    const rules = this.sheets[sheet]?.info.ruleValidations ?? [];
    const index = rules.findIndex((r) => r.rects.some((rect) => inRect(rect, row, col)));
    return index < 0 ? undefined : { rule: rules[index], index };
  }

  /** Applies a rule's alert as Excel does; resolves false when the value is rejected. */
  private async alert(dv: ValidationAlert, fallbackTitle: string): Promise<boolean> {
    const message = dv.error ?? DEFAULT_LIST_ERROR;
    if (dv.errorStyle === 'stop') {
      this.toast('error', dv.errorTitle ? `${dv.errorTitle}: ${message}` : message);
      return false;
    }
    if (dv.errorStyle === 'warning') {
      return this.confirm({ title: dv.errorTitle ?? fallbackTitle, body: `${message} Keep this value anyway?`, confirmLabel: 'Keep value' });
    }
    this.toast('info', message);
    return true;
  }

  /** Checks typed text against the cell's Excel validation rules. Resolves false to reject it. */
  private async passesValidation(row: number, col: number, text: string): Promise<boolean> {
    if (text === '' || text.startsWith('=')) return true;
    const dv = this.listAt(this.active, row, col);
    if (dv?.showError) {
      const options = dv.source.kind === 'formula' ? await this.formulaList(dv, row, col).ready : this.listOptions(dv, row, col);
      if (options && !listAllows(options, text.replace(/^'/, ''))) return this.alert(dv, 'Not in the list');
    }
    const rule = this.ruleAt(this.active, row, col);
    if (rule?.rule.showError) {
      const res = await this.request({ type: 'validate', sheet: this.active, row, col, text, rule: rule.index } as Omit<WorkerRequest, 'id'>).catch(() => undefined);
      if (res?.type === 'validated' && !res.ok) return this.alert(rule.rule, 'Value not allowed');
    }
    return true;
  }

  // ---------------------------------------------------------------------------
  // Schemas

  /** Tries each registered schema against the open workbook; the first that fits applies. */
  bindSchemas(): void {
    for (const v of this.sheets) v.constraint = undefined;
    this.schema = undefined;
    this.schemaMismatch = undefined;
    this.schemaNote = undefined;
    if (this.status !== 'ready' || !this.schemas.length) return this.notify();
    let closest: { name: string; problems: string[]; found: number } | undefined;
    // A file stamped with a schema when it was saved tries that schema first.
    for (const s of stampedFirst(this.schemas, this.properties)) {
      const r = bindSchema(s, this.sheets);
      if (r.binding) {
        this.schemaNote = versionNote(s, this.properties);
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

  /** The stamp saving writes into the file: the applied schema's id and version, if it has a version. */
  private schemaStamp(): Record<string, string> | undefined {
    return stampFor(this.schema?.schema);
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

  /** Shows a dialog and resolves with the button chosen. */
  ask(req: Omit<ConfirmRequest, 'resolve'>): Promise<Answer> {
    return new Promise((resolve) => {
      this.confirmRequest?.resolve('cancel');
      this.confirmRequest = {
        ...req,
        resolve: (answer) => {
          this.confirmRequest = undefined;
          this.notify();
          resolve(answer);
        },
      };
      this.notify();
    });
  }

  async confirm(req: Omit<ConfirmRequest, 'resolve'>): Promise<boolean> {
    return (await this.ask(req)) === 'confirm';
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

  set<
    K extends
      | 'panel'
      | 'bannerDismissed'
      | 'highlightEdits'
      | 'exactNumbers'
      | 'designerOpen'
      | 'menu'
      | 'filterPicker'
      | 'findOpen'
      | 'shortcutsOpen'
      | 'columnWidthFor'
      | 'schemaDraft'
      | 'designerPick',
  >(
    key: K,
    value: WorkbookStore[K],
  ): void {
    (this as unknown as Record<string, unknown>)[key] = value;
    this.notify();
  }
}
