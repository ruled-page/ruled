import { mergeEditResults } from '../core/editResult';
import type { EditorTheme } from '../editor/theme';
import type { IconOverrides } from '../ui/icons';
import { invertIndexMap, isSheetOp, nearestVisible, nextSheetName, sheetIndexMap, sheetOpProblem, type SheetOp, type WorkbookOp } from '../core/sheetRules';
import { shiftFormula } from '../core/formula';
import { formulaReferences } from '../core/formulaAssist';
import type { NameTarget } from '../core/names';
import { linkDestination, replaceLinks, resolvePlace, type LinkFields } from '../core/links';
import { inRect, MAX_COLS, MAX_ROWS, parseRange, toA1 } from '../core/refs';
import { bindSchema, type Binding, type BindResult } from '../schema/binding';
import { describeMatch, textMatches } from '../schema/json';
import { withHooksFrom } from '../schema/hooks';
import { SchemaController } from '../schema/controller';
import { hasWildcards, matchesWhole, matchSpan, parsePattern, searchIn } from '../core/wildcard';
import { stepDecimals } from './numberFormats';
import type { FieldValue, NumberOptions, SchemaValues, WorkbookSchema } from '../schema/types';
import {
  CellFlag,
  VT,
  type CalcReport,
  type DefinedName,
  type DisplayStyle,
  type CellChange,
  type CommentChange,
  type EditOp,
  type EditResult,
  type BorderEdge,
  type BorderPlacement,
  type Finding,
  type FilterState,
  type FormatChange,
  type PastePart,
  type LayoutOp,
  type ReadOnlyCells,
  type SeriesOptions,
  type ListValidation,
  type NameChange,
  type NameRef,
  type Rect,
  type RuleValidation,
  type ValidationAlert,
} from '../core/types';
import type { OpenOptions } from '../core/session';
import { CSV_FILE, type CsvOptions } from '../core/csv';
import type { WorkerRequest, WorkerResponse } from '../worker/protocol';
import { deleteDraft, fileKey, loadDraft, saveDraft, type DraftOp } from './drafts';
import { namedRange } from './names';
import { pasteRisks } from './pasteRisks';
import { parseClipboardText, transposeGrid } from './clipboard';
import { csvFileName, sheetCsv } from './csv';
import { displayText, editText, exactText, isDateStyle } from './format';
import { builtInLists, cleanLists, loadUserLists, saveUserLists } from './fillLists';
import { seriesLine } from '../core/fill';
import { outOfDateCells, workbookReaders, type WorkbookReaders } from './readers';
import { stampedFirst, stampFor, versionNote } from '../schema/stamp';
import { DEFAULT_LIST_ERROR, listAllows, listOptions, listValidationAt } from './validation';
import { SheetView } from './sheetView';
import { cellLink, type CellLink, type LinkClick } from './links';
import type { i18n as I18n } from 'i18next';
import {
  createI18n,
  english,
  errorText,
  msg,
  TextError,
  translator,
  type Localized,
  type MessageKey,
  type Text,
  type Translations,
  type Translator,
} from '../i18n';
import { isNarrow } from './viewport';

export interface Toast {
  id: number;
  kind: 'info' | 'warn' | 'error';
  text: Text;
}

/** Something said to screen readers and not shown, such as what an action did. */
export interface Spoken {
  id: number;
  text: Text;
}

export type Answer = 'confirm' | 'alternate' | 'cancel';

export interface ConfirmRequest {
  title: Text;
  body: Text;
  confirmLabel: Text;
  /** A second action besides confirming, e.g. "Keep as text". */
  alternateLabel?: Text;
  cancelLabel?: Text;
  /** A table shown under the message; its rows are the workbook's own values. */
  details?: { header: Text[]; rows: string[][]; more?: number };
  /** Style the confirm button as destructive. Default true. */
  danger?: boolean;
  resolve: (answer: Answer) => void;
}

/** One action in the change history, as hosts receive it: the action in the editor's language. */
export interface HistoryEntry {
  at: number;
  user?: string;
  action: string;
  cells: number;
  /** Up to five of the cells it changed, e.g. "Prices!B4". */
  refs: string[];
}

/** One action in the change history as the editor keeps it, translated when shown. */
export interface HistoryItem extends Omit<HistoryEntry, 'action'> {
  action: Text;
}

const SAVE_CONFLICT = 'Someone else saved a newer version of this workbook.';

/**
 * Throw from onSave when the stored file changed since it was opened. The
 * editor keeps the user's changes and offers to download them as a copy.
 * Without a message of the host's own, the editor explains it in its language.
 */
export class SaveConflictError extends Error {
  constructor(message = SAVE_CONFLICT) {
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
  /** What was done, by whom and when, oldest first; actions are in the editor's language. */
  history: HistoryEntry[];
  /** The schema the workbook was checked against, and what it found, when one applied. */
  schema?: { id: string; name: string; version?: string; errors: number; warnings: number };
  /** The values in the shape of that schema: each table's records by field, each single cell. */
  values?: SchemaValues;
  /**
   * Formulas whose saved value may be out of date: this editor can't calculate
   * them, a cell they depend on changed, and they still hold Excel's last
   * result (as does any formula reading one). Excel recalculates them when the
   * file is next opened; until then, don't rely on their values.
   */
  outOfDate: { sheet: string; ref: string }[];
}

/** What opening a workbook came to: open, declined by the user, or refused with the reason. */
export type OpenResult = { opened: true } | { opened: false; cancelled: true } | { opened: false; error: Text };

/**
 * What saving came to: handed to the host's onSave, downloaded (without
 * one), cancelled by the user, blocked by the schema's checks, or failed
 * (onSave threw, or the stored copy had changed).
 */
export type SaveResult = 'saved' | 'downloaded' | 'cancelled' | 'blocked' | 'failed';

export interface EditorFeatures {
  /** Open button, file picker and drag-and-drop. Default true. */
  open: boolean;
  /** Save button and Cmd/Ctrl+S. Default true. */
  save: boolean;
  /** With a host handling saves, a separate button to download a copy (an export). Default true. */
  download: boolean;
  /** Schema designer button. Default true. */
  schemas: boolean;
  /**
   * "What's in this file": what the workbook holds that the editor keeps but
   * doesn't show, and how its formulas are calculated, with the banner that
   * sums it up. Useful to people who know Excel; hosts whose users only fill
   * in a template can turn it off. Default true.
   */
  report: boolean;
  /**
   * Charts drawn from the workbook's current values. The chart library loads
   * only when a workbook has a chart to draw. Off, charts show as placeholders
   * with their title. Default true.
   */
  charts: boolean;
  /**
   * The format bar: fonts, colours, borders, alignment and number formats.
   * Hosts whose users only fill in values can turn it off. Default true.
   */
  format: boolean;
  /**
   * CSV files: the Open button and dropping take .csv, .tsv and .txt files as
   * well as workbooks, and a button downloads the sheet shown as CSV (when
   * downloads are on). Default true.
   */
  csv: boolean;
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
  saveLabel?: Localized;
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
  /**
   * Language of the editor's own text, such as "de" or "fr-CA"; defaults to
   * the locale. English and German are built in; others need `translations`.
   */
  language?: string;
  /**
   * Messages by language (i18next JSON, keys as in the English catalogue):
   * whole catalogues for more languages, or a few messages worded differently.
   */
  translations?: Translations;
  /** Name recorded against each action in the change history. */
  user?: string;
  /**
   * Lists that dragging the fill handle and AutoFill continue, as Excel's
   * custom lists do, after the built-in day and month names: each list is
   * its entries in order, e.g. [['North', 'East', 'South', 'West']]. Users
   * can add their own in the Custom Lists dialog.
   */
  customLists?: string[][];
  /** Keep unsaved edits in the browser and offer them back when the same file is reopened. Default true. */
  drafts?: boolean;
  /** Size and complexity limits for opened files; see OpenOptions. */
  limits?: OpenOptions['limits'];
  /** Macros (VBA or Excel 4.0 macro sheets): keep them (default), remove them when saving, or refuse such files. */
  macros?: OpenOptions['macros'];
  /** ActiveX controls and embedded objects: keep them (default), or refuse such files. */
  activeContent?: OpenOptions['activeContent'];
  /**
   * Live recalculation: 'auto' (default) turns it off for workbooks too large
   * to recalculate quickly, where the user can turn it on; 'on' and 'off' force it.
   */
  recalc?: OpenOptions['recalc'];
  /**
   * Called when the editor stops working: its worker couldn't load, crashed,
   * ran out of memory or stopped answering. The editor shows what happened
   * and offers to reopen the file, which brings back unsaved edits when
   * drafts are on.
   */
  onError?: (error: Error) => void;
  /**
   * How long the worker may spend on one task, in milliseconds, before it is
   * taken to have stopped. Default 120,000; 0 waits for ever.
   */
  workerTimeout?: number;
  /** Hears what the editor does and how long it took, for the host's logging or analytics. */
  onEvent?: (event: EditorEvent) => void;
  /** How the sheet is written when downloaded as CSV: the delimiter, and whether formulas are escaped (they are by default). */
  csvExport?: CsvOptions;
  /**
   * Following hyperlinks. 'follow' (the default): Ctrl+click (⌘-click on a
   * Mac), Ctrl+Enter (⌘Enter) or the cell menu goes to a place in the
   * workbook, or opens a web page or email in a new tab, asking first the
   * first time for each address; links to anything else are refused. 'off':
   * links are shown and can be edited, but never followed.
   */
  links?: 'follow' | 'off';
  /**
   * Called when the user follows a link to a web page or an email, instead
   * of the editor opening it; links to anything else are refused before it is
   * called. Return (or resolve) true to have the editor ask and open it after all.
   */
  onLinkClick?: (link: LinkClick) => boolean | void | Promise<boolean | void>;
}

/**
 * What the editor tells onEvent. Messages are in English, whatever the
 * editor's language, so logs read the same.
 */
export type EditorEvent =
  | { type: 'opened'; fileName: string; ms: number; sheets: number }
  | { type: 'openFailed'; fileName: string; reason: string }
  /** The calculation engine is built (or live recalculation is off): how many formulas it keeps up to date. */
  | { type: 'calculated'; ms: number; formulas: number; live: number; notCalculated: number }
  | { type: 'saved'; result: SaveResult; ms: number; editedCount: number }
  /** An error the editor showed the user. */
  | { type: 'error'; message: string };

/** Default for StoreOptions.workerTimeout. */
const WORKER_TIMEOUT = 120_000;

function defaultWorker(): Worker {
  return new Worker(new URL('../worker/workbook.worker.ts', import.meta.url), { type: 'module' });
}

export interface SchemaMismatch {
  /** The schema that came closest, if any of its sheets were found. */
  name?: Localized;
  problems: Text[];
  tried: number;
}

export type Panel = 'issues' | 'changes' | 'report';

/** The dialog that defines a name, or edits the one `from` identifies, as it opens. */
export interface NameDialog {
  from?: NameRef;
  name: string;
  localSheet?: number;
  /** What it refers to, as typed: "=Data!$A$1:$B$3". */
  refersTo: string;
  comment: string;
}

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

/**
 * A change to the sheets, as undo and redo need it: where each sheet went
 * (by its position before), and the sheet shown before and after.
 */
interface SheetMove {
  map: (number | null)[];
  before: number;
  after: number;
}

/** An undoable step: cell edits, undone by restoring cells, or a structural change, undone in the worker. */
export type UndoEntry = EditResult | { structure: Text; sheets?: SheetMove };
const isStructure = (e: UndoEntry): e is { structure: Text; sheets?: SheetMove } => 'structure' in e;

/** What the schema returns for a record change: the edit to make, or why it can't be made. */
const isEditOp = (x: EditOp | Text): x is EditOp => typeof x === 'object' && typeof (x as { type?: unknown }).type === 'string';

type Pending = { resolve: (r: WorkerResponse) => void; reject: (e: Error) => void; timer?: ReturnType<typeof setTimeout> };

const rectCells = (r: Rect) => (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);

/** Excel's default cell format, until a workbook is open. */
const DEFAULT_STYLE: DisplayStyle = { numFmt: 'General', bold: false, italic: false, underline: false, strike: false, fontSizePx: 14.7, fontFamily: 'Calibri', wrap: false, indent: 0 };

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

/** A structural change in words, for the history and undo, from the sheets as they were before it. */
function describeStructure(op: WorkbookOp, sheets: SheetView[]): Text {
  const sheet = 'sheet' in op ? sheets[op.sheet]?.info.name : undefined;
  switch (op.type) {
    case 'addSheet':
      return msg('sheets.history.added', { sheet: op.name });
    case 'renameSheet':
      return msg('sheets.history.renamed', { from: sheet, to: op.name });
    case 'deleteSheet':
      return msg('sheets.history.deleted', { sheet });
    case 'moveSheet':
      return msg('sheets.history.moved', { sheet });
    case 'hideSheet':
      return msg(op.hidden ? 'sheets.history.hidden' : 'sheets.history.unhidden', { sheet });
    case 'insertRows':
      return msg('store.history.insertedRows', { count: op.count });
    case 'deleteRows':
      return msg('store.history.deletedRows', { count: op.count });
    case 'insertCols':
      return msg('store.history.insertedColumns', { count: op.count });
    case 'deleteCols':
      return msg('store.history.deletedColumns', { count: op.count });
    case 'growTable':
      return msg('store.history.tableGrew', { table: op.table, count: op.rows });
    case 'move': {
      const range = rangeLabel(op.src);
      if (op.toSheet === op.sheet) return msg('store.history.moved', { range });
      const sheet = sheets[op.toSheet]?.info.name;
      return sheet === undefined ? msg('store.history.movedElsewhere', { range }) : msg('store.history.movedTo', { range, sheet });
    }
  }
}

/** The history's name for each kind of cell edit. */
const EDIT_ACTIONS: Record<EditOp['type'], MessageKey> = {
  input: 'store.history.edit',
  inputs: 'store.history.edit',
  batch: 'store.history.edit',
  pasteText: 'store.history.paste',
  copy: 'store.history.paste',
  fill: 'store.history.fill',
  series: 'fill.history.series',
  clear: 'store.history.clear',
  restore: 'store.history.restore',
  sort: 'store.history.sort',
  layout: 'store.history.layout',
  format: 'store.history.format',
  comment: 'store.history.comment',
  // Changes to names are always labelled with the name.
  name: 'store.history.edit',
  linkFormat: 'links.history.edit',
};

/** Has the browser download a file. */
function download(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

const rangeLabel = (r: Rect) => (r.r1 === r.r2 && r.c1 === r.c2 ? toA1(r.r1, r.c1) : `${toA1(r.r1, r.c1)}:${toA1(r.r2, r.c2)}`);

export class WorkbookStore {
  status: 'empty' | 'loading' | 'ready' | 'error' = 'empty';
  /** What is happening while a workbook loads. */
  stage: Text = '';
  /** Why the workbook could not be opened. */
  error?: Text;
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
  /** The latest things said to screen readers and not shown, newest last. */
  spoken: Spoken[] = [];
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
  /** Asking for a row's height, from the keyboard. */
  rowHeightFor?: number;
  /**
   * The cell whose note or comments are open to read or change, and where
   * to show them; `start` is what to add when the cell has neither.
   */
  commentFor?: { sheet: number; row: number; col: number; x: number; y: number; start: 'note' | 'thread' };
  /** A sheet tab's menu, while open: which sheet, and where. */
  sheetMenu?: { sheet: number; x: number; y: number };
  /** The sheet being renamed, while its dialog is open. */
  renameSheetFor?: number;
  /** Choosing a hidden sheet to show again. */
  unhideOpen = false;
  /** The workbook's structure is protected: its sheets can't be added, deleted, renamed, moved, hidden or unhidden. */
  structureLocked = false;
  /** The workbook's defined names, as they now are. */
  names: DefinedName[] = [];
  /** Defining or editing a name, while its dialog is open. */
  nameDialog?: NameDialog;
  /** Excel's Name Manager, while open. */
  nameManagerOpen = false;
  /** Asking for a link to give the selected cells. */
  linkDialog = false;
  /** Excel's Fill Series dialog, for the selection. */
  seriesDialog = false;
  /** Excel's Custom Lists dialog, opened from the Fill Series dialog. */
  customListsOpen = false;
  /** The user's lists couldn't be kept in the browser's storage, so they last until the page closes. */
  listsNotKept = false;
  /** Addresses the user has agreed to open while this workbook is open. */
  private readonly trustedLinks = new Set<string>();
  designerOpen = false;
  bannerDismissed = false;
  /** The schema banner (no schema matched, or the file was saved under another version), until the next file. */
  schemaBannerDismissed = false;
  highlightEdits = true;
  /** The user's override for exact numbers; undefined follows the schema or defaults. */
  exactNumbers?: boolean;
  schemas: WorkbookSchema[];
  schema?: SchemaController;
  schemaMismatch?: SchemaMismatch;
  /** Set when the file was last saved with another version of the schema now applied to it. */
  schemaNote?: Text;
  /** The schema being edited in the designer, kept while the designer is closed or picking cells. */
  schemaDraft?: WorkbookSchema;
  /** The designer is waiting for cells to be selected on the sheet, for a new table or a single cell. */
  designerPick?: 'table' | 'cell';
  /** The workbook's custom document properties, as opened. */
  properties: Record<string, string> = {};
  /** Changes whenever schema issues change, so the grid redraws its markers. */
  schemaVersion = 0;
  /** What was done, oldest first; hosts receive it translated, as HistoryEntry. */
  history: HistoryItem[] = [];
  /** The latest list of changed cells, refreshed while the Changes panel is open. */
  changeList: CellChange[] = [];
  private draftKey?: string;
  private draftOps: DraftOp[] = [];
  private draftTimer?: ReturnType<typeof setTimeout>;
  /** The draft being written, which reopening waits for. */
  private draftWrite?: Promise<unknown>;
  /** The file as opened, to open again if the worker fails. */
  private source?: { fileName: string; bytes: ArrayBuffer };
  /** Reopen with live recalculation off: building the engine is what failed. */
  private reopenWithoutRecalc = false;
  /** Waits for the engine's report after opening, in case building it hangs. */
  private calcTimer?: ReturnType<typeof setTimeout>;
  version = 0;
  private readonly options: StoreOptions;

  private worker?: Worker;
  private seq = 0;
  private pending = new Map<number, Pending>();
  /** Schemas the host gave, by id, with their code hooks. */
  private readonly hostSchemas = new Map<string, WorkbookSchema>();
  private listeners = new Set<() => void>();
  private clip?: Clip;
  private queue: Promise<unknown> = Promise.resolve();
  private toastSeq = 0;
  /** Set while the pointer or focus is on a message, so it stays until they leave. */
  private toastsHeld = false;
  /** What each undoable step did, said when it is undone or redone. */
  private readonly stepLabels = new WeakMap<object, Text>();
  /** The cell just typed in, until its checks have run: a problem with it is said at once. */
  private checking?: { sheet: number; row: number; col: number };
  private readonly listCache = new WeakMap<ListValidation, { version: number; options: string[] | undefined }>();
  private readonly formulaLists = new Map<string, FormulaList>();
  /** Bumped whenever cell values change, so formula-based lists are re-evaluated. */
  private dataVersion = 0;

  readonly features: EditorFeatures;

  constructor(options: StoreOptions = {}) {
    // A copy: changing the theme or user later doesn't touch the host's object.
    this.options = { ...options };
    this.schemas = [...(options.schemas ?? [])];
    for (const s of this.schemas) this.hostSchemas.set(s.id, s);
    this.features = { open: true, save: true, download: true, schemas: true, report: true, charts: true, format: true, csv: true, ...options.features };
    this.i18n = createI18n(options.language ?? this.locale, options.translations);
    this.t = translator(this.i18n);
  }

  /** The editor's own i18next instance. */
  readonly i18n: I18n;
  /** Translation functions, in the editor's current language. */
  readonly t: Translator;

  /**
   * Whether the editor is phone-sized (its own width, which in a host's layout
   * can be narrower than the window): the side panel then covers the sheet.
   */
  narrow = isNarrow();

  setNarrow(narrow: boolean): void {
    if (narrow === this.narrow) return;
    this.narrow = narrow;
    this.notify();
  }

  /** Changes the language of the editor's text while it is open. */
  setLanguage(language: string): void {
    void this.i18n.changeLanguage(language);
    // Schema messages name host-written text in the language they were made in; make them again.
    const ctl = this.schema;
    if (ctl) {
      if (this.schemaNote) this.schemaNote = versionNote(ctl.schema, this.properties, this.t.language);
      ctl.validate();
    }
    this.notify();
  }

  /** The theme tokens set through options or setTheme. */
  get theme(): EditorTheme | undefined {
    return this.options.theme;
  }

  /** Bumped when the theme changes, so the canvas grid reads its colours again. */
  themeVersion = 0;

  /** Changes the theme while the editor is open. */
  /** Changes the name recorded against each action in the change history. */
  setUser(user: string | undefined): void {
    this.options.user = user;
  }

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
    if (!this.hasSaveHandler) return this.t('toolbar.saveAsExcel');
    return this.options.saveLabel ? this.t.localized(this.options.saveLabel) : this.t('toolbar.save');
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
    this.stopWorker(new Error('Superseded'));
    this.listeners.clear();
  }

  /** Whether the file can be opened again after the editor stopped working. */
  get canReopen(): boolean {
    return this.status === 'error' && !!this.source && !this.worker;
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
          this.toast('info', msg('store.cf.tooManyCells', { count: res.result.skipped, sheet: view.info.name }));
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

  private get timeout(): number {
    return this.options.workerTimeout ?? WORKER_TIMEOUT;
  }

  /** A timer that declares the worker stopped if it is still running when the time is up. */
  private watchdog(worker: Worker): ReturnType<typeof setTimeout> | undefined {
    if (!this.timeout) return undefined;
    return setTimeout(() => this.workerFailed(worker, msg('store.worker.timedOut', { count: Math.max(1, Math.round(this.timeout / 1000)) })), this.timeout);
  }

  private request(req: Omit<WorkerRequest, 'id'> & { type: WorkerRequest['type'] }, transfer: Transferable[] = []): Promise<WorkerResponse> {
    const worker = this.worker;
    if (!worker) return Promise.reject(new TextError(this.status === 'error' && this.error ? this.error : msg('store.noWorkbook')));
    const id = ++this.seq;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject, timer: this.watchdog(worker) });
      worker.postMessage({ ...req, id } as WorkerRequest, transfer);
    });
  }

  private onMessage = (e: MessageEvent<WorkerResponse>) => {
    const res = e.data;
    if (res.type === 'calc') {
      clearTimeout(this.calcTimer);
      this.applyCalc(res.report);
      this.notify();
      return;
    }
    if (res.type === 'failed') {
      if (this.worker) this.workerFailed(this.worker, msg('store.worker.calcFailed', { reason: res.text }), true);
      return;
    }
    const p = this.pending.get(res.id);
    if (!p) return;
    this.pending.delete(res.id);
    clearTimeout(p.timer);
    if (res.type === 'error') p.reject(new TextError(res.text ?? res.message));
    else p.resolve(res);
    // The engine is built after the workbook is shown; its report should follow.
    if ((res.type === 'opened' || (res.type === 'restructured' && !res.report)) && this.worker) {
      clearTimeout(this.calcTimer);
      this.calcTimer = this.watchdog(this.worker);
    }
  };

  /** Creates a worker for one file, listening for it failing as well as for its answers. */
  private startWorker(): Worker {
    const worker = (this.options.createWorker ?? defaultWorker)();
    worker.onmessage = this.onMessage;
    // A worker that can't load, throws outside a request, or runs out of memory reports here.
    worker.onerror = (e: ErrorEvent) => {
      e.preventDefault?.();
      this.workerFailed(worker, msg('store.worker.stopped'));
    };
    worker.onmessageerror = () => this.workerFailed(worker, msg('store.worker.stopped'));
    return worker;
  }

  /** Ends the worker and fails every request still waiting on it. */
  private stopWorker(reason: Error): void {
    this.worker?.terminate();
    this.worker = undefined;
    clearTimeout(this.calcTimer);
    for (const p of this.pending.values()) {
      clearTimeout(p.timer);
      p.reject(reason);
    }
    this.pending.clear();
  }

  /**
   * The worker stopped working: it is ended, waiting requests fail, and the
   * editor shows what happened with an offer to reopen the file. Unsaved
   * edits are written to the draft now rather than after the usual pause.
   */
  private workerFailed(worker: Worker, reason: Text, calculating = false): void {
    if (worker !== this.worker) return;
    const error = new TextError(reason);
    this.stopWorker(error);
    if (this.draftKey && this.draftOps.length) {
      clearTimeout(this.draftTimer);
      this.draftWrite = saveDraft({ key: this.draftKey, fileName: this.fileName, updatedAt: Date.now(), editedCount: this.editedCount, ops: this.draftOps }).catch(() => undefined);
    }
    this.reopenWithoutRecalc = calculating;
    Object.assign(this, { status: 'error', error: reason, busy: false, stage: '' });
    this.queue = Promise.resolve();
    this.notify();
    this.options.onError?.(new Error(this.t.text(reason)));
  }

  /** Tells the host about an error the editor recovered from, such as one while drawing. */
  reportError(error: Error): void {
    this.options.onError?.(error);
  }

  /** Opens the file again after the editor stopped working, offering back unsaved edits. */
  async reopen(): Promise<void> {
    if (!this.canReopen || !this.source) return;
    await this.draftWrite;
    await this.load(this.source.fileName, this.source.bytes, this.reopenWithoutRecalc ? 'off' : undefined);
  }

  // ---------------------------------------------------------------------------
  // Opening and saving

  /**
   * Opens a workbook, or a CSV file as a workbook of one sheet: whether it
   * opened, or why not. The user can decline to discard edits to the workbook
   * already open; a file that can't be read leaves the editor showing why.
   * `csv: false` refuses CSV files (the editor's own Open button and drop
   * pass the `csv` feature).
   */
  async open(fileName: string, bytes: ArrayBuffer, { csv = true }: { csv?: boolean } = {}): Promise<OpenResult> {
    if (this.editedCount > 0 && !(await this.confirm({
      title: msg('store.open.discardTitle'),
      body: msg('store.open.discardBody', { fileName: this.fileName }),
      confirmLabel: msg('store.open.discardConfirm'),
    }))) {
      return { opened: false, cancelled: true };
    }
    if (!/\.xls[xm]$/i.test(fileName) && !(csv && CSV_FILE.test(fileName))) {
      const error = msg(csv ? 'store.open.notSupported' : 'store.open.notExcel');
      this.toast('error', error);
      return { opened: false, error };
    }
    if (this.draftKey && this.editedCount > 0) void deleteDraft(this.draftKey);
    return this.load(fileName, bytes);
  }

  /** Opens a file in a fresh worker; `recalc` overrides the option, for reopening after the engine failed. */
  private async load(fileName: string, bytes: ArrayBuffer, recalc?: OpenOptions['recalc']): Promise<OpenResult> {
    // Hash before the bytes move to the worker; the hash identifies this file's drafts.
    const started = performance.now();
    const key = this.options.drafts === false ? undefined : await fileKey(bytes).catch(() => undefined);
    this.stopWorker(new Error('Superseded'));
    // A fresh worker per file releases the previous workbook's memory.
    this.worker = this.startWorker();
    // The worker gets a copy, leaving the caller's buffer intact and one to reopen from.
    this.source = { fileName, bytes };
    this.reopenWithoutRecalc = false;
    Object.assign(this, {
      status: 'loading',
      stage: msg('store.open.reading', { fileName }),
      error: undefined,
      fileName,
      calc: undefined,
      undoStack: [],
      redoStack: [],
      editedCount: 0,
      bannerDismissed: false,
      schemaBannerDismissed: false,
      panel: null,
      clip: undefined,
      schema: undefined,
      schemaMismatch: undefined,
      sheetMenu: undefined,
      renameSheetFor: undefined,
      unhideOpen: false,
      history: [],
      changeList: [],
      names: [],
      nameDialog: undefined,
      nameManagerOpen: false,
      draftKey: key,
      draftOps: [],
      commentFor: undefined,
    });
    this.selections.clear();
    this.notify();
    try {
      const copy = bytes.slice(0);
      const res = await this.request(
        { type: 'open', bytes: copy, fileName, locale: this.locale || 'en-US', options: { limits: this.options.limits, macros: this.options.macros, activeContent: this.options.activeContent, recalc: recalc ?? this.options.recalc } } as Omit<WorkerRequest, 'id'>,
        [copy],
      );
      if (res.type !== 'opened') throw new Error('Unexpected response');
      const r = res.result;
      this.sheets = r.sheets.map((info, i) => new SheetView(info, r.grids[i]));
      for (const v of this.sheets) v.isUnlocked = (style) => !!this.styles[style]?.unlocked;
      this.names = r.names;
      this.styles = r.styles;
      this.findings = r.findings;
      this.date1904 = r.date1904;
      this.properties = r.properties ?? {};
      this.structureLocked = !!r.structureLocked;
      this.names = r.names ?? [];
      this.trustedLinks.clear();
      this.timings = r.timings;
      this.active = Math.max(0, r.sheets.findIndex((s) => s.state === 'visible'));
      this.status = 'ready';
      this.stage = '';
      this.cfRuns.clear();
      this.bindSchemas();
      this.refreshConditionalFormats(this.active);
      this.announce(msg('store.announce.opened', { fileName: this.fileName, count: this.sheets.length }));
    } catch (err) {
      if ((err as Error).message === 'Superseded') return { opened: false, cancelled: true };
      // A failed worker has already said so.
      if (this.status !== 'error') {
        this.status = 'error';
        this.error = errorText(err);
      }
    }
    this.notify();
    if (this.status !== 'ready') {
      const error = this.error ?? msg('store.noWorkbook');
      this.emit({ type: 'openFailed', fileName, reason: english.text(error) });
      return { opened: false, error };
    }
    this.emit({ type: 'opened', fileName, ms: Math.round(performance.now() - started), sheets: this.sheets.length });
    if (key) await this.offerDraft(key);
    return { opened: true };
  }

  /** Offers back edits left unsaved the last time this exact file was open. */
  private async offerDraft(key: string): Promise<void> {
    const draft = await loadDraft(key);
    if (!draft?.ops.length || this.draftKey !== key) return;
    const when = this.t.date(draft.updatedAt, { dateStyle: 'medium', timeStyle: 'short' });
    const answer = await this.ask({
      title: msg('store.draft.title'),
      body: msg('store.draft.body', { when, count: draft.editedCount }),
      confirmLabel: msg('store.draft.restore'),
      cancelLabel: msg('store.draft.discard'),
      danger: false,
    });
    if (answer !== 'confirm') {
      void deleteDraft(key);
      return;
    }
    // Cell edits replay as batches between structural changes, which replay on their own.
    let batch: EditOp[] = [];
    const flush = async () => {
      if (batch.length) await this.edit({ type: 'batch', ops: batch }, 'undo', { label: msg('store.history.restoredDraft'), replay: true });
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

  /** The workbook as it is now, as onSave would receive it, without running save checks or onSave. */
  async snapshot(): Promise<SavedWorkbook | undefined> {
    await this.settled();
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
        history: this.history.map((h) => ({ ...h, action: this.t.text(h.action) })),
        editedCount: this.editedCount,
        outOfDate: outOfDateCells(this.sheets),
        ...(ctl && summary
          ? {
              schema: { id: ctl.schema.id, name: this.t.localized(ctl.schema.name), version: ctl.schema.version, errors: summary.errors, warnings: summary.warnings },
              values: ctl.values(),
            }
          : {}),
        ...this.readers(),
      };
    } catch (err) {
      this.toast('error', msg('store.save.failed', { reason: errorText(err) }));
      return undefined;
    } finally {
      this.busy = false;
      this.notify();
    }
  }

  /**
   * The Save button: schema checks, an optional review of every change,
   * then the host's onSave or a download. Resolves with what happened; the
   * editor has already told the user.
   */
  async save(): Promise<SaveResult> {
    const started = performance.now();
    const result = await this.saveSteps();
    this.emit({ type: 'saved', result, ms: Math.round(performance.now() - started), editedCount: this.editedCount });
    return result;
  }

  private async saveSteps(): Promise<SaveResult> {
    if (this.schema) {
      this.busy = true;
      this.notify();
      const check = await this.schema.checkBeforeSave();
      this.busy = false;
      this.notify();
      if (!check.ok) {
        if (!check.canOverride) {
          this.set('panel', 'issues');
          this.toast('error', msg('store.save.blocked', { reason: check.message }));
          return 'blocked';
        }
        const go = await this.confirm({
          title: msg('store.save.withErrorsTitle'),
          body: msg('store.save.withErrorsBody', { reason: check.message }),
          confirmLabel: msg('store.save.withErrorsConfirm'),
        });
        if (!go) return 'cancelled';
      }
    }
    await this.settled();
    const outOfDate = outOfDateCells(this.sheets).length;
    if (
      outOfDate &&
      !(await this.confirm({
        title: msg('store.save.outOfDateTitle'),
        body: msg('store.save.outOfDateBody', { count: outOfDate }),
        confirmLabel: msg('store.save.outOfDateConfirm'),
      }))
    )
      return 'cancelled';
    const saved = await this.snapshot();
    if (!saved) return 'failed';
    if (this.numbers.reviewBeforeSave && saved.changes.length) {
      const LIMIT = 300;
      const answer = await this.ask({
        title: msg('store.save.reviewTitle', { count: saved.changes.length }),
        body: msg('store.save.reviewBody'),
        confirmLabel: msg('store.save.reviewConfirm'),
        danger: false,
        details: {
          header: [msg('store.save.reviewCell'), msg('store.save.reviewWas'), msg('store.save.reviewNow')],
          rows: saved.changes
            .slice(0, LIMIT)
            .map((c) => [
              `${c.sheet}!${c.ref}`,
              c.before,
              c.comment
                ? c.after
                  ? this.t('store.save.commentChanged', { change: this.t.text(c.comment), text: c.after })
                  : this.t.text(c.comment)
                : c.formatOnly
                  ? this.t('store.save.formatChanged', { value: c.after })
                  : c.after,
            ]),
          more: Math.max(0, saved.changes.length - LIMIT),
        },
      });
      if (answer !== 'confirm') return 'cancelled';
    }
    if (this.options.onSave) {
      this.busy = true;
      this.notify();
      try {
        await this.options.onSave(saved);
        this.afterSave();
        this.toast('info', msg('store.save.saved', { fileName: saved.fileName }));
        return 'saved';
      } catch (err) {
        if (err instanceof SaveConflictError || (err as Error)?.name === 'SaveConflictError') {
          const own = (err as Error).message;
          const answer = await this.ask({
            title: msg('store.conflict.title'),
            body: msg('store.conflict.body', { reason: !own || own === SAVE_CONFLICT ? msg('store.conflict.reason') : own }),
            confirmLabel: msg('store.conflict.download'),
            cancelLabel: msg('store.conflict.keepEditing'),
            danger: false,
          });
          if (answer === 'confirm') {
            const [, name, extension] = /^(.*?)(\.xls[xm])?$/i.exec(saved.fileName)!;
            this.downloadFile({ ...saved, fileName: this.t('store.conflict.copyName', { name, extension: extension ?? '' }) });
          }
        } else {
          this.toast('error', msg('store.save.notSaved', { reason: errorText(err) }));
        }
        return 'failed';
      } finally {
        this.busy = false;
        this.notify();
      }
    }
    this.downloadFile(saved);
    this.afterSave();
    this.toast(
      'info',
      saved.editedCount
        ? msg('store.save.savedChanges', { fileName: saved.fileName, count: saved.editedCount })
        : msg('store.save.savedUnchanged', { fileName: saved.fileName }),
    );
    return 'downloaded';
  }

  private downloadFile(saved: SavedWorkbook): void {
    const type = /\.xlsm$/i.test(saved.fileName)
      ? 'application/vnd.ms-excel.sheet.macroEnabled.12'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    download(new Blob([saved.bytes as BlobPart], { type }), saved.fileName);
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
    const saved = await this.snapshot();
    if (!saved) return;
    this.downloadFile(saved);
    if (!this.hasSaveHandler) this.afterSave();
    this.toast('info', msg('store.save.downloaded', { fileName: saved.fileName }));
  }

  /**
   * The sheet shown, as CSV text (see sheetCsv) and the name to give the
   * file, once pending edits are in; undefined when it isn't a worksheet.
   * Options override the `csvExport` option.
   */
  async csv(options: CsvOptions = {}): Promise<{ text: string; fileName: string; sheet: string } | undefined> {
    await this.settled();
    const view = this.view;
    if (this.status !== 'ready' || !view || view.info.kind !== 'worksheet') return undefined;
    const text = sheetCsv(view, this.styles, this.date1904, this.locale, { ...this.options.csvExport, ...options });
    return { text, fileName: csvFileName(this.fileName, view.info.name, this.sheets.length), sheet: view.info.name };
  }

  /**
   * Downloads the sheet shown as a CSV file, UTF-8 with a byte order mark so
   * that Excel reads it as UTF-8. It is an export: it doesn't count as saving.
   */
  async exportCsv(options: CsvOptions = {}): Promise<void> {
    const csv = await this.csv(options);
    if (!csv) return;
    download(new Blob(['\uFEFF' + csv.text], { type: 'text/csv;charset=utf-8' }), csv.fileName);
    this.toast('info', msg('store.csv.downloaded', { fileName: csv.fileName, sheet: csv.sheet }));
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
    opts: { label?: Text; replay?: boolean; announce?: Text } = {},
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
        const changed = r.undo.length > 0 || !!r.layout?.undo.length || !!r.noteUndo?.length || !!r.names?.undo.length;
        if (record === 'undo' && changed) {
          this.undoStack.push(r);
          this.redoStack = [];
          this.stepLabels.set(r, opts.label ?? msg(EDIT_ACTIONS[op.type]));
          // What was done is said, except typing, which is heard as it is typed.
          const said = opts.announce ?? opts.label;
          if (said && !opts.replay) this.announce(said);
        } else if (record === 'merge' && changed) {
          const last = this.undoStack.pop();
          if (last && isStructure(last)) this.undoStack.push(last, r);
          else {
            const merged = last ? mergeEditResults([last, r]) : r;
            const label = last && this.stepLabels.get(last);
            if (label) this.stepLabels.set(merged, label);
            this.undoStack.push(merged);
          }
        }
        if (changed) {
          this.recordHistory(op, record, r, opts.label);
          // A new comment's id and time are made in the worker: the draft keeps the comment as made.
          this.draftOps.push(op.type === 'comment' ? { type: 'restore', cells: [], notes: r.noteRedo } : op);
          this.scheduleDraft();
          if (this.panel === 'changes') void this.refreshChanges();
        }
        if (ctl) {
          if (record === 'undo' && before.length && !opts.replay) void ctl.afterEdit(before);
          else ctl.validate();
        }
        return r;
      } catch (err) {
        if ((err as Error).message !== 'Superseded') this.toast('error', errorText(err));
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

  private recordHistory(op: EditOp, record: 'undo' | 'none' | 'merge', r: EditResult, label?: Text): void {
    const action = label ?? msg(record === 'merge' ? 'store.history.businessLogic' : EDIT_ACTIONS[op.type]);
    const cells = [...r.redo, ...(r.noteRedo ?? [])];
    const refs = cells.slice(0, 5).map((c) => `${this.sheets[c.sheet]?.info.name ?? '?'}!${toA1(c.row, c.col)}`);
    this.history.push({ at: Date.now(), user: this.options.user, action, cells: cells.length, refs });
  }

  private applyResult(r: EditResult): void {
    this.dataVersion++;
    if (r.undo.length || r.layout?.applied.length || r.noteUndo?.length || r.names?.applied.length) this.changeSeq++;
    if (r.names) this.names = r.names.list;
    for (const c of r.layout?.applied ?? []) {
      this.sheets[c.sheet]?.applyLayout(c.op);
      this.colVersion++;
    }
    for (const s of r.newStyles) this.styles[s.index] = s.style;
    for (const u of r.updates) this.sheets[u.sheet]?.apply(u);
    for (const n of r.noteRedo ?? []) this.sheets[n.sheet]?.setNote(n.row, n.col, n.note);
    this.editedCount = r.editedCount;
    if (r.rejected) this.toast('error', r.rejected);
    if (r.skipped) {
      const reason = r.skippedReasons[0] ?? '';
      this.toast('warn', r.skipped === 1 ? reason : msg('store.edit.skipped', { count: r.skipped, reason }));
    }
    for (const w of r.warnings) this.toast('warn', w);
    this.refreshConditionalFormats(this.active);
  }

  async undo(): Promise<void> {
    await this.queue;
    const entry = this.undoStack.pop();
    if (!entry) return;
    if (isStructure(entry)) {
      const sheets = entry.sheets && { map: invertIndexMap(entry.sheets.map, this.sheets.length), active: entry.sheets.before };
      if (await this.swapStructure('undoStructure', msg('store.history.undoStep', { action: entry.structure }), sheets)) {
        this.redoStack.push(entry);
        this.announce(msg('store.announce.undone', { action: entry.structure }));
      } else {
        this.undoStack = [];
      }
      this.notify();
      return;
    }
    this.showSheet(entry.undo[0]?.sheet ?? entry.noteUndo?.[0]?.sheet ?? entry.layout?.undo[0]?.sheet);
    const notes = entry.noteUndo?.length ? { notes: entry.noteUndo } : {};
    await this.edit({ type: 'restore', cells: entry.undo, ...notes, layout: entry.layout?.undo, names: entry.names?.undo }, 'none', { label: msg('store.history.undo') });
    this.redoStack.push(entry);
    this.announce(msg('store.announce.undone', { action: this.stepLabels.get(entry) ?? msg('store.history.edit') }));
    this.notify();
  }

  async redo(): Promise<void> {
    await this.queue;
    const entry = this.redoStack.pop();
    if (!entry) return;
    if (isStructure(entry)) {
      const sheets = entry.sheets && { map: entry.sheets.map, active: entry.sheets.after };
      if (await this.swapStructure('redoStructure', msg('store.history.redoStep', { action: entry.structure }), sheets)) {
        this.undoStack.push(entry);
        this.announce(msg('store.announce.redone', { action: entry.structure }));
      }
      this.notify();
      return;
    }
    this.showSheet(entry.redo[0]?.sheet ?? entry.noteRedo?.[0]?.sheet ?? entry.layout?.applied[0]?.sheet);
    const notes = entry.noteRedo?.length ? { notes: entry.noteRedo } : {};
    await this.edit({ type: 'restore', cells: entry.redo, ...notes, layout: entry.layout?.applied, names: entry.names?.applied }, 'none', { label: msg('store.history.redo') });
    this.undoStack.push(entry);
    this.announce(msg('store.announce.redone', { action: this.stepLabels.get(entry) ?? msg('store.history.edit') }));
    this.notify();
  }

  // ---------------------------------------------------------------------------
  // Structural changes: inserting and deleting rows and columns, moving cells,
  // and adding, renaming, deleting, moving, hiding and unhiding sheets

  /**
   * Inserts or deletes rows or columns, moves cells, or changes the sheets,
   * with every reference to what changed following it. The worker rebuilds
   * the workbook, so all the sheets are reloaded here. Resolves false if it
   * was refused.
   */
  restructure(op: WorkbookOp, opts: { replay?: boolean } = {}): Promise<boolean> {
    const run = async () => {
      if (this.status !== 'ready') return false;
      this.busy = true;
      this.notify();
      try {
        const label = describeStructure(op, this.sheets);
        const before = this.active;
        const res = await this.request({ type: 'structure', op } as Omit<WorkerRequest, 'id'>);
        if (res.type !== 'restructured') return false;
        let sheets: SheetMove | undefined;
        if (isSheetOp(op)) {
          const map = sheetIndexMap(op, this.sheets.length);
          sheets = { map, before, after: this.activeAfter(op, map, res.result.sheets) };
          this.adoptWorkbook(res, { map, active: sheets.after });
        } else {
          this.adoptWorkbook(res);
        }
        this.undoStack.push({ structure: label, sheets });
        this.redoStack = [];
        this.history.push({ at: Date.now(), user: this.options.user, action: label, cells: 0, refs: [] });
        this.draftOps.push({ type: 'structure', op });
        this.scheduleDraft();
        if (!opts.replay) {
          this.selectAfter(op);
          this.announce(label);
        }
        for (const w of res.warnings) this.toast('warn', w);
        if (this.panel === 'changes') void this.refreshChanges();
        return true;
      } catch (err) {
        if ((err as Error).message !== 'Superseded') this.toast('error', errorText(err));
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

  /** The sheet shown after a change to the sheets, as Excel shows it. */
  private activeAfter(op: SheetOp, map: (number | null)[], sheets: { state: SheetView['info']['state'] }[]): number {
    if (op.type === 'addSheet') return op.at;
    if (op.type === 'hideSheet' && !op.hidden) return op.sheet;
    const now = map[this.active];
    // The sheet deleted or hidden was the one shown: the next one along is.
    if (now === null || sheets[now]?.state !== 'visible') return nearestVisible(sheets, now ?? op.sheet);
    return now;
  }

  private async swapStructure(type: 'undoStructure' | 'redoStructure', label: Text, sheets?: { map: (number | null)[]; active: number }): Promise<boolean> {
    try {
      this.busy = true;
      this.notify();
      const res = await this.request({ type } as Omit<WorkerRequest, 'id'>);
      if (res.type !== 'restructured') return false;
      this.adoptWorkbook(res, sheets);
      this.history.push({ at: Date.now(), user: this.options.user, action: label, cells: 0, refs: [] });
      if (this.panel === 'changes') void this.refreshChanges();
      return true;
    } catch (err) {
      this.toast('error', errorText(err));
      return false;
    } finally {
      this.busy = false;
    }
  }

  /**
   * Replaces every sheet's view with the worker's rebuilt workbook, keeping
   * the active sheet and selections; after a change to the sheets, `sheets`
   * says where each went and which to show.
   */
  private adoptWorkbook(res: Extract<WorkerResponse, { type: 'restructured' }>, sheets?: { map: (number | null)[]; active: number }): void {
    const r = res.result;
    if (sheets) {
      const selections = [...this.selections].flatMap(([i, sel]) => (sheets.map[i] == null ? [] : [[sheets.map[i], sel] as const]));
      this.selections = new Map(selections);
      this.active = sheets.active;
    }
    this.structureLocked = !!r.structureLocked;
    this.sheets = r.sheets.map((info, i) => new SheetView(info, r.grids[i]));
    for (const v of this.sheets) v.isUnlocked = (style) => !!this.styles[style]?.unlocked;
    this.names = r.names;
    this.styles = r.styles;
    this.date1904 = r.date1904;
    this.names = r.names ?? [];
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
    // The cell may have moved.
    this.commentFor = undefined;
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
      this.toast('error', errorText(err));
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
    this.emit({ type: 'calculated', ms: report.buildMs, formulas: report.formulaCells, live: report.liveCells, notCalculated: report.unsupported + report.mismatched });
    this.findings = [...this.findings.filter((f) => f.kind !== 'formulaEngine'), ...report.findings];
    for (const u of report.updates) this.sheets[u.sheet]?.apply(u);
    this.dataVersion++;
    this.schema?.validate();
    this.refreshConditionalFormats(this.active);
  }

  /** Selects what a structural change made or left, as Excel does. */
  private selectAfter(op: WorkbookOp): void {
    // A table growing under the cell being typed into leaves the selection where it is; a sheet keeps its own.
    if (op.type === 'growTable' || isSheetOp(op)) return;
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
  // Sheets

  /** What a new sheet is called: "Sheet" in the editor's language, numbered as Excel numbers it. */
  get newSheetName(): string {
    return nextSheetName(this.t('sheets.defaultName'), this.sheets.map((s) => s.info.name));
  }

  /** Hidden sheets that can be shown again: not those hidden so only Excel's Visual Basic editor can show them. */
  get hiddenSheets(): number[] {
    return this.sheets.flatMap((s, i) => (s.info.state === 'hidden' ? [i] : []));
  }

  /**
   * Why Excel would refuse a change to the sheets, or the schema that
   * applies would: it finds its sheets by name, so a sheet it checks keeps a
   * name it matches, and one it needs can't be deleted.
   */
  sheetProblem(op: SheetOp): Text | undefined {
    const excel = sheetOpProblem(op, this.sheets.map((s) => s.info), this.structureLocked);
    if (excel) return excel;
    const ctl = this.schema;
    const bound = ctl && 'sheet' in op ? ctl.binding.sheets.find((b) => b.sheet === op.sheet) : undefined;
    if (!ctl || !bound) return undefined;
    if (op.type === 'renameSheet' && !textMatches(bound.schema.name, op.name)) {
      return msg('sheets.refuse.schemaRename', { schema: ctl.schema.name, expected: describeMatch(bound.schema.name) });
    }
    if (op.type === 'deleteSheet' && !bound.schema.optional) return msg('sheets.refuse.schemaDelete', { schema: ctl.schema.name });
    return undefined;
  }

  /** Makes a change to the sheets, or shows why it can't be made. */
  private changeSheets(op: SheetOp): Promise<boolean> {
    const problem = this.sheetProblem(op);
    if (!problem) return this.restructure(op);
    this.toast('error', problem);
    return Promise.resolve(false);
  }

  /** Adds an empty sheet at `at`, after the active one unless told otherwise, and shows it. */
  addSheet(at = this.active + 1): Promise<boolean> {
    return this.changeSheets({ type: 'addSheet', at, name: this.newSheetName });
  }

  renameSheet(sheet: number, name: string): Promise<boolean> {
    if (name === this.sheets[sheet]?.info.name) return Promise.resolve(true);
    return this.changeSheets({ type: 'renameSheet', sheet, name });
  }

  /** Deletes a sheet, asking first unless it is empty, as Excel does. */
  async deleteSheet(sheet: number): Promise<boolean> {
    const op: SheetOp = { type: 'deleteSheet', sheet };
    const problem = this.sheetProblem(op);
    if (problem) {
      this.toast('error', problem);
      return false;
    }
    const view = this.sheets[sheet];
    const empty = view.info.kind === 'worksheet' && view.extent().maxRow < 0 && !view.info.objects.length;
    if (!empty && !(await this.confirm({ title: msg('sheets.delete.title', { sheet: view.info.name }), body: msg('sheets.delete.body'), confirmLabel: msg('sheets.delete.confirm') }))) {
      return false;
    }
    return this.restructure(op);
  }

  /** Moves a sheet so it ends up at position `to`. */
  moveSheet(sheet: number, to: number): Promise<boolean> {
    if (to === sheet) return Promise.resolve(false);
    return this.changeSheets({ type: 'moveSheet', sheet, to });
  }

  /** Moves a sheet past the next visible sheet to its left (-1) or right (1). */
  moveSheetBy(sheet: number, step: -1 | 1): Promise<boolean> {
    const visible = this.sheets.flatMap((s, i) => (s.info.state === 'visible' ? [i] : []));
    const other = visible[visible.indexOf(sheet) + step];
    return other === undefined ? Promise.resolve(false) : this.moveSheet(sheet, other);
  }

  /** Moves a sheet to just before the sheet now at `before` (the number of sheets for the end), as dropping its tab there does. */
  dropSheet(sheet: number, before: number): Promise<boolean> {
    return this.moveSheet(sheet, before > sheet ? before - 1 : before);
  }

  hideSheet(sheet: number): Promise<boolean> {
    return this.changeSheets({ type: 'hideSheet', sheet, hidden: true });
  }

  /** Shows a hidden sheet again, and makes it the active one. */
  unhideSheet(sheet: number): Promise<boolean> {
    return this.changeSheets({ type: 'hideSheet', sheet, hidden: false });
  }

  // ---------------------------------------------------------------------------
  // Layout: sizes, hidden rows and columns, frozen panes and merged cells

  /** Changes the active sheet's layout as one undoable step, after `before` (edits to cells that go with it). */
  private changeLayout(op: LayoutOp, label: Text, before: EditOp[] = [], announce?: Text): Promise<EditResult | undefined> {
    const ops: EditOp[] = [...before, { type: 'layout', sheet: this.active, op }];
    return this.edit(ops.length === 1 ? ops[0] : { type: 'batch', ops }, 'undo', { label, announce });
  }

  /**
   * The rows (or columns) the selection spans. A whole column selected
   * spans every row; only those to just past the used range count.
   */
  private selectedLines(axis: 'rows' | 'cols'): number[] {
    const view = this.view;
    const sel = this.selection;
    const limit = axis === 'rows' ? Math.max(view?.info.maxRow ?? 0, sel.row) + 200 : Math.max(view?.info.maxCol ?? 0, sel.col) + 50;
    const lines = new Set<number>();
    for (const r of sel.rects.length ? sel.rects : [{ r1: sel.row, r2: sel.row, c1: sel.col, c2: sel.col }]) {
      const [a, b] = axis === 'rows' ? [r.r1, r.r2] : [r.c1, r.c2];
      for (let i = a; i <= Math.min(b, limit); i++) lines.add(i);
    }
    return [...lines].sort((x, y) => x - y);
  }

  /** Hides the selected rows or columns. */
  async hideLines(axis: 'rows' | 'cols'): Promise<void> {
    const lines = this.selectedLines(axis);
    const op: LayoutOp = { kind: 'hidden', axis, lines: Object.fromEntries(lines.map((l) => [l, true])) };
    await this.changeLayout(op, msg(axis === 'rows' ? 'store.history.hideRows' : 'store.history.hideColumns', { count: lines.length }));
  }

  /**
   * Shows the hidden rows or columns in the selection, and those hidden
   * just before or after it, so rows hidden at the top of a sheet, which
   * can't be selected around, can still be shown.
   */
  async unhideLines(axis: 'rows' | 'cols'): Promise<void> {
    const view = this.view;
    if (!view) return;
    const hidden = new Set(axis === 'rows' ? view.info.hiddenRows : view.info.hiddenCols);
    const selected = this.selectedLines(axis);
    let a = selected[0];
    let b = selected[selected.length - 1];
    while (hidden.has(a - 1)) a--;
    while (hidden.has(b + 1)) b++;
    const lines = [...hidden].filter((l) => l >= a && l <= b);
    if (!lines.length) return;
    const op: LayoutOp = { kind: 'hidden', axis, lines: Object.fromEntries(lines.map((l) => [l, false])) };
    await this.changeLayout(op, msg(axis === 'rows' ? 'store.history.unhideRows' : 'store.history.unhideColumns', { count: lines.length }));
  }

  /** Whether there are hidden rows or columns that unhideLines would show. */
  hasHiddenNear(axis: 'rows' | 'cols'): boolean {
    const view = this.view;
    const hidden = axis === 'rows' ? view?.info.hiddenRows : view?.info.hiddenCols;
    if (!hidden?.length) return false;
    const selected = this.selectedLines(axis);
    const a = selected[0] - 1;
    const b = selected[selected.length - 1] + 1;
    return hidden.some((l) => l >= a && l <= b);
  }

  /**
   * Freezes the rows above and the columns left of the active cell, the top
   * row, or the first column, so they stay in view; "none" unfreezes.
   */
  async freezePanes(at: 'cell' | 'topRow' | 'firstColumn' | 'none'): Promise<void> {
    const { row, col } = this.selection;
    const [rows, cols] = at === 'cell' ? [row, col] : at === 'topRow' ? [1, 0] : at === 'firstColumn' ? [0, 1] : [0, 0];
    await this.changeLayout({ kind: 'freeze', rows, cols }, msg(rows || cols ? 'store.history.freeze' : 'store.history.unfreeze'));
  }

  /** The selection, grown to take in any merged cells it covers part of, as Excel's selection does. */
  private mergeTarget(): Rect | undefined {
    const view = this.view;
    const r = this.selection.rects[0];
    if (!view || !r) return undefined;
    const rect = { ...r };
    for (let grew = true; grew; ) {
      grew = false;
      for (const m of view.info.merges) {
        if (m.r1 > rect.r2 || m.r2 < rect.r1 || m.c1 > rect.c2 || m.c2 < rect.c1) continue;
        if (m.r1 < rect.r1 || m.r2 > rect.r2 || m.c1 < rect.c1 || m.c2 > rect.c2) {
          Object.assign(rect, { r1: Math.min(rect.r1, m.r1), r2: Math.max(rect.r2, m.r2), c1: Math.min(rect.c1, m.c1), c2: Math.max(rect.c2, m.c2) });
          grew = true;
        }
      }
    }
    return rect;
  }

  /**
   * Merges the selected cells into one. As in Excel, only the top-left
   * cell's value stays, so it asks first when others hold something.
   */
  async mergeCells(): Promise<void> {
    const view = this.view;
    const rect = this.mergeTarget();
    if (!view || !rect || (rect.r1 === rect.r2 && rect.c1 === rect.c2)) return;
    if (rectCells(rect) > 1_000_000) return this.toast('warn', msg('store.merge.tooLarge'));
    let lost = 0;
    view.forEach(rect, (row, col, rec) => {
      if ((row !== rect.r1 || col !== rect.c1) && rec.t !== VT.None) lost++;
    });
    if (lost && !(await this.confirm({ title: msg('store.merge.title'), body: msg('store.merge.body', { count: lost }), confirmLabel: msg('store.merge.confirm') }))) return;
    const others: Rect[] = [
      { r1: rect.r1, r2: rect.r1, c1: rect.c1 + 1, c2: rect.c2 },
      { r1: rect.r1 + 1, r2: rect.r2, c1: rect.c1, c2: rect.c2 },
    ].filter((r) => r.r1 <= r.r2 && r.c1 <= r.c2);
    const inside = view.info.merges.filter((m) => m.r1 >= rect.r1 && m.r2 <= rect.r2 && m.c1 >= rect.c1 && m.c2 <= rect.c2);
    const clear: EditOp[] = lost ? [{ type: 'clear', sheet: this.active, rects: others }] : [];
    await this.changeLayout({ kind: 'merges', add: [rect], remove: inside }, msg('store.history.merge'), clear);
  }

  /** Splits merged cells in the selection back into their cells. */
  async unmergeCells(): Promise<void> {
    const remove = this.mergesInSelection();
    if (remove.length) await this.changeLayout({ kind: 'merges', add: [], remove }, msg('store.history.unmerge'));
  }

  // ---------------------------------------------------------------------------
  // Formatting

  /** The active cell's format, which the format bar shows. */
  get activeStyle(): DisplayStyle {
    const view = this.view;
    const { row, col } = this.selection;
    return this.styles[view ? view.styleAt(row, col) : 0] ?? this.styles[0] ?? DEFAULT_STYLE;
  }

  /** Changes the selected cells' formats, as one undoable step. */
  async formatCells(change: FormatChange, opts: { borders?: { where: BorderPlacement; edge?: BorderEdge }; clear?: boolean } = {}): Promise<void> {
    if (!this.view) return;
    const rects = this.selection.rects.length ? this.selection.rects : [{ r1: this.selection.row, c1: this.selection.col, r2: this.selection.row, c2: this.selection.col }];
    const label = msg(opts.clear ? 'store.history.clearFormats' : 'store.history.format');
    await this.edit({ type: 'format', sheet: this.active, rects, change, ...opts }, 'undo', { label });
  }

  /** Turns bold, italic, underline or strikethrough on for the selection, or off when the active cell has it, as Excel does. */
  toggleFormat(key: 'bold' | 'italic' | 'underline' | 'strike'): Promise<void> {
    return this.formatCells({ [key]: !this.activeStyle[key] });
  }

  /** One more or one fewer decimal place in the active cell's number format, for the whole selection. */
  stepDecimals(delta: 1 | -1): Promise<void> {
    return this.formatCells({ numFmt: stepDecimals(this.activeStyle.numFmt, delta) });
  }

  /** Merged cells the selection covers some of. */
  mergesInSelection(): Rect[] {
    const view = this.view;
    const rects = this.selection.rects;
    return view?.info.merges.filter((m) => rects.some((r) => m.r1 <= r.r2 && m.r2 >= r.r1 && m.c1 <= r.c2 && m.c2 >= r.c1)) ?? [];
  }

  // ---------------------------------------------------------------------------
  // Hyperlinks

  /** Whether links are followed here, rather than only shown. */
  get followsLinks(): boolean {
    return this.options.links !== 'off';
  }

  /** The link on a cell of the active sheet: the sheet's own link there, or a HYPERLINK formula's. */
  linkAt(row = this.selection.row, col = this.selection.col): CellLink | undefined {
    const view = this.view;
    return view ? cellLink(view, row, col) : undefined;
  }

  /** The selected ranges, or the active cell, as links are given to them. */
  private linkRects(): Rect[] {
    const sel = this.selection;
    return sel.rects.length ? sel.rects : [{ r1: sel.row, c1: sel.col, r2: sel.row, c2: sel.col }];
  }

  /** How many of the sheet's links the selection covers some of. */
  linksInSelection(): number {
    const view = this.view;
    return view ? replaceLinks(view.info.links, this.linkRects()).remove.length : 0;
  }

  /**
   * Gives the selected cells a link, replacing any they had, as one undoable
   * step. `text`, when it differs from what the active cell shows, goes into
   * that cell first, as Excel's "Text to display" does; and as in Excel, the
   * cells take its Hyperlink look.
   */
  async setLink(link: LinkFields, text?: string): Promise<void> {
    const view = this.view;
    if (!view) return;
    const rects = this.linkRects();
    const change = replaceLinks(view.info.links, rects, link);
    const { row, col } = this.selection;
    const before: EditOp[] = [];
    if (text !== undefined && text !== this.shownText(view, row, col)) {
      // Shown as written: text starting "=" stays text rather than becoming a formula.
      const typed = text.startsWith('=') ? `'${text}` : text;
      const prepared = this.schema?.prepareInput(this.active, row, col, typed) ?? { text: typed };
      if ('reject' in prepared) return this.toast('warn', prepared.reject);
      before.push({ type: 'input', sheet: this.active, row, col, text: prepared.text });
    }
    // A protected sheet refuses the link; its cells keep their look.
    if (!view.info.protected) before.push({ type: 'linkFormat', sheet: this.active, rects, on: true });
    await this.changeLayout({ kind: 'links', ...change }, msg(change.remove.length ? 'links.history.edit' : 'links.history.insert'), before);
  }

  /**
   * Takes the links off the selected cells, which keep what they hold, as
   * one undoable step. As Excel's Remove Hyperlink does, the cells that had
   * a link lose its look.
   */
  async removeLinks(): Promise<void> {
    const view = this.view;
    const rects = this.linkRects();
    const change = view && replaceLinks(view.info.links, rects);
    if (!view || !change?.remove.length) return;
    const linked = change.remove.flatMap((l) =>
      rects
        .map((r) => ({ r1: Math.max(r.r1, l.rect.r1), c1: Math.max(r.c1, l.rect.c1), r2: Math.min(r.r2, l.rect.r2), c2: Math.min(r.c2, l.rect.c2) }))
        .filter((r) => r.r1 <= r.r2 && r.c1 <= r.c2),
    );
    const before: EditOp[] = view.info.protected ? [] : [{ type: 'linkFormat', sheet: this.active, rects: linked, on: false }];
    await this.changeLayout({ kind: 'links', ...change }, msg('links.history.remove'), before);
  }

  /**
   * Follows the link on a cell: to its place in the workbook, or to a web
   * page or email in a new tab (with no access back to this page), asking the
   * first time for each address, since the workbook is untrusted. Links to
   * anything else are refused, saying why. A host's onLinkClick takes web
   * pages and emails instead.
   */
  async followLink(row = this.selection.row, col = this.selection.col): Promise<void> {
    const view = this.view;
    const link = view && cellLink(view, row, col);
    if (!view || !link) return this.toast('info', msg('links.follow.none'));
    if (!this.followsLinks) return this.toast('info', msg('links.follow.off'));
    const to = linkDestination(link);
    if (!to) return this.toast('warn', msg('links.follow.unknown'));
    if (to.kind === 'refused') return this.toast('warn', msg('links.follow.refused', { address: to.address }));
    if (to.kind === 'place') {
      if (!this.goToPlace(to.location)) this.toast('warn', msg('links.follow.noPlace', { location: to.location }));
      return;
    }
    const host = this.options.onLinkClick;
    if (host) {
      try {
        if ((await host({ url: to.url, kind: to.kind, sheet: view.info.name, ref: toA1(row, col) })) !== true) return;
      } catch (err) {
        return this.toast('error', errorText(err));
      }
    }
    if (!this.trustedLinks.has(to.url)) {
      const email = to.kind === 'email';
      const ok = await this.confirm({
        title: msg(email ? 'links.confirm.emailTitle' : 'links.confirm.title'),
        body: msg(email ? 'links.confirm.emailBody' : 'links.confirm.body'),
        confirmLabel: msg(email ? 'links.confirm.email' : 'links.confirm.open'),
        danger: false,
        details: { header: [msg('links.confirm.address')], rows: [[to.url]] },
      });
      if (!ok) return;
      this.trustedLinks.add(to.url);
    }
    globalThis.open(to.url, '_blank', 'noopener,noreferrer');
  }

  /** Goes to a place in the workbook a link names: false when it isn't there, or is on a sheet that isn't shown. */
  private goToPlace(location: string): boolean {
    const place = resolvePlace(location, this.sheets.map((v) => v.info.name), this.names, this.active);
    const view = place && this.sheets[place.sheet];
    if (!place || !view || view.info.kind !== 'worksheet' || view.info.state !== 'visible' || place.rect.r2 >= MAX_ROWS || place.rect.c2 >= MAX_COLS) return false;
    if (place.sheet !== this.active) this.setActive(place.sheet);
    this.revealFocus = true;
    this.revealSeq++;
    this.setSelection({ row: place.rect.r1, col: place.rect.c1, rects: [place.rect] });
    return true;
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
      const r1 = rowHas(rect.r1 - 1) ? rect.r1 - 1 : rect.r1;
      const r2 = rowHas(rect.r2 + 1) ? rect.r2 + 1 : rect.r2;
      const c1 = colHas(rect.c1 - 1) ? rect.c1 - 1 : rect.c1;
      const c2 = colHas(rect.c2 + 1) ? rect.c2 + 1 : rect.c2;
      grew = r1 !== rect.r1 || r2 !== rect.r2 || c1 !== rect.c1 || c2 !== rect.c2;
      Object.assign(rect, { r1, r2, c1, c2 });
    }
    return rect;
  }

  /**
   * Sorts the data region by the active cell's column. A table's, a filter's
   * or a block's first row stays put as headings; in a range the user
   * selected, only when it looks like headings, as Excel judges it.
   */
  sort(descending: boolean, col = this.selection.col): Promise<EditResult | undefined> {
    const rect = this.dataRegion();
    if (!rect || rect.r2 <= rect.r1) {
      this.toast('info', msg('store.sort.noData'));
      return Promise.resolve(undefined);
    }
    const sel = this.selection.rects[0];
    const selected = !!sel && sel.r1 === rect.r1 && sel.r2 === rect.r2 && sel.c1 === rect.c1 && sel.c2 === rect.c2;
    const header = !selected || this.looksLikeHeadings(rect);
    return this.edit({ type: 'sort', sheet: this.active, rect, keys: [{ col, descending }], header }, 'undo', {
      label: msg(descending ? 'store.history.sortDescending' : 'store.history.sortAscending'),
    });
  }

  /**
   * Whether a range's first row looks like headings rather than data, as
   * Excel guesses it: text above a column of numbers, dates or logicals, or
   * bold above plain.
   */
  private looksLikeHeadings(rect: Rect): boolean {
    const view = this.view;
    if (!view || view.tableAt(rect.r1, rect.c1)?.rect.r1 === rect.r1) return true;
    const filled = (r: number, c: number) => {
      const rec = view.get(r, c);
      return rec && rec.t !== VT.None ? rec : null;
    };
    const bold = (r: number) => {
      for (let c = rect.c1; c <= rect.c2; c++) if (filled(r, c) && this.styles[view.styleAt(r, c)]?.bold) return true;
      return false;
    };
    if (bold(rect.r1) && !bold(rect.r1 + 1)) return true;
    for (let c = rect.c1; c <= rect.c2; c++) {
      if (filled(rect.r1, c)?.t !== VT.String) continue;
      let values = 0;
      let texts = 0;
      for (let r = rect.r1 + 1; r <= Math.min(rect.r2, rect.r1 + 50); r++) {
        const rec = filled(r, c);
        if (rec?.t === VT.String) texts++;
        else if (rec) values++;
      }
      if (values && !texts) return true;
    }
    return false;
  }

  /** The text a cell shows, as filters and find compare it, in the locale's format as Excel saves filters. */
  shownText(view: SheetView, row: number, col: number): string {
    const rec = view.get(row, col);
    if (!rec) return '';
    return displayText(rec, this.styles[rec.style] ?? this.styles[0], this.date1904, false, this.locale).text;
  }

  /** What the selected cells show, row by row, leaving out empty ones: the entries of a list imported from them. */
  selectedTexts(limit = 1000): string[] {
    const view = this.view;
    const rect = this.selection.rects[0];
    if (!view || !rect) return [];
    const cells: { row: number; col: number }[] = [];
    view.forEach(rect, (row, col) => cells.push({ row, col }));
    cells.sort((a, b) => a.row - b.row || a.col - b.col);
    const out: string[] = [];
    for (const c of cells) {
      const text = this.shownText(view, c.row, c.col).trim();
      if (text) out.push(text);
      if (out.length >= limit) break;
    }
    return out;
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

  /**
   * Shows only rows whose value in `col` is one of `allowed`; undefined
   * shows them all again. A filter in a table is the table's, as Excel
   * keeps it; the file keeps the filter, and the rows it hides, when saved.
   */
  async setFilter(rect: Rect, col: number, allowed: Set<string> | undefined): Promise<void> {
    const view = this.view;
    if (!view) return;
    const columns = new Map(view.filter?.columns ?? []);
    if (allowed) columns.set(col, allowed);
    else columns.delete(col);
    let target = view.filter?.rect ?? rect;
    let table = view.filter?.table;
    if (!view.filter) {
      const t = view.info.tables.find((x) => x.rect.r1 <= target.r2 && x.rect.r2 >= target.r1 && x.rect.c1 <= target.c2 && x.rect.c2 >= target.c1);
      if (t) {
        target = { ...t.rect, r2: t.rect.r2 - t.totalsRows };
        table = t.name;
      }
    }
    const rows = columns.size ? SheetView.rowsFilteredBy(target, columns, (row, c) => this.shownText(view, row, c)) : [];
    const filter: FilterState | null = columns.size
      ? { rect: target, columns: Object.fromEntries([...columns].map(([c, values]) => [c, [...values]])), ...(table ? { table } : {}) }
      : null;
    // How many rows the filter hides is said, as Excel's status bar shows how many it leaves.
    const hidden = msg(filter ? 'store.announce.filtered' : 'store.announce.filterCleared', { count: rows.length });
    await this.changeLayout({ kind: 'filter', filter, rows }, msg(filter ? 'store.history.filter' : 'store.history.clearFilter'), [], hidden);
  }

  async clearFilter(): Promise<void> {
    if (!this.view?.filter) return;
    await this.changeLayout({ kind: 'filter', filter: null, rows: [] }, msg('store.history.clearFilter'), [], msg('store.announce.filterCleared'));
  }

  // ---------------------------------------------------------------------------
  // Find and replace

  /**
   * Cells matching a search, in reading order: by sheet, then row, then
   * column. As in Excel, * stands for any run of characters and ? for any
   * one, and ~ before either means it literally.
   */
  find(q: FindQuery): { sheet: number; row: number; col: number }[] {
    if (!q.text) return [];
    const out: { sheet: number; row: number; col: number }[] = [];
    const needle = q.matchCase ? q.text : q.text.toLowerCase();
    const pattern = hasWildcards(needle) ? parsePattern(needle) : undefined;
    const matches = (text: string) =>
      pattern ? (q.wholeCell ? matchesWhole(pattern, text) : searchIn(pattern, text) !== -1) : q.wholeCell ? text === needle : text.includes(needle);
    const sheets = q.scope === 'workbook' ? this.sheets.map((_, i) => i) : [this.active];
    for (const s of sheets) {
      const view = this.sheets[s];
      if (view.info.kind !== 'worksheet') continue;
      const hits: { row: number; col: number }[] = [];
      view.forEach({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }, (row, col, rec) => {
        if (rec.t === VT.None && rec.f === undefined) return;
        let text = q.lookIn === 'formulas' ? editText(rec, this.styles[rec.style] ?? this.styles[0], this.date1904, this.locale) : this.shownText(view, row, col);
        if (!q.matchCase) text = text.toLowerCase();
        if (matches(text)) hits.push({ row, col });
      });
      hits.sort((a, b) => a.row - b.row || a.col - b.col);
      out.push(...hits.map((h) => ({ sheet: s, ...h })));
    }
    return out;
  }

  /** Replaces the search text in matching cells' contents (formulas included). Resolves to how many cells changed. */
  async replace(q: FindQuery, replacement: string, only?: { sheet: number; row: number; col: number }): Promise<number> {
    const hits = only ? [only] : this.find({ ...q, lookIn: 'formulas' });
    const pattern = parsePattern(q.matchCase ? q.text : q.text.toLowerCase());
    /** Every match in the text replaced, left to right, as the search finds them. */
    const replaceAll = (text: string) => {
      const hay = q.matchCase ? text : text.toLowerCase();
      let out = '';
      let from = 0;
      for (let span = matchSpan(pattern, hay, 0); span; span = span[1] < hay.length ? matchSpan(pattern, hay, span[1]) : undefined) {
        out += text.slice(from, span[0]) + replacement;
        from = span[1];
      }
      return out + text.slice(from);
    };
    const bySheet = new Map<number, { row: number; col: number; text: string }[]>();
    let skipped = 0;
    for (const h of hits) {
      const view = this.sheets[h.sheet];
      const rec = view.get(h.row, h.col);
      if (!rec || view.readOnlyReason(h.row, h.col)) {
        skipped++;
        continue;
      }
      const before = editText(rec, this.styles[rec.style] ?? this.styles[0], this.date1904, this.locale);
      const after = q.wholeCell ? replacement : replaceAll(before);
      if (after === before) continue;
      const list = bySheet.get(h.sheet) ?? [];
      list.push({ row: h.row, col: h.col, text: after });
      bySheet.set(h.sheet, list);
    }
    const ops: EditOp[] = [...bySheet].map(([sheet, cells]) => ({ type: 'inputs', sheet, cells }));
    const n = ops.reduce((k, o) => k + (o.type === 'inputs' ? o.cells.length : 0), 0);
    if (n) await this.edit(ops.length === 1 ? ops[0] : { type: 'batch', ops }, 'undo', { label: msg('store.history.replace') });
    if (skipped) this.toast('info', msg('store.replace.readOnly', { count: skipped }));
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

  /** Sets a column's width, in characters, and the other selected columns' if it is one of them. */
  async setColumnWidth(col: number, chars: number): Promise<void> {
    if (!this.view || !(chars >= 0)) return;
    await this.resizeLines('cols', col, Math.min(2000, Math.round(chars * 7)));
  }

  /** A row's height as shown here, in points as Excel gives it. */
  rowHeightPt(row: number): number {
    return Math.round((this.view?.rowHeight(row) ?? 20) * 75) / 100;
  }

  /** Sets a row's height, in points, and the other selected rows' if it is one of them. */
  async setRowHeight(row: number, points: number): Promise<void> {
    if (!this.view || !(points >= 0)) return;
    await this.resizeLines('rows', row, Math.min(546, Math.round((points * 4) / 3)));
  }

  /**
   * Keeps a size set in the grid, as dragging a column's edge does: the
   * grid shows it already. The other selected columns take it too.
   */
  resizedColumn(col: number, px: number): Promise<void> {
    return this.resizeLines('cols', col, Math.round(px));
  }

  /** Sets a row's height, and the other selected rows', dragged to in the grid. */
  resizedRow(row: number, px: number): Promise<void> {
    return this.resizeLines('rows', row, Math.min(546, Math.round(px)));
  }

  private async resizeLines(axis: 'rows' | 'cols', line: number, px: number): Promise<void> {
    const selected = this.selectedLines(axis);
    const lines = selected.includes(line) ? selected : [line];
    const sizes = Object.fromEntries(lines.map((l) => [l, px]));
    const op: LayoutOp = axis === 'cols' ? { kind: 'colWidths', widths: sizes } : { kind: 'rowHeights', heights: sizes };
    await this.changeLayout(op, msg(axis === 'cols' ? 'store.history.columnWidth' : 'store.history.rowHeight', { count: lines.length }));
  }

  /** Bumped when column widths change outside the grid. */
  colVersion = 0;

  /**
   * Goes to a reference typed in the name box, as Excel does: a cell or range,
   * on another sheet with "Prices!B4", or the range a defined name refers
   * to. The grid takes the focus, so typing goes into the cell. False when
   * the text is neither, or names something other than a range.
   */
  goTo(text: string): boolean {
    const to = this.referenceTo(text) ?? namedRange(this, text);
    if (!to) return false;
    if (to.sheet !== this.active) this.setActive(to.sheet);
    this.revealFocus = true;
    this.revealSeq++;
    this.setSelection({ row: to.rects[0].r1, col: to.rects[0].c1, rects: to.rects });
    return true;
  }

  /** The cells a reference such as "B4", "A1:C3" or "Prices!B4" names, if they are on a sheet here. */
  private referenceTo(text: string): NameTarget | undefined {
    const m = /^(?:(?:'((?:[^']|'')+)'|([^!']+))!)?([^!]+)$/.exec(text.trim());
    if (!m) return undefined;
    const name = m[1]?.replace(/''/g, "'") ?? m[2];
    const sheet = name === undefined ? this.active : this.sheets.findIndex((s) => s.info.name.toLowerCase() === name.toLowerCase());
    const rect = parseRange(m[3].trim());
    if (sheet < 0 || !rect || rect.r1 < 0 || rect.r2 >= MAX_ROWS || rect.c2 >= MAX_COLS) return undefined;
    return { sheet, rects: [rect] };
  }

  /** Adds, changes or deletes a defined name, as one step to undo. */
  changeName(change: NameChange, label: Text): Promise<EditResult | undefined> {
    return this.edit({ type: 'name', change }, 'undo', { label });
  }

  /** The selected cells as they are copied: the text each shows, and a number's exact value. */
  private selectionCells(rect: Rect): { text: string; value?: number }[][] {
    const view = this.view;
    if (!view) return [];
    const rows: { text: string; value?: number }[][] = [];
    // The file's used range, grown by any cells typed or pasted beyond it since it was opened.
    const extent = view.extent();
    const lastRow = Math.max(extent.maxRow, rect.r1);
    const lastCol = Math.max(extent.maxCol, rect.c1);
    for (let row = rect.r1; row <= Math.min(rect.r2, lastRow); row++) {
      if (view.gridRow(row) < 0) continue;
      const cells: { text: string; value?: number }[] = [];
      for (let col = rect.c1; col <= Math.min(rect.c2, lastCol); col++) {
        const rec = view.get(row, col);
        const style = rec ? (this.styles[rec.style] ?? this.styles[0]) : undefined;
        const raw = rec && style && this.numbers.copy === 'raw' ? exactText(rec, style, false, this.locale) : undefined;
        const text = rec && style ? (raw ?? displayText(rec, style, this.date1904, false, this.locale).text) : '';
        cells.push(rec?.t === VT.Number && rec.n !== undefined ? { text, value: rec.n } : { text });
      }
      rows.push(cells);
    }
    return rows;
  }

  /** The selected cells as tab-separated text, as they are copied to the clipboard. */
  selectionText(rect: Rect): string {
    return this.selectionCells(rect)
      .map((cells) => cells.map(({ text }) => (/[\t\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text)).join('\t'))
      .join('\n');
  }

  /**
   * The selected cells as the HTML table Excel puts on the clipboard: each
   * cell's text, escaped so nothing in a workbook becomes markup where it is
   * pasted, and each number's exact value (x:num), which Excel reads back.
   */
  selectionHtml(rect: Rect): string {
    const escape = (s: string) => s.replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);
    const rows = this.selectionCells(rect).map(
      (cells) => `<tr>${cells.map(({ text, value }) => `<td${value !== undefined ? ` x:num="${value}"` : ''}>${escape(text).replace(/\n/g, '<br>')}</td>`).join('')}</tr>`,
    );
    return `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><body><table>${rows.join('')}</table></body></html>`;
  }

  /** Copies or cuts the selection from the menu, where the keyboard's clipboard events aren't available. */
  async copySelection(cut: boolean): Promise<void> {
    const rect = this.selection.rects[0];
    if (!rect) return;
    this.noteCopy(rect, cut);
    try {
      const text = this.selectionText(rect);
      if (typeof ClipboardItem === 'undefined') await navigator.clipboard.writeText(text);
      else {
        const html = this.selectionHtml(rect);
        await navigator.clipboard.write([
          new ClipboardItem({ 'text/plain': new Blob([text], { type: 'text/plain' }), 'text/html': new Blob([html], { type: 'text/html' }) }),
        ]);
      }
    } catch {
      // Pasting within this workbook still works; other apps won't see it.
      this.toast('info', msg(cut ? 'store.clipboard.cutBlocked' : 'store.clipboard.copyBlocked'));
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
      this.toast('warn', msg('store.clipboard.readBlocked'));
      return;
    }
    await this.paste({ row: rect.r1, col: rect.c1 }, parseClipboardText(text), rect);
  }

  /**
   * Excel's Paste Special: only the values (formulas' results), the
   * formulas or the formats of what was copied in this workbook, or all of
   * it turned on its side. From another app only text is to hand, so values
   * and turning on its side are what can be pasted from there.
   */
  async pasteSpecial(what: PastePart, transpose = false): Promise<void> {
    const view = this.view;
    const clip = this.clip;
    const sel = this.selection;
    const at = { row: sel.rects[0]?.r1 ?? sel.row, col: sel.rects[0]?.c1 ?? sel.col };
    if (!view) return;
    if (clip?.cut) return this.toast('info', msg('store.pasteSpecial.notAfterCut'));
    if (!clip) {
      if (what === 'formulas' || what === 'formats') return this.toast('info', msg('store.pasteSpecial.copyHere'));
      let text: string;
      try {
        text = await navigator.clipboard.readText();
      } catch {
        return this.toast('warn', msg('store.clipboard.readBlocked'));
      }
      const values = parseClipboardText(text);
      const cell = { r1: at.row, c1: at.col, r2: at.row, c2: at.col };
      return this.paste(at, transpose ? transposeGrid(values) : values, cell);
    }
    const h = clip.rect.r2 - clip.rect.r1 + 1;
    const w = clip.rect.c2 - clip.rect.c1 + 1;
    const dst: Rect = { r1: at.row, c1: at.col, r2: at.row + (transpose ? w : h) - 1, c2: at.col + (transpose ? h : w) - 1 };
    if (dst.r2 >= MAX_ROWS || dst.c2 >= MAX_COLS) return this.toast('warn', msg('store.pasteSpecial.noRoom'));
    if (what !== 'formats') {
      const src = this.sheets[clip.sheet];
      const keeps = (row: number, col: number) => {
        if (what === 'values') return false;
        const [down, across] = transpose ? [col - dst.c1, row - dst.r1] : [row - dst.r1, col - dst.c1];
        return !!((src.get(clip.rect.r1 + down, clip.rect.c1 + across)?.flags ?? 0) & CellFlag.Formula);
      };
      if (!(await this.okToReplace(this.formulasLost(view, [dst], keeps)))) return;
    }
    const label = msg(transpose && what === 'all' ? 'store.history.pasteTransposed' : `store.history.paste.${what}`);
    const readOnly = this.schemaReadOnly(dst);
    await this.edit({ type: 'copy', srcSheet: clip.sheet, src: clip.rect, sheet: this.active, dst, what, transpose, ...(readOnly && { readOnly }) }, 'undo', { label });
  }

  /** The references in the formula being typed, by sheet, which the grid outlines in their colours. */
  formulaRefs: { sheet: number; rect: Rect; color: number }[] = [];

  /**
   * Outlines on the grid the references in a formula being typed in a cell
   * of the active sheet, as Excel does; anything else (or nothing) clears
   * them. References to sheets the workbook doesn't have are left out.
   */
  showFormulaRefs(text: string | undefined): void {
    const refs: WorkbookStore['formulaRefs'] = [];
    for (const r of text ? formulaReferences(text) : []) {
      const name = r.sheet?.toLowerCase();
      const sheet = name === undefined ? this.active : this.sheets.findIndex((s) => s.info.name.toLowerCase() === name);
      if (r.rect && sheet >= 0) refs.push({ sheet, rect: r.rect, color: r.color });
    }
    // Only a change of references redraws, not every key typed.
    if (JSON.stringify(refs) === JSON.stringify(this.formulaRefs)) return;
    this.formulaRefs = refs;
    this.notify();
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
      title: msg('store.replaceFormula.title', { count: n }),
      body: msg('store.replaceFormula.body', { count: n }),
      confirmLabel: msg('store.replaceFormula.confirm', { count: n }),
    });
  }

  /** Typing still being checked before its edit is queued. */
  private readonly typing = new Set<Promise<void>>();

  /** Once every edit asked for, typing still being checked included, has been made. */
  private async settled(): Promise<void> {
    await Promise.all(this.typing);
    await this.queue;
  }

  async input(row: number, col: number, text: string): Promise<void> {
    const typed = this.typed(row, col, text);
    this.typing.add(typed);
    try {
      await typed;
    } finally {
      this.typing.delete(typed);
    }
  }

  private async typed(row: number, col: number, text: string): Promise<void> {
    const view = this.view;
    if (!view) return;
    const reason = view.readOnlyReason(row, col);
    if (reason) return this.toast('warn', reason);
    const prepared = this.schema?.prepareInput(this.active, row, col, text) ?? { text };
    if ('reject' in prepared) return this.toast('warn', prepared.reject);
    if (!(await this.passesValidation(row, col, text))) return;
    const rec = view.get(row, col);
    if (rec?.flags && rec.flags & CellFlag.Formula && !text.startsWith('=') && !(await this.okToReplace(1))) return;
    if (text !== '' && !(await this.growTableTo({ r1: row, c1: col, r2: row, c2: col }))) return;
    if (this.schema) this.checking = { sheet: this.active, row, col };
    await this.edit({ type: 'input', sheet: this.active, row, col, text: prepared.text });
  }

  /**
   * As in Excel, typing or pasting in the row just below an Excel table
   * (one without a totals row) makes the table take that row in, with its
   * calculated columns and formats. A step of its own, undone separately.
   * False when it couldn't grow.
   */
  private async growTableTo(rect: Rect): Promise<boolean> {
    const view = this.view;
    const table = view?.info.tables.find((t) => !t.totalsRows && rect.r1 === t.rect.r2 + 1 && rect.c1 <= t.rect.c2 && rect.c2 >= t.rect.c1);
    if (!table) return true;
    return this.restructure({ type: 'growTable', sheet: this.active, table: table.name, rows: rect.r2 - rect.r1 + 1 });
  }

  async clear(rects: Rect[]): Promise<void> {
    const view = this.view;
    if (!view) return;
    if (!(await this.okToReplace(this.formulasLost(view, rects)))) return;
    // A schema's read-only cells are left as they are, as a protected sheet's locked ones are.
    const readOnly = rects.flatMap((r) => this.schemaReadOnly(r) ?? []);
    await this.edit({ type: 'clear', sheet: this.active, rects, ...(readOnly.length && { readOnly }) });
  }

  /** Records what was copied (or cut) so a paste back into the workbook can carry formulas, or move the cells. */
  noteCopy(rect: Rect, cut = false): void {
    const view = this.view;
    if (!view) return;
    const rec = view.get(rect.r1, rect.c1);
    const style = rec ? (this.styles[rec.style] ?? this.styles[0]) : undefined;
    const raw = rec && style && this.numbers.copy === 'raw' ? exactText(rec, style, false, this.locale) : undefined;
    const first = rec && style ? (raw ?? displayText(rec, style, this.date1904, false, this.locale).text) : '';
    this.clip = { sheet: this.active, rect, first, cut };
    this.notify();
  }

  /**
   * Pastes clipboard text, or what was cut or copied in this workbook when
   * the text is the same (so formulas come along, and a cut moves cells).
   * With `ownClipboard`, pastes what was cut or copied here without the text,
   * for when the browser won't let the page read the clipboard.
   */
  /**
   * Pastes cells. `exact` has, for pastes from Excel, the full value of each
   * number the text shows rounded (from the clipboard's HTML).
   */
  async paste(
    target: { row: number; col: number },
    values: readonly (readonly string[])[],
    selection: Rect,
    ownClipboard = false,
    exact?: readonly (readonly (number | null)[])[],
    formats?: readonly (readonly (FormatChange | null)[])[],
  ): Promise<void> {
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
      if (!(await this.growTableTo(dst))) return;
      const readOnly = this.schemaReadOnly(dst);
      await this.edit({ type: 'copy', srcSheet: clip.sheet, src: clip.rect, sheet: this.active, dst, ...(readOnly && { readOnly }) });
      return;
    }

    let grid = values.map((line) => [...line]);
    if (this.numbers.paste === 'strict') {
      // Ask before storing values that wouldn't read back the way they were pasted.
      const risks = pasteRisks(values, this.locale, this.t.language);
      if (risks.length) {
        const answer = await this.ask({
          title: msg('store.paste.riskTitle', { count: risks.length }),
          body: msg('store.paste.riskBody'),
          confirmLabel: msg('store.paste.asShown'),
          alternateLabel: msg('store.paste.keepText'),
          danger: false,
          details: {
            header: [msg('store.paste.riskPasted'), msg('store.paste.riskBecomes'), msg('store.paste.riskBecause')],
            rows: risks.slice(0, 20).map((r) => [r.text, r.becomes, this.t.text(r.reason)]),
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
    let numbers = exact?.length === values.length ? exact.map((line) => [...line]) : undefined;
    if (tile) {
      const block = grid;
      const blockNumbers = numbers;
      grid = [];
      numbers = blockNumbers && [];
      for (let r = 0; r < selH; r++) {
        grid.push(Array.from({ length: selW }, (_, c) => block[r % h]?.[c % w] ?? ''));
        numbers?.push(Array.from({ length: selW }, (_, c) => blockNumbers![r % h]?.[c % w] ?? null));
      }
    }
    if (this.schema) {
      for (let i = 0; i < grid.length; i++) {
        for (let j = 0; j < grid[i].length; j++) {
          const prepared = this.schema.prepareInput(this.active, dst.r1 + i, dst.c1 + j, grid[i][j]);
          if ('reject' in prepared) return this.toast('warn', msg('store.paste.rejected', { reason: prepared.reject }));
          grid[i][j] = prepared.text;
        }
      }
    }
    const keeps = (row: number, col: number) => (grid[row - dst.r1]?.[col - dst.c1] ?? '').startsWith('=');
    if (!(await this.okToReplace(this.formulasLost(view, [dst], keeps)))) return;
    if (rectCells(dst) > 200_000) this.toast('info', msg('store.paste.large', { count: rectCells(dst) }));
    if (!(await this.growTableTo(dst))) return;
    const readOnly = this.schemaReadOnly(dst);
    const text: EditOp = { type: 'pasteText', sheet: this.active, row: dst.r1, col: dst.c1, values: grid, ...(numbers ? { exact: numbers } : {}), ...(readOnly && { readOnly }) };
    const formatOps = formats && this.features.format && rectCells(dst) <= 100_000 ? this.pastedFormats(dst, formats, h, w) : [];
    await this.edit(formatOps.length ? { type: 'batch', ops: [text, ...formatOps] } : text, 'undo', { label: msg('store.history.paste') });
  }

  /**
   * The formats pasted from another app, as format changes over the cells
   * they land on, tiled as the values are: cells in a row with the same
   * formats share one change.
   */
  private pastedFormats(dst: Rect, formats: readonly (readonly (FormatChange | null)[])[], h: number, w: number): EditOp[] {
    const byChange = new Map<string, { change: FormatChange; rects: Rect[] }>();
    for (let row = dst.r1; row <= dst.r2; row++) {
      let run: { key: string; rect: Rect } | undefined;
      for (let col = dst.c1; col <= dst.c2 + 1; col++) {
        const change = col <= dst.c2 ? formats[(row - dst.r1) % h]?.[(col - dst.c1) % w] : null;
        const key = change ? JSON.stringify(change) : '';
        if (run && run.key === key) {
          run.rect.c2 = col;
          continue;
        }
        if (run) {
          const group = byChange.get(run.key) ?? { change: JSON.parse(run.key) as FormatChange, rects: [] };
          group.rects.push(run.rect);
          byChange.set(run.key, group);
        }
        run = key ? { key, rect: { r1: row, r2: row, c1: col, c2: col } } : undefined;
      }
    }
    return [...byChange.values()].map(({ change, rects }) => ({ type: 'format', sheet: this.active, rects, change }));
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
    const readOnly = this.schemaReadOnly(dst);
    await this.edit({ type: 'copy', srcSheet: this.active, src, sheet: this.active, dst, ...(readOnly && { readOnly }) }, 'undo', {
      label: msg(across ? 'store.history.fillRight' : 'store.history.fillDown'),
    });
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
    const readOnly = this.schemaReadOnly(dst);
    await this.edit({ type: 'fill', sheet: this.active, src, dst, lists: this.fillLists, ...(readOnly && { readOnly }) });
  }

  /** What Excel's Fill Series dialog starts with for the selection: rows for a wide one, and dates when it starts with one. */
  seriesStart(): SeriesOptions {
    const sel = this.selection;
    const rect = sel.rects[0] ?? { r1: sel.row, r2: sel.row, c1: sel.col, c2: sel.col };
    const rec = this.view?.get(rect.r1, rect.c1);
    const date = rec?.t === VT.Number && isDateStyle(this.styles[rec.style] ?? this.styles[0]);
    return { by: rect.c2 - rect.c1 > rect.r2 - rect.r1 ? 'rows' : 'columns', type: date ? 'date' : 'linear', dateUnit: 'day', step: 1, trend: false };
  }

  /**
   * Excel's Fill Series over the selection (its first range): each row or
   * column becomes a series from the cells it starts with. With one cell
   * selected, the series runs on from it, down or across, until the stop
   * value, as in Excel. One step to undo; locked cells of a protected sheet
   * and the schema's read-only cells are left as they are, as filling by
   * dragging leaves them.
   */
  async fillSeries(series: SeriesOptions): Promise<void> {
    const view = this.view;
    if (!view) return;
    const sel = this.selection;
    let rect = sel.rects[0] ?? { r1: sel.row, r2: sel.row, c1: sel.col, c2: sel.col };
    const across = series.by === 'rows';
    if (rect.r1 === rect.r2 && rect.c1 === rect.c2 && series.stop !== undefined && series.type !== 'autofill') {
      const rec = view.get(rect.r1, rect.c1);
      const start = rec?.t === VT.Number ? rec.n : undefined;
      const run = start === undefined ? undefined : seriesLine([start], across ? MAX_COLS - rect.c1 : MAX_ROWS - rect.r1, series, this.date1904);
      // A series that doesn't move would never reach the stop value.
      if (run?.direction) rect = across ? { ...rect, c2: rect.c1 + run.values.length } : { ...rect, r2: rect.r1 + run.values.length };
    }
    // The cells a series starts from stay; a trend replaces its numbers, and AutoFill keeps what it continues.
    const length = across ? rect.c2 - rect.c1 + 1 : rect.r2 - rect.r1 + 1;
    const runs = new Map<number, number>();
    /** How many cells the row or column starts with before an empty one. */
    const run = (line: number) => {
      let n = runs.get(line);
      if (n === undefined) {
        n = 0;
        while (n < length) {
          const rec = across ? view.get(line, rect.c1 + n) : view.get(rect.r1 + n, line);
          if (!rec || rec.t === VT.None) break;
          n++;
        }
        runs.set(line, n);
      }
      return n;
    };
    const keeps = (row: number, col: number) => {
      const at = across ? col - rect.c1 : row - rect.r1;
      if (series.type === 'autofill') return at < run(across ? row : col);
      return at === 0 && !(series.trend && series.type !== 'date');
    };
    if (!(await this.okToReplace(this.formulasLost(view, [rect], keeps)))) return;
    const readOnly = this.schemaReadOnly(rect);
    await this.edit({ type: 'series', sheet: this.active, rect, series, lists: this.fillLists, ...(readOnly && { readOnly }) }, 'undo', {
      label: msg('fill.history.series'),
    });
  }

  /**
   * The cells in `rect` the schema keeps read-only, in runs, for an edit to
   * leave as they are; the worker knows the sheet's own locked cells.
   */
  private schemaReadOnly(rect: Rect): ReadOnlyCells[] | undefined {
    const view = this.view;
    if (!this.schema || !view?.constraint) return undefined;
    type Run = ReadOnlyCells & { key: string };
    const out: Run[] = [];
    let above: Run[] = [];
    for (let row = rect.r1; row <= rect.r2; row++) {
      const runs: Run[] = [];
      for (let col = rect.c1; col <= rect.c2; col++) {
        const reason = view.constraint(row, col);
        const key = reason === undefined ? '' : JSON.stringify(reason);
        const last = runs[runs.length - 1];
        if (key && last?.key === key && last.c2 === col - 1) last.c2 = col;
        else if (key) runs.push({ r1: row, r2: row, c1: col, c2: col, reason: reason!, key });
      }
      // A row with the same runs as the one above makes those taller.
      if (runs.length && runs.length === above.length && runs.every((r, i) => r.c1 === above[i].c1 && r.c2 === above[i].c2 && r.key === above[i].key)) {
        for (const r of above) r.r2 = row;
        continue;
      }
      out.push(...runs);
      above = runs;
    }
    return out.length ? out.map(({ key: _key, ...r }) => r) : undefined;
  }

  // ---------------------------------------------------------------------------
  // Lists that filling continues

  /** Day and month names in the editor's language and locale, and in English. */
  get builtInFillLists(): string[][] {
    return builtInLists([this.t.language, this.locale]);
  }

  /** The host's lists, from the `customLists` option. */
  get hostFillLists(): string[][] {
    return cleanLists(this.options.customLists);
  }

  private ownLists?: { user: string | undefined; lists: string[][] };

  /** The user's own lists, kept in this browser for each user. */
  get userLists(): string[][] {
    const user = this.options.user;
    if (!this.ownLists || this.ownLists.user !== user) this.ownLists = { user, lists: loadUserLists(user) };
    return this.ownLists.lists;
  }

  /** Every list filling continues, the first to match winning: the built-in ones, the host's, then the user's. */
  get fillLists(): string[][] {
    return [...this.builtInFillLists, ...this.hostFillLists, ...this.userLists];
  }

  /** Replaces the user's own lists, keeping them in the browser where it can. */
  setUserLists(lists: string[][]): void {
    const user = this.options.user;
    const clean = cleanLists(lists);
    this.ownLists = { user, lists: clean };
    this.listsNotKept = !saveUserLists(user, clean);
    this.notify();
  }

  /** Deletes one of the user's lists, once they confirm it. */
  async deleteUserList(index: number): Promise<boolean> {
    const list = this.userLists[index];
    if (!list) return false;
    const sure = await this.confirm({
      title: msg('fill.lists.deleteTitle'),
      body: msg('fill.lists.deleteBody', { list: list.join(', ') }),
      confirmLabel: msg('fill.lists.deleteConfirm'),
    });
    if (sure) this.setUserLists(this.userLists.filter((_, i) => i !== index));
    return sure;
  }

  /** Changes the host's lists while the editor is open. */
  setCustomLists(lists: string[][] | undefined): void {
    this.options.customLists = lists;
    this.notify();
  }

  // ---------------------------------------------------------------------------
  // Notes and comment threads

  /**
   * Who new notes and comments are by: the host's user. Without one, notes
   * have no author, and threads (which need someone) are by "Anonymous".
   */
  get commentAuthor(): string {
    return this.options.user || this.t('store.comment.anonymous');
  }

  /**
   * Opens a cell's note or comments, on the active sheet, to read or change.
   * Where the cell has neither and can't have them, says why instead.
   */
  openComment(row: number, col: number, at: { x: number; y: number }, start: 'note' | 'thread' = 'note'): void {
    const view = this.view;
    if (!view) return;
    const reason = view.commentReadOnlyReason(row, col);
    if (reason && !view.noteAt(row, col)) return this.toast('warn', reason);
    this.set('commentFor', { sheet: this.active, row, col, x: at.x, y: at.y, start });
  }

  /** Adds a note to a cell on the active sheet, or changes its text. */
  setNote(row: number, col: number, text: string): Promise<boolean> {
    const had = this.view?.noteAt(row, col);
    return this.changeComment(row, col, { action: 'setNote', text, author: this.options.user || undefined }, had ? 'store.history.editNote' : 'store.history.addNote');
  }

  /** Deletes a cell's note, or its whole comment thread. */
  deleteComment(row: number, col: number): Promise<boolean> {
    const thread = this.view?.noteAt(row, col)?.kind === 'thread';
    return this.changeComment(row, col, { action: 'delete' }, thread ? 'store.history.deleteThread' : 'store.history.deleteNote');
  }

  startThread(row: number, col: number, text: string): Promise<boolean> {
    return this.changeComment(row, col, { action: 'startThread', text, author: this.commentAuthor }, 'store.history.startThread');
  }

  reply(row: number, col: number, text: string): Promise<boolean> {
    return this.changeComment(row, col, { action: 'reply', text, author: this.commentAuthor }, 'store.history.reply');
  }

  /** Changes the text of one comment in a thread: one of the user's own. */
  editComment(row: number, col: number, id: string, text: string): Promise<boolean> {
    return this.changeComment(row, col, { action: 'editEntry', id, text }, 'store.history.editComment');
  }

  /** Deletes one reply; the thread's first comment takes the thread with it. */
  deleteReply(row: number, col: number, id: string): Promise<boolean> {
    const first = this.view?.noteAt(row, col)?.entries[0]?.id === id;
    return this.changeComment(row, col, { action: 'deleteEntry', id }, first ? 'store.history.deleteThread' : 'store.history.deleteReply');
  }

  resolveThread(row: number, col: number, resolved: boolean): Promise<boolean> {
    return this.changeComment(row, col, { action: 'resolve', resolved }, resolved ? 'store.history.resolveThread' : 'store.history.reopenThread');
  }

  /** Makes a change to a note or thread on the active sheet, as a step of its own; resolves whether it was made. */
  private async changeComment(row: number, col: number, change: CommentChange, label: MessageKey): Promise<boolean> {
    const view = this.view;
    if (!view || this.status !== 'ready') return false;
    const reason = view.commentReadOnlyReason(row, col);
    if (reason) {
      this.toast('warn', reason);
      return false;
    }
    const r = await this.edit({ type: 'comment', sheet: this.active, row, col, change }, 'undo', { label: msg(label) });
    return !!r?.noteRedo?.length;
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
      const options = listOptions(dv, this.sheets, this.styles, this.date1904, this.locale);
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
          entry.options = src ? listOptions({ ...dv, source: src }, this.sheets, this.styles, this.date1904, this.locale) : undefined;
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
  private async alert(dv: ValidationAlert, fallbackTitle: Text): Promise<boolean> {
    // The workbook's own title and message are shown as written.
    const message: Text = dv.error ?? DEFAULT_LIST_ERROR;
    if (dv.errorStyle === 'stop') {
      this.toast('error', dv.errorTitle ? msg('store.validation.titled', { title: dv.errorTitle, message }) : message);
      return false;
    }
    if (dv.errorStyle === 'warning') {
      return this.confirm({
        title: dv.errorTitle ?? fallbackTitle,
        body: msg('store.validation.keepAnyway', { message }),
        confirmLabel: msg('store.validation.keepValue'),
      });
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
      if (options && !listAllows(options, text.replace(/^'/, ''))) return this.alert(dv, msg('store.validation.notInList'));
    }
    const rule = this.ruleAt(this.active, row, col);
    if (rule?.rule.showError) {
      const res = await this.request({ type: 'validate', sheet: this.active, row, col, text, rule: rule.index } as Omit<WorkerRequest, 'id'>).catch(() => undefined);
      if (res?.type === 'validated' && !res.ok) return this.alert(rule.rule, msg('store.validation.notAllowed'));
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
    let closest: { name: Localized; problems: Text[]; found: number } | undefined;
    // A file stamped with a schema when it was saved tries that schema first.
    for (const s of stampedFirst(this.schemas, this.properties)) {
      let r: BindResult;
      try {
        r = bindSchema(s, this.sheets);
      } catch (e) {
        // A schema that can't be read against this workbook is passed over, saying why.
        r = { problems: [errorText(e)], found: 0 };
      }
      if (r.binding) {
        this.schemaNote = versionNote(s, this.properties, this.t.language);
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
      language: () => this.t.language,
      views: () => this.sheets,
      styles: () => this.styles,
      date1904: () => this.date1904,
      followUp: async (op) => {
        await this.edit(op, 'merge');
      },
      changed: () => {
        this.schemaVersion++;
        this.sayProblemsWithTyped();
        this.notify();
      },
    });
    this.schema = ctl;
    this.sheets.forEach((v, i) => {
      v.constraint = (row, col) => ctl.readOnlyReason(i, row, col);
    });
    ctl.validate();
    // On a phone the panel would cover the sheet; the schema chip opens it.
    if (ctl.summary().errors && !this.narrow) this.panel = 'issues';
    this.notify();
  }

  /** The stamp saving writes into the file: the applied schema's id and version, if it has a version. */
  private schemaStamp(): Record<string, string> | undefined {
    return stampFor(this.schema?.schema);
  }

  /** Adds a schema, replacing any with the same id, and re-checks the open workbook. */
  registerSchema(schema: WorkbookSchema): void {
    this.hostSchemas.set(schema.id, schema);
    // Newest first, so a schema just applied takes priority over older ones that also fit.
    this.schemas = [schema, ...this.schemas.filter((s) => s.id !== schema.id)];
    this.bindSchemas();
  }

  /**
   * Applies a schema made or imported in the schema designer. A copy of a
   * schema the host provided keeps the host's code hooks (its async checks,
   * business logic and last check before saving), which JSON can't carry.
   */
  applyDesignedSchema(schema: WorkbookSchema): void {
    const host = this.hostSchemas.get(schema.id);
    const applied = host ? withHooksFrom(schema, host) : schema;
    this.schemas = [applied, ...this.schemas.filter((s) => s.id !== schema.id)];
    this.bindSchemas();
  }

  /** Whether the host provided a schema with this id, whose code hooks a designed copy keeps. */
  isHostSchema(id: string): boolean {
    return this.hostSchemas.has(id);
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
    if (!isEditOp(op)) return this.toast('warn', op);
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
    if (!isEditOp(op)) return this.toast('warn', op);
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

  toast(kind: Toast['kind'], text: Text): void {
    const id = ++this.toastSeq;
    this.toasts = [...this.toasts.slice(-3), { id, kind, text }];
    this.notify();
    // Errors stay until dismissed, so they can be read at any pace; warnings stay longer than news.
    if (kind === 'error') this.emit({ type: 'error', message: english.text(text) });
    else {
      const expire = () => (this.toastsHeld ? setTimeout(expire, 1000) : this.dismissToast(id));
      setTimeout(expire, kind === 'warn' ? 10_000 : 5000);
    }
  }

  /** Keeps messages on screen while the pointer or focus is on one, as WCAG asks of content that times out. */
  holdToasts(held: boolean): void {
    this.toastsHeld = held;
  }

  /** Says something to screen readers without showing it, such as what an action just did. */
  announce(text: Text): void {
    this.spoken = [...this.spoken.slice(-2), { id: ++this.toastSeq, text }];
    this.notify();
  }

  /** Says at once what is wrong with the value just typed, once its checks have run. */
  private sayProblemsWithTyped(): void {
    const at = this.checking;
    const ctl = this.schema;
    if (!at || !ctl || ctl.isPending(at.sheet, at.row, at.col)) return;
    this.checking = undefined;
    const issues = ctl.cellIssues(at.sheet, at.row, at.col);
    if (!issues.length) return;
    const problems = issues
      .map((i): Text => msg(i.severity === 'error' ? 'grid.spokenError' : 'grid.spokenWarning', { message: i.message }))
      .reduce((said, next) => msg('store.announce.andThen', { said, next }));
    // Shown by the cell's mark and in the formula bar; said here.
    this.announce(msg('store.announce.problems', { ref: toA1(at.row, at.col), problems }));
  }

  private emit(event: EditorEvent): void {
    try {
      this.options.onEvent?.(event);
    } catch {
      // The host's logging mustn't break the editor.
    }
  }

  dismissToast(id: number): void {
    this.toasts = this.toasts.filter((t) => t.id !== id);
    this.notify();
  }

  set<
    K extends
      | 'panel'
      | 'bannerDismissed'
      | 'schemaBannerDismissed'
      | 'highlightEdits'
      | 'exactNumbers'
      | 'designerOpen'
      | 'menu'
      | 'filterPicker'
      | 'findOpen'
      | 'shortcutsOpen'
      | 'columnWidthFor'
      | 'rowHeightFor'
      | 'commentFor'
      | 'sheetMenu'
      | 'renameSheetFor'
      | 'unhideOpen'
      | 'nameDialog'
      | 'nameManagerOpen'
      | 'linkDialog'
      | 'seriesDialog'
      | 'customListsOpen'
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
