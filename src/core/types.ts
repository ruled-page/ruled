/** Inclusive, 0-based cell rectangle. */
export interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/**
 * What a cell holds, as it is stored in the file. Formula text is in file
 * form: no leading "=", and newer functions keep their `_xlfn.` prefix.
 */
export type CellContent =
  | { kind: 'empty' }
  | { kind: 'number'; value: number }
  | { kind: 'string'; value: string }
  | { kind: 'bool'; value: boolean }
  | { kind: 'error'; value: string }
  | { kind: 'formula'; formula: string };

export const EMPTY: CellContent = { kind: 'empty' };

/** Value type tags used in the compact grid arrays. */
export const VT = { None: 0, Number: 1, String: 2, Bool: 3, Error: 4 } as const;

export const CellFlag = {
  /** The cell holds a formula. */
  Formula: 1,
  /** The in-app engine cannot calculate this formula; Excel's saved value is shown. */
  Unsupported: 2,
  /** Part of a multi-cell array or spill range. */
  ArrayMember: 4,
  /** Inputs changed since Excel last calculated this; Excel recalculates on open. */
  Stale: 8,
  /** The engine's result differed from Excel's saved value when the file was opened. */
  Mismatch: 16,
  /** The file had no saved value for this formula; the engine calculated it. */
  NoCache: 32,
  /** Changed by the user in this session. */
  Edited: 64,
} as const;

/**
 * A sheet's cells in compressed sparse row form: the cells of row r are at
 * indices rowPtr[r] .. rowPtr[r + 1] - 1, ordered by column.
 */
export interface SheetGrid {
  rowCount: number;
  rowPtr: Int32Array;
  cols: Int32Array;
  types: Uint8Array;
  nums: Float64Array;
  strs: (string | undefined)[];
  styles: Int32Array;
  flags: Uint8Array;
  /** Cell index → formula in file form. Shared formulas are expanded. */
  formulas: Map<number, string>;
}

export interface DisplayStyle {
  numFmt: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  fontSizePx: number;
  fontFamily: string;
  color?: string;
  fill?: string;
  hAlign?: string;
  vAlign?: string;
  wrap: boolean;
  indent: number;
}

export interface Anchor {
  row: number;
  col: number;
  /** Offset inside the cell, in pixels. */
  rowOffPx: number;
  colOffPx: number;
}

export interface SheetObject {
  kind: 'chart' | 'image' | 'shape' | 'pivot' | 'control';
  name: string;
  title?: string;
  chartType?: string;
  from?: Anchor;
  to?: Anchor;
  /** Size in pixels when the object is anchored to one cell only. */
  sizePx?: { w: number; h: number };
}

export interface BlockedRegion extends Rect {
  kind: 'pivot' | 'array' | 'tableHeader';
  reason: string;
}

export interface MergeRange extends Rect {}

/** Where an Excel list validation gets its allowed values. */
export type ListSource =
  | { kind: 'values'; values: string[] }
  /** A range, possibly on another sheet; read live so edits to it show up. */
  | { kind: 'range'; sheet: number; rect: Rect }
  /** A formula (OFFSET, INDIRECT, …) the engine evaluates when the list is needed. */
  | { kind: 'formula'; formula: string };

/** Alert behaviour shared by every Excel data validation rule. */
export interface ValidationAlert {
  allowBlank: boolean;
  /** Whether Excel shows an alert for values that break the rule. */
  showError: boolean;
  errorStyle: 'stop' | 'warning' | 'information';
  errorTitle?: string;
  error?: string;
  promptTitle?: string;
  prompt?: string;
}

/** A non-list Excel data validation rule: number, date, time, text length or custom formula. */
export interface RuleValidation extends ValidationAlert {
  rects: Rect[];
  kind: 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'custom';
  operator: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';
  /** Bounds or the custom formula, relative to the rule's top-left cell. */
  formula1: string;
  formula2?: string;
}

/** An Excel "list" data validation rule on a sheet. */
export interface ListValidation extends ValidationAlert {
  rects: Rect[];
  source: ListSource;
  /** False when Excel's "In-cell dropdown" option is off. */
  inCellDropdown: boolean;
}

export interface SheetInfo {
  index: number;
  name: string;
  kind: 'worksheet' | 'chartsheet' | 'other';
  state: 'visible' | 'hidden' | 'veryHidden';
  /** Used range, inclusive; -1 when the sheet is empty. */
  maxRow: number;
  maxCol: number;
  defaultColWidthPx: number;
  defaultRowHeightPx: number;
  /** Column index → width in pixels, for columns with a custom width. */
  colWidths: Record<number, number>;
  /** Row index → height in pixels, for rows with a custom height. */
  rowHeights: Record<number, number>;
  hiddenCols: number[];
  hiddenRows: number[];
  frozenCols: number;
  frozenRows: number;
  merges: MergeRange[];
  listValidations: ListValidation[];
  ruleValidations: RuleValidation[];
  blocked: BlockedRegion[];
  objects: SheetObject[];
  protected: boolean;
  showGridLines: boolean;
  tabColor?: string;
}

export type FindingKind =
  | 'chart'
  | 'chartsheet'
  | 'pivot'
  | 'table'
  | 'image'
  | 'shape'
  | 'control'
  | 'macro'
  | 'externalLink'
  | 'connection'
  | 'slicer'
  | 'conditionalFormat'
  | 'dataValidation'
  | 'comment'
  | 'sparkline'
  | 'hyperlink'
  | 'protection'
  | 'hiddenSheet'
  | 'arrayFormula'
  | 'mergedCells'
  | 'frozenRows'
  | 'autoFilter'
  | 'richText'
  | 'cellImage'
  | 'formulaEngine';

export interface Finding {
  kind: FindingKind;
  /** "warn" items are surfaced in the banner when the file opens. */
  severity: 'info' | 'warn';
  title: string;
  detail: string;
  sheet?: string;
  where?: string;
}

export interface LoadResult {
  fileName: string;
  sheets: SheetInfo[];
  grids: (SheetGrid | null)[];
  styles: DisplayStyle[];
  findings: Finding[];
  date1904: boolean;
  timings: Record<string, number>;
}

export interface CalcSample {
  sheet: string;
  ref: string;
  formula: string;
  reason: string;
  excel?: string;
  engine?: string;
}

export interface CalcReport {
  buildMs: number;
  formulaCells: number;
  /** Formulas the engine calculates and keeps up to date. */
  liveCells: number;
  unsupported: number;
  mismatched: number;
  noCachedValue: number;
  unsupportedFunctions: string[];
  samples: CalcSample[];
  /** Value corrections for formulas the file had no saved value for. */
  updates: CellUpdate[];
  findings: Finding[];
  /** Why each formula is not calculated here, keyed by "sheet:row:col". */
  cellReasons: Map<string, string>;
}

export interface CellUpdate {
  sheet: number;
  row: number;
  col: number;
  t: number;
  n?: number;
  s?: string;
  style: number;
  /** Formula in file form, when the cell holds one. */
  f?: string;
  flags: number;
}

/** A cell that differs from the opened file. Values are as people read them. */
export interface CellChange {
  sheet: string;
  ref: string;
  before: string;
  after: string;
  /** Only the cell's format changed. */
  formatOnly?: boolean;
}

export interface CellSnapshot {
  sheet: number;
  row: number;
  col: number;
  content: CellContent;
  style: number;
}

export type EditOp =
  /** Text typed into a cell, parsed the way Excel parses input. */
  | { type: 'input'; sheet: number; row: number; col: number; text: string }
  /** Text pasted from outside the app; each value is parsed like typed input. */
  | { type: 'pasteText'; sheet: number; row: number; col: number; values: string[][] }
  | { type: 'clear'; sheet: number; rects: Rect[] }
  /** Copy cells within the workbook, tiling the source over the destination. */
  | { type: 'copy'; srcSheet: number; src: Rect; sheet: number; dst: Rect }
  /** Drag-fill from the source into the destination, extending series. */
  | { type: 'fill'; sheet: number; src: Rect; dst: Rect }
  /** Put cells back to an exact earlier state (undo/redo). */
  | { type: 'restore'; cells: CellSnapshot[] }
  /** Text for scattered cells, parsed like typed input (used by schema hooks). */
  | { type: 'inputs'; sheet: number; cells: { row: number; col: number; text: string }[] }
  /** Several operations applied in order and undone as one. */
  | { type: 'batch'; ops: EditOp[] };

export interface EditResult {
  updates: CellUpdate[];
  undo: CellSnapshot[];
  redo: CellSnapshot[];
  skipped: number;
  skippedReasons: string[];
  warnings: string[];
  newStyles: { index: number; style: DisplayStyle }[];
  editedCount: number;
  /** Set when the input could not be accepted at all, e.g. a malformed formula. */
  rejected?: string;
}
