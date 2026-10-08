import type { Text } from '../i18n/message';
import type { ChartSpec } from './charts';
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

/** One side of a cell border, as drawn: width in pixels, line pattern and colour. */
export interface BorderEdge {
  width: 1 | 2 | 3;
  pattern: 'solid' | 'dashed' | 'dotted' | 'double';
  /** Unset for Excel's automatic colour, which follows the theme's grid text. */
  color?: string;
}

export interface CellBorders {
  left?: BorderEdge;
  right?: BorderEdge;
  top?: BorderEdge;
  bottom?: BorderEdge;
}

/**
 * A differential format: only the parts of a style that a table style or a
 * conditional format changes. Inner `horizontal` and `vertical` edges apply
 * between the cells of the region it covers.
 */
export interface DiffStyle {
  fill?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  border?: CellBorders & { horizontal?: BorderEdge; vertical?: BorderEdge };
}

export type TableElement =
  | 'wholeTable'
  | 'headerRow'
  | 'totalRow'
  | 'firstColumn'
  | 'lastColumn'
  | 'firstRowStripe'
  | 'secondRowStripe'
  | 'firstColumnStripe'
  | 'secondColumnStripe'
  | 'firstHeaderCell'
  | 'lastHeaderCell'
  | 'firstTotalCell'
  | 'lastTotalCell';

/** A table style resolved to formats per element, with its stripe sizes. */
export interface TableLook {
  elements: Partial<Record<TableElement, DiffStyle>>;
  rowStripe: number;
  colStripe: number;
}

/** An Excel table (ListObject). */
export interface TableInfo {
  name: string;
  rect: Rect;
  headerRows: number;
  totalsRows: number;
  columns: string[];
  /** Each column's calculated formula, which Excel fills into every row of the column, in file form. */
  calculated?: (string | undefined)[];
  /** The table's part in the package. */
  path?: string;
  style?: string;
  showFirstColumn: boolean;
  showLastColumn: boolean;
  showRowStripes: boolean;
  showColumnStripes: boolean;
  /** How the table is drawn; missing when its style is unknown or "none". */
  look?: TableLook;
}

/** A threshold in a colour scale, data bar or icon set. */
export interface Cfvo {
  type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula' | 'autoMin' | 'autoMax';
  val?: string;
  /** For icon sets: the value must be >= the threshold (else >). */
  gte: boolean;
}

/** One conditional formatting rule, as Excel stores it. */
export interface CfRule {
  rects: Rect[];
  type: string;
  /** Lower numbers win. */
  priority: number;
  stopIfTrue: boolean;
  style?: DiffStyle;
  operator?: string;
  /** In file form, written for the top-left cell of the first range. */
  formulas: string[];
  text?: string;
  rank?: number;
  percent?: boolean;
  bottom?: boolean;
  aboveAverage?: boolean;
  equalAverage?: boolean;
  stdDev?: number;
  colorScale?: { cfvo: Cfvo[]; colors: string[] };
  dataBar?: { cfvo: Cfvo[]; color: string; negativeColor?: string; showValue: boolean; gradient: boolean; minLength: number; maxLength: number };
  iconSet?: { set: string; cfvo: Cfvo[]; reverse: boolean; showValue: boolean };
}

/** What conditional formatting does to one cell. */
export interface CfCell {
  row: number;
  col: number;
  /** Index into CfResult.styles. */
  style?: number;
  /** Data bar extent as fractions of the cell width. */
  bar?: { from: number; to: number; color: string; gradient: boolean };
  icon?: { set: string; index: number };
  hideValue?: boolean;
}

export interface CfResult {
  sheet: number;
  styles: DiffStyle[];
  cells: CfCell[];
  /** Rules too large to evaluate here. */
  skipped: number;
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
  border?: CellBorders;
  /** Left unlocked for editing when its sheet is protected; cells are locked unless the format says otherwise. */
  unlocked?: boolean;
  /** Its formula is hidden from the formula bar when its sheet is protected. */
  formulaHidden?: boolean;
  /** Its font is in the theme's hyperlink colour, as Excel's Hyperlink cell style has it. */
  linkFont?: boolean;
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
  /** The kind of chart ("Bar", "Pie"), to show as "{{type}} chart"; Excel's own name for kinds without a message. */
  chartType?: Text;
  /** What a chart plots and how, for drawing it. */
  chart?: ChartSpec;
  from?: Anchor;
  to?: Anchor;
  /** Size in pixels when the object is anchored to one cell only. */
  sizePx?: { w: number; h: number };
}

export interface BlockedRegion extends Rect {
  kind: 'pivot' | 'array' | 'tableHeader';
  reason: Text;
}

export interface MergeRange extends Rect {}

/**
 * Cells the page keeps read-only, a schema's, which an edit leaves as they
 * are and counts as skipped, as it does a protected sheet's locked cells.
 */
export interface ReadOnlyCells extends Rect {
  reason: Text;
}

/** Excel's Fill Series: how each row or column of a selection is filled from the cells it starts with. */
export interface SeriesOptions {
  /** Each row of the selection is a series of its own, or each column. */
  by: 'rows' | 'columns';
  /**
   * Linear adds the step, Growth multiplies by it, Date moves a date on by
   * the step in `dateUnit`, and AutoFill continues the cells as dragging the
   * fill handle would, ignoring the step and stop.
   */
  type: 'linear' | 'growth' | 'date' | 'autofill';
  dateUnit: 'day' | 'weekday' | 'month' | 'year';
  step: number;
  /** The series ends before a value that would pass this one. */
  stop?: number;
  /**
   * Linear and Growth: fits a straight line or an exponential curve to the
   * numbers the row or column starts with, ignoring the step, and replaces
   * them with the fitted values too, as Excel does.
   */
  trend: boolean;
}

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

export interface NoteEntry {
  author?: string;
  text: string;
  /** ISO date, for threaded comments. */
  date?: string;
  /** A threaded comment's id ({GUID}), which its replies point at. */
  id?: string;
  /** Who wrote a threaded comment: their id in the workbook's list of people. */
  personId?: string;
  /** When a threaded comment was written, as the file has it: UTC, to hundredths of a second. */
  time?: string;
}

/** A note, or a threaded comment with its replies, on one cell. */
export interface CellNote {
  row: number;
  col: number;
  kind: 'note' | 'thread';
  entries: NoteEntry[];
  resolved?: boolean;
}

/** A change to a cell's note or comment thread. */
export type CommentChange =
  /** Adds a note, or changes its text. */
  | { action: 'setNote'; text: string; author?: string }
  | { action: 'startThread'; text: string; author: string }
  | { action: 'reply'; text: string; author: string }
  /** Changes the text of one comment in the thread. */
  | { action: 'editEntry'; id: string; text: string }
  /** Deletes a reply; deleting the first comment deletes the thread. */
  | { action: 'deleteEntry'; id: string }
  | { action: 'resolve'; resolved: boolean }
  /** Deletes the cell's note, or its whole thread. */
  | { action: 'delete' };

/** A cell's note or thread at one moment, for undo and redo; null when it has none. */
export interface NoteSnapshot {
  sheet: number;
  row: number;
  col: number;
  note: CellNote | null;
}

/** A hyperlink on a range of cells, as a sheet's <hyperlink> element has it. */
export interface Hyperlink {
  /** The cells it covers. */
  rect: Rect;
  /** An address outside the workbook, as written: a web page, an email or a file. */
  target?: string;
  /** A place in the workbook ("Sheet2!A1" or a defined name); with `target`, a part of that page. */
  location?: string;
  /** The text the link was given to show. */
  display?: string;
  /** Its screen tip. */
  tooltip?: string;
  /** Which of the sheet's <hyperlink> elements it is, so it is saved as it was; unset for one made here. */
  id?: number;
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
  /** Formats of whole rows and columns, used by cells that have none of their own. */
  rowStyles: Record<number, number>;
  colStyles: { min: number; max: number; style: number }[];
  merges: MergeRange[];
  listValidations: ListValidation[];
  ruleValidations: RuleValidation[];
  blocked: BlockedRegion[];
  objects: SheetObject[];
  notes: CellNote[];
  tables: TableInfo[];
  conditionalFormats: CfRule[];
  links: Hyperlink[];
  /** The range of the sheet's own filter (autoFilter), if it has one. */
  autoFilter?: Rect;
  /** The filter that hides rows, the sheet's or a table's, when it picks values the editor can show and change. */
  filter?: FilterState;
  protected: boolean;
  /** Its protection keeps objects as they are (Excel's "Edit objects" left unticked), notes and comments included. */
  objectsLocked?: boolean;
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
  | 'activeContent'
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
  | 'autoFilter'
  | 'richText'
  | 'cellImage'
  | 'formulaEngine';

export interface Finding {
  kind: FindingKind;
  /** "warn" items are surfaced in the banner when the file opens. */
  severity: 'info' | 'warn';
  title: Text;
  detail: Text;
  sheet?: string;
  where?: string;
}

/** A defined name, as the workbook has it now. */
export interface DefinedName {
  name: string;
  /** What it refers to, as file-form formula text without "=": "Data!$A$1:$A$10". */
  formula: string;
  /** Index of the sheet the name is local to, if any. */
  localSheet?: number;
  /** Hidden from Excel's Name Manager. */
  hidden?: boolean;
  comment?: string;
  /** Its place among the names the file defines, for one read from the file. */
  fileIndex?: number;
}

/** Which name: its name, in any case, and the sheet it is local to, if any. */
export interface NameRef {
  name: string;
  localSheet?: number;
}

/** A change to the workbook's names: `from` is the name as it was (absent when one is added), `to` what it becomes (null when it is deleted). */
export interface NameChange {
  from?: NameRef;
  to: DefinedName | null;
}

export interface LoadResult {
  fileName: string;
  sheets: SheetInfo[];
  /** The workbook's defined names, which links and the name box can go to. */
  names: DefinedName[];
  grids: (SheetGrid | null)[];
  styles: DisplayStyle[];
  findings: Finding[];
  date1904: boolean;
  timings: Record<string, number>;
  /** Custom document properties, such as the schema the file was last saved with. */
  properties: Record<string, string>;
  /** The workbook's structure is protected, so its sheets can't be added, deleted, renamed, moved, hidden or unhidden. */
  structureLocked?: boolean;
}

export interface CalcSample {
  sheet: string;
  ref: string;
  formula: string;
  reason: Text;
  /** Excel's saved value and the engine's, as describeValue words them. */
  excel?: Text;
  engine?: Text;
}

export interface CalcReport {
  buildMs: number;
  formulaCells: number;
  /** Formulas the engine calculates and keeps up to date. */
  liveCells: number;
  unsupported: number;
  mismatched: number;
  noCachedValue: number;
  /** What the formulas that aren't calculated use, such as "LET function", each once. */
  unsupportedFunctions: Text[];
  samples: CalcSample[];
  /** Value corrections for formulas the file had no saved value for. */
  updates: CellUpdate[];
  findings: Finding[];
  /** Why each formula is not calculated here, keyed by "sheet:row:col". */
  cellReasons: Map<string, Text>;
  /** Live recalculation is off for this workbook; no engine was built. */
  off?: boolean;
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
  /** For a HYPERLINK formula, the address it links to now. */
  link?: string;
}

/** A cell that differs from the opened file. Values are as people read them. */
export interface CellChange {
  sheet: string;
  ref: string;
  before: string;
  after: string;
  /** Only the cell's format changed. */
  formatOnly?: boolean;
  /** A change to rows, columns or where cells are, rather than to one cell; `ref` is the rows, columns or range. */
  note?: Text;
  /** A change to the cell's note or comment thread, in words; `before` and `after` are its text. */
  comment?: Text;
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
  /** Text pasted from elsewhere; `exact` has the full value of numbers the text shows rounded, where the source gave it. */
  | { type: 'pasteText'; sheet: number; row: number; col: number; values: string[][]; exact?: (number | null)[][]; readOnly?: ReadOnlyCells[] }
  | { type: 'clear'; sheet: number; rects: Rect[]; readOnly?: ReadOnlyCells[] }
  /**
   * Copy cells within the workbook, tiling the source over the destination:
   * everything, or only the values (formulas' results), the formulas or the
   * formats, and turned on its side with `transpose`.
   */
  | { type: 'copy'; srcSheet: number; src: Rect; sheet: number; dst: Rect; what?: PastePart; transpose?: boolean; readOnly?: ReadOnlyCells[] }
  /**
   * Drag-fill from the source into the destination, extending series and
   * continuing the `lists` (by default, English day and month names).
   */
  | { type: 'fill'; sheet: number; src: Rect; dst: Rect; lists?: string[][]; readOnly?: ReadOnlyCells[] }
  /** Fills each row or column of `rect` with a series, as Excel's Fill Series dialog does. */
  | { type: 'series'; sheet: number; rect: Rect; series: SeriesOptions; lists?: string[][]; readOnly?: ReadOnlyCells[] }
  /** Put cells, their notes and threads, the layout and the names back to an exact earlier state (undo/redo). */
  | { type: 'restore'; cells: CellSnapshot[]; notes?: NoteSnapshot[]; layout?: LayoutChange[]; names?: NameChange[] }
  /** Adds, changes or deletes a cell's note or comment thread. */
  | { type: 'comment'; sheet: number; row: number; col: number; change: CommentChange }
  /** Text for scattered cells, parsed like typed input (used by schema hooks). */
  | { type: 'inputs'; sheet: number; cells: { row: number; col: number; text: string }[] }
  /**
   * Reorders the rows of `rect` by the given columns, as Excel's sort does:
   * cells move with their formats, and formulas move as if copied.
   */
  | { type: 'sort'; sheet: number; rect: Rect; keys: { col: number; descending: boolean }[]; header: boolean }
  /** Changes how a sheet is laid out, rather than what its cells hold. */
  | { type: 'layout'; sheet: number; op: LayoutOp }
  /**
   * Changes the cells' formats, keeping what they hold: `borders` puts the
   * edge on the given sides of the range's cells; `clear` returns them to
   * the workbook's default format.
   */
  | { type: 'format'; sheet: number; rects: Rect[]; change: FormatChange; borders?: { where: BorderPlacement; edge?: BorderEdge }; clear?: boolean }
  /**
   * Adds, changes or deletes a defined name. Renaming one renames it in the
   * formulas that use it, as Excel does.
   */
  | { type: 'name'; change: NameChange }
  /** Gives the cells that hold something Excel's Hyperlink look, as inserting a link does, or takes it away, as removing one does. */
  | { type: 'linkFormat'; sheet: number; rects: Rect[]; on: boolean }
  /** Several operations applied in order and undone as one. */
  | { type: 'batch'; ops: EditOp[] };

/**
 * A change to a sheet's layout, or to its hyperlinks. Sizes are in pixels,
 * null being the sheet's default; each lists only the rows or columns it
 * changes, so the same kind of change, with the values from before, undoes it.
 */
export type LayoutOp =
  | { kind: 'colWidths'; widths: Record<number, number | null> }
  | { kind: 'rowHeights'; heights: Record<number, number | null> }
  | { kind: 'hidden'; axis: 'rows' | 'cols'; lines: Record<number, boolean> }
  /** Rows above and columns left of the panes that scroll; 0 and 0 unfreezes. */
  | { kind: 'freeze'; rows: number; cols: number }
  /** A filter, or none, and the rows it hides. */
  | { kind: 'filter'; filter: FilterState | null; rows: number[] }
  | { kind: 'merges'; add: Rect[]; remove: Rect[] }
  /** Hyperlinks added and removed; the same change the other way round undoes it. */
  | { kind: 'links'; add: Hyperlink[]; remove: Hyperlink[] };

/**
 * A change to cells' formats; what it leaves out stays as it was. Colours
 * are "#RRGGBB"; null is the automatic font colour, no fill, the general
 * alignment, or no border on that edge.
 */
export interface FormatChange {
  numFmt?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  color?: string | null;
  fill?: string | null;
  hAlign?: 'left' | 'center' | 'right' | 'justify' | null;
  vAlign?: 'top' | 'center' | 'bottom' | null;
  wrap?: boolean;
  indent?: number;
  border?: Partial<Record<'left' | 'right' | 'top' | 'bottom', BorderEdge | null>>;
}

/** What a paste takes from what was copied, as Excel's Paste Special offers. */
export type PastePart = 'all' | 'values' | 'formulas' | 'formats';

/** Which edges of a range's cells a border goes on. */
export type BorderPlacement = 'all' | 'outside' | 'inside' | 'top' | 'bottom' | 'left' | 'right' | 'none';

/**
 * A filter by values on a range whose first row holds the headings: for
 * each filtered column (a sheet column), the values whose rows show, as the
 * cells show them ('' for blanks).
 */
export interface FilterState {
  rect: Rect;
  columns: Record<number, string[]>;
  /** The table whose filter it is, when the range is a table's; its part holds the filter. */
  table?: string;
}

export interface LayoutChange {
  sheet: number;
  op: LayoutOp;
}

export interface EditResult {
  updates: CellUpdate[];
  undo: CellSnapshot[];
  redo: CellSnapshot[];
  skipped: number;
  skippedReasons: Text[];
  warnings: Text[];
  newStyles: { index: number; style: DisplayStyle }[];
  editedCount: number;
  /** Set when the input could not be accepted at all, e.g. a malformed formula. */
  rejected?: Text;
  /** Layout changes made, in order, and the ones that undo them, in the order to apply them. */
  layout?: { applied: LayoutChange[]; undo: LayoutChange[] };
  /** Notes and threads as they were before the edit, for undo. */
  noteUndo?: NoteSnapshot[];
  /** Notes and threads as the edit left them, to show and for redo. */
  noteRedo?: NoteSnapshot[];
  /** Changes made to the names, those that undo them, and every name as it now is. */
  names?: { applied: NameChange[]; undo: NameChange[]; list: DefinedName[] };
}
