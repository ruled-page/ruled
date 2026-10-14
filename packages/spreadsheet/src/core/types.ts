import type { Text } from '../i18n/message';
import type { ChartSpec } from './charts';
import type { NewChartKind } from './chartWriter';
import type { ChartEdit } from './chartEdit';
import type { PivotSetup } from './pivotTable';

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
  /**
   * Shows a value spilled from a dynamic array formula in another cell; the
   * cell holds nothing of its own. In a file's grid, a cell holding only the
   * value Excel saved for a spill.
   */
  Spilled: 128,
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
  /** Excel's line styles: solid, dashes, dots, a hair's fine dots, dash-dot and dash-dot-dot, slanted dash-dot, and double. */
  pattern: 'solid' | 'dashed' | 'dotted' | 'double' | 'hair' | 'dashDot' | 'dashDotDot' | 'slantDashDot';
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
  /** Whether its heading row shows filter buttons, as a table with an <autoFilter> does. */
  filterButtons?: boolean;
  /** Its filter picks rows in a way the editor can't change, so it is kept as the file has it. */
  filterKept?: boolean;
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
  /** The file's format for it (its dxfId), kept when the rule is written again; a new rule's format is added. */
  dxfId?: number;
  /** The id of its Excel 2010 settings, kept in the sheet's extensions. */
  extId?: string;
  /** Written only in the sheet's Excel 2010 extensions, which keep it as they have it: it can be moved, re-ranged or deleted, not changed. */
  x14Only?: boolean;
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
  /** What the sheet's sparklines draw, for the current values. */
  sparklines?: SparkCell[];
}

/** One of Excel's sparkline groups: sparklines drawn alike, each in its cell from a range of values. */
export interface SparklineGroup {
  /** `stacked` is Excel's win/loss. */
  type: 'line' | 'column' | 'stacked';
  colors: Record<'series' | 'negative' | 'axis' | 'markers' | 'first' | 'last' | 'high' | 'low', string>;
  /** Points marked in their own colour, and the axis at zero. */
  show: Record<'markers' | 'high' | 'low' | 'first' | 'last' | 'negative' | 'axis', boolean>;
  /** In points. */
  lineWeight: number;
  /** Empty cells as a gap in the line, as zero, or with the line carried across them. */
  empty: 'gap' | 'zero' | 'span';
  /** The range drawn against: each sparkline's own, the group's, or set. */
  minAxis: 'individual' | 'group' | 'custom';
  maxAxis: 'individual' | 'group' | 'custom';
  manualMin?: number;
  manualMax?: number;
  /** Values in hidden rows and columns are drawn too. */
  displayHidden: boolean;
  /** Points run from right to left. */
  rightToLeft: boolean;
  /** Each sparkline's cell and the range (a formula, as the file has it) its values come from. */
  lines: { row: number; col: number; source: string }[];
}

/** What a sparkline draws in its cell. */
export interface SparkCell {
  row: number;
  col: number;
  /** Index into the sheet's sparkline groups. */
  group: number;
  /** Null for an empty cell, or one that isn't a number. */
  values: (number | null)[];
  /** The range its values are drawn against. */
  min: number;
  max: number;
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
  /** Excel 365's checkbox: TRUE shows ticked, FALSE or nothing unticked. */
  checkbox?: boolean;
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
  /** Where it is in the file: its sheet's drawing part, and which of the part's anchors it is. Missing for objects added since. */
  source?: { drawing: string; anchor: number };
  /** A chart's part in the file. */
  chartPart?: string;
  /** A chart added or changed in the editor: its chart part, as it is saved. */
  chartXml?: string;
  /** A picture: its image part, and how it is cropped (fractions of the image cut from each edge), turned (degrees clockwise) and flipped. */
  picture?: { part: string; crop?: { l: number; t: number; r: number; b: number }; rotate?: number; flipH?: boolean; flipV?: boolean };
  /** A shape or text box of a kind the editor draws, as it is drawn. */
  shape?: ShapeSpec;
}

/** A shape as drawn: one of Excel's preset outlines, its fill, its line and its text. Colours are resolved, sizes in pixels. */
export interface ShapeSpec {
  /** The preset, by the name Excel's files use (rect, roundRect, rightArrow…). */
  geometry: string;
  /** The preset's adjustments (a rounded corner's radius, an arrow's head), by name, in 100,000ths. */
  adjust?: Record<string, number>;
  /** None when missing. */
  fill?: string;
  /** None when missing; `dash` is Excel's preset name (dash, sysDot…), and a line's ends may have arrowheads. */
  line?: { color: string; width: number; dash?: string; head?: 'arrow' | 'oval'; tail?: 'arrow' | 'oval' };
  /** Degrees clockwise. */
  rotate?: number;
  flipH?: boolean;
  flipV?: boolean;
  text?: ShapeText;
}

/** The text in a shape: its paragraphs of runs, where it sits, and how far in from the box's edges (left, top, right, bottom). */
export interface ShapeText {
  paragraphs: {
    align: 'left' | 'center' | 'right' | 'justify';
    /** A run's size in pixels; its font missing for the workbook's own. A run of "\n" is a line break. */
    runs: { text: string; size?: number; bold?: boolean; italic?: boolean; underline?: boolean; color?: string; font?: string }[];
    /** An empty paragraph's size. */
    size?: number;
  }[];
  anchor: 'top' | 'middle' | 'bottom';
  insets: [number, number, number, number];
  nowrap?: boolean;
  vertical?: boolean;
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
  /** Whether Excel shows the input message when the cell is selected (unset: it does, where there is one). */
  showPrompt?: boolean;
  /** Whether Excel shows an alert for values that break the rule. */
  showError: boolean;
  errorStyle: 'stop' | 'warning' | 'information';
  errorTitle?: string;
  error?: string;
  promptTitle?: string;
  prompt?: string;
}

/**
 * A non-list Excel data validation rule: number, date, time, text length or
 * custom formula, or "any value", which only shows an input message.
 */
export interface RuleValidation extends ValidationAlert {
  rects: Rect[];
  kind: 'any' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'custom';
  operator: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';
  /** Bounds or the custom formula, relative to the rule's top-left cell. */
  formula1: string;
  formula2?: string;
}

/** An Excel "list" data validation rule on a sheet. */
export interface ListValidation extends ValidationAlert {
  rects: Rect[];
  source: ListSource;
  /** The source as the file writes it (`"a,b"`, `$A$1:$A$5`, a name or a formula), relative to the rule's top-left cell. */
  formula?: string;
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

/** One direction of a sheet's outline: how deep in groups each row (or column) is. */
export interface OutlineAxis {
  /** Row (or column) → its outline level, 1 to 7, for those in a group. */
  levels: Record<number, number>;
  /** Rows (or columns) marked as the summary of a collapsed group. */
  collapsed: number[];
  /** A group's summary row is below it (Excel's default), or its summary column right of it; false when above or left. */
  summaryAfter: boolean;
}

/** A pivot table as Excel's PivotTable Fields pane shows it. */
export interface PivotInfo {
  /** Its definition's part, which names it in changes to it. */
  path: string;
  name: string;
  /** The cells it covers: its table, and its report filters' above it. */
  regions: Rect[];
  /** Why its fields can't be changed here, when they can't. */
  kept?: Text;
  fields: PivotFieldInfo[];
  setup: PivotSetup;
}

export interface PivotFieldInfo {
  name: string;
  /** Its values are all numbers, so ticking it sums it, as Excel's pane does. */
  numeric: boolean;
  /** Its items as the source has them, for its filter, up to a limit. */
  items: { key: string; label: string }[];
  /** It has more items than are listed. */
  more?: boolean;
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
  /** Row index → height in pixels, for rows with a height of their own. */
  rowHeights: Record<number, number>;
  /**
   * Rows whose height Excel fitted to their contents rather than one set by
   * hand (no customHeight), which are fitted again as their cells change,
   * as are rows with no height of their own.
   */
  fittedRows?: Record<number, true>;
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
  /** The sheet has validation rules the editor can't read, so its rules are kept as the file has them and can't be changed. */
  validationsKept?: boolean;
  blocked: BlockedRegion[];
  objects: SheetObject[];
  /** Its pivot tables, as Excel's PivotTable Fields pane shows them. */
  pivots?: PivotInfo[];
  notes: CellNote[];
  tables: TableInfo[];
  conditionalFormats: CfRule[];
  /** Excel 2010's sparkline groups. */
  sparklines?: SparklineGroup[];
  links: Hyperlink[];
  /** How its rows and columns are grouped, when they are or the file says where groups' summaries go. */
  outline?: Record<'rows' | 'cols', OutlineAxis>;
  /** The range of the sheet's own filter (autoFilter), if it has one. */
  autoFilter?: Rect;
  /** The sheet's own filter picks rows in a way the editor can't change (dates in a period, colours, icons), so it is kept as the file has it. */
  filterKept?: boolean;
  /** The filter that hides rows, the sheet's or a table's, when it picks values the editor can show and change. */
  filter?: FilterState;
  protected: boolean;
  /** Its protection keeps objects as they are (Excel's "Edit objects" left unticked), notes and comments included. */
  objectsLocked?: boolean;
  /** Its protection still lets rows, or columns, be formatted (Excel's "Format rows" and "Format columns"): hidden, shown and grouped. */
  formatAllowed?: { rows: boolean; cols: boolean };
  showGridLines: boolean;
  /** The sheet reads right to left, column A at the right, as Excel shows sheets in Arabic or Hebrew. */
  rightToLeft?: boolean;
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
  | 'encryption'
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
  /** The workbook was opened with a password, and is saved encrypted with it. */
  encrypted?: boolean;
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
  /** The file's dynamic array formulas that are calculated here: their saved spill ranges, no longer locked. */
  liveArrays?: { sheet: number; rect: Rect }[];
  /** Where circular references close, as Excel lists them; none when there are none. */
  circular?: { sheet: number; row: number; col: number }[];
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
  /**
   * The range a dynamic array formula's results spill into, from its own
   * cell: given for the formula's cell and for each cell it spills into.
   * For a formula showing #SPILL!, the range its results need.
   */
  spill?: Rect;
  /**
   * The range an array formula entered with Ctrl+Shift+Enter fills, which
   * keeps its size and is changed only as a whole: given for each of its
   * cells, the formula being in the first.
   */
  array?: Rect;
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

/**
 * What a sort orders by: a column's values (a row's, sorting across), smallest
 * or A first unless `descending`, or the order of a custom list such as the
 * months, values not in the list coming after the ones that are.
 */
/**
 * A level of a sort: by a column's (or a row's) values, in a custom list's
 * order where `list` is given; or, with `marked`, the lines whose cell has a
 * colour or icon first, or last when descending, the rest keeping their order.
 */
export type SortKey = ({ col: number } | { row: number }) & { descending: boolean; list?: readonly string[]; marked?: readonly number[] };

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
   * cells move with their formats, and formulas move as if copied. With
   * `across`, its columns are reordered by the given rows instead, as Excel's
   * "Sort left to right" does, and a header row doesn't apply. Text compares
   * ignoring case unless `caseSensitive`, which puts lower case first.
   */
  | { type: 'sort'; sheet: number; rect: Rect; keys: SortKey[]; header: boolean; caseSensitive?: boolean; across?: boolean }
  /** Excel's Refresh (one sheet's pivot tables) and Refresh All (every sheet's): each worked out again from its source. */
  | { type: 'refreshPivots'; sheet?: number }
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
  /** `fitted`: the rows among `heights` fitted to their contents, as AutoFit sizes them, rather than set by hand. */
  | { kind: 'rowHeights'; heights: Record<number, number | null>; fitted?: number[] }
  | { kind: 'hidden'; axis: 'rows' | 'cols'; lines: Record<number, boolean> }
  /** Rows above and columns left of the panes that scroll; 0 and 0 unfreezes. */
  | { kind: 'freeze'; rows: number; cols: number }
  /** A filter, or none, and the rows it hides. */
  | { kind: 'filter'; filter: FilterState | null; rows: number[] }
  | { kind: 'merges'; add: Rect[]; remove: Rect[] }
  /** Rows' or columns' outline levels (0: in no group), and whether each is marked as a collapsed group's summary. */
  | { kind: 'outline'; axis: 'rows' | 'cols'; levels: Record<number, number>; collapsed: Record<number, boolean> }
  /** Hyperlinks added and removed; the same change the other way round undoes it. */
  | { kind: 'links'; add: Hyperlink[]; remove: Hyperlink[] }
  /** The sheet's data validation rules, all of them, as Excel's Data Validation dialog leaves them; the ones before undo it. */
  | { kind: 'validations'; lists: ListValidation[]; rules: RuleValidation[] }
  /** The sheet's conditional formatting rules, all of them, in priority order; the ones before undo it. */
  | { kind: 'conditionalFormats'; rules: CfRule[] }
  /** The sheet's charts, pictures and shapes, where they now sit; any left out are deleted. */
  | { kind: 'objects'; objects: SheetObject[] }
  /** Excel's Insert Chart: a chart of `chart`'s kind plotting `rect`, placed between `from` and `to`; `title` is Excel's "Chart Title" in the user's language. */
  | { kind: 'insertChart'; chart: NewChartKind; rect: Rect; from: Anchor; to: Anchor; title: string }
  /** Excel's chart dialogs: the chart with key `key` changed as `edit` says; `title` is Excel's "Chart Title" in the user's language. */
  | { kind: 'editChart'; key: string; edit: ChartEdit; title: string }
  /** Excel's PivotTable Fields pane: the pivot table whose part is `path` with its fields set as `setup` says; the setup before undoes it. */
  | { kind: 'pivot'; path: string; setup: PivotSetup }
  /**
   * Excel's Insert PivotTable: an empty pivot table named `name` (the next
   * PivotTableN if not given) over `source`, a range such as Data!$A$1:$E$101
   * or a table's or a name's name, with its top-left cell at `row` and `col`.
   * `path` is its part's, when it is put back; `removePivot` undoes it.
   */
  | { kind: 'insertPivot'; source: string; row: number; col: number; name?: string; path?: string }
  /** A pivot table inserted here taken out again. */
  | { kind: 'removePivot'; path: string };

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
  /** Excel 365's checkbox format on (or off). */
  checkbox?: boolean;
}

/** What a paste takes from what was copied, as Excel's Paste Special offers. */
export type PastePart = 'all' | 'values' | 'formulas' | 'formats';

/** Which edges of a range's cells a border goes on. */
export type BorderPlacement = 'all' | 'outside' | 'inside' | 'top' | 'bottom' | 'left' | 'right' | 'none';

/**
 * A filter by values on a range whose first row holds the headings: for
 * each filtered column (a sheet column), the values whose rows show, as the
 * cells show them ('' for blanks). With no columns, the filter's buttons are
 * there but nothing is filtered, as after Excel's Filter command or Clear.
 */
export interface FilterState {
  rect: Rect;
  columns: Record<number, string[]>;
  /** Columns filtered by a condition rather than by values, as Excel's Text, Number and Date Filters do. */
  conditions?: Record<number, FilterCondition>;
  /** The table whose filter it is, when the range is a table's; its part holds the filter. */
  table?: string;
}

/** How a custom filter compares a cell with its value, as Excel stores it. */
export type FilterOperator = 'equal' | 'notEqual' | 'greaterThan' | 'greaterThanOrEqual' | 'lessThan' | 'lessThanOrEqual';

/**
 * A column's filter by condition, as Excel keeps it:
 * - `custom`: one or two rules, both or either to hold. A rule's value is
 *   text, with Excel's wildcards for "begins with" (`abc*`), "contains"
 *   (`*abc*`) and the like, or a number (a date's serial number).
 * - `top10`: the largest (or smallest) so many numbers, or so many percent of
 *   them, with the cut-off the filter was applied with.
 * - `average`: the numbers above (or below) the average, with the average it
 *   was applied with.
 * - `color`: cells whose fill (or, with `font`, text) is a colour, or has
 *   none (`color` missing); `dxfId` is the format in the file that holds it.
 * - `icon`: cells that conditional formatting gives an icon of a set, or no
 *   icon (`index` missing).
 * - `period`: dates in a period, as Excel's Date Filters pick them; one
 *   counted from today keeps the dates it was applied with, `from` up to but
 *   not including `to`.
 */
export type FilterCondition =
  | { kind: 'custom'; and: boolean; rules: { operator: FilterOperator; value: string }[] }
  | { kind: 'top10'; top: boolean; percent: boolean; count: number; threshold?: number }
  | { kind: 'average'; above: boolean; average?: number }
  | { kind: 'color'; font: boolean; color?: string; dxfId?: number }
  | { kind: 'icon'; set: string; index?: number }
  | { kind: 'period'; period: DatePeriod; from?: number; to?: number };

/** Excel's Date Filters by period, by the names its files use. */
export type DatePeriod =
  | 'tomorrow'
  | 'today'
  | 'yesterday'
  | 'nextWeek'
  | 'thisWeek'
  | 'lastWeek'
  | 'nextMonth'
  | 'thisMonth'
  | 'lastMonth'
  | 'nextQuarter'
  | 'thisQuarter'
  | 'lastQuarter'
  | 'nextYear'
  | 'thisYear'
  | 'lastYear'
  | 'yearToDate'
  | `Q${1 | 2 | 3 | 4}`
  | `M${1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 | 11 | 12}`;

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
  /** Where circular references close now, when formulas are being calculated. */
  circular?: { sheet: number; row: number; col: number }[];
  /** Sheets whose read-only regions changed (a pivot table grown or shrunk), with all of them as they now are. */
  blocked?: { sheet: number; regions: BlockedRegion[] }[];
  /** Sheets' pivot tables, where they were worked out again or changed. */
  pivots?: { sheet: number; list: PivotInfo[] }[];
}
