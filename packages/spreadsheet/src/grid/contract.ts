import type { SheetView } from '../state/sheetView';
import type { Selection, WorkbookStore } from '../state/store';

/**
 * What the editor and the grid agree on. The grid draws one sheet, frozen
 * panes included, and owns scrolling and where things are on screen; what
 * is drawn over it (charts, outline bars, touch handles) asks it through
 * GridHandle. Everything crosses in sheet rows and columns and in client
 * pixels, never in the grid's own indices, so hidden lines and frozen panes
 * stay the grid's own business.
 */

/** A box on screen, in client pixels. */
export interface Box {
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * A band of rows that scrolls as one: the frozen rows, when the sheet has
 * any in view, then the rows below them.
 */
export interface Pane {
  /** The first and last sheet rows the band holds, in view or not. */
  range: [number, number];
  /** The sheet rows in view, top to bottom, partly shown ones included. */
  rows: number[];
  /** The band below the column letters, the row numbers and scrollbars included. */
  box: Box;
  /** The room its scrollbars take at its right and bottom edges. */
  scrollbars: { right: number; bottom: number };
}

/** What a grid shows now. */
export interface GridView {
  /** Where the column letters meet the row numbers. */
  corner: Box | undefined;
  panes: Pane[];
  /** The sheet columns in view, the frozen ones first, partly shown ones included. */
  cols: number[];
  /** How many of `cols` are frozen. */
  frozenCols: number;
  /** Where the frozen columns end on screen, or the row numbers without any: columns that scroll pass under them. */
  freezeX: number;
  /** The sheet reads right to left: its columns run leftwards from the row numbers, at the right. */
  rtl?: boolean;
}

/** What the editor asks of a grid. */
export interface GridHandle {
  view(): GridView;
  /**
   * Where a cell is drawn, or would be: any column, and any row of its band
   * from the first in view down. A merged cell's own box, not the merge's.
   */
  cellBox(row: number, col: number): Box | undefined;
  /** A row's number, or a column's letters, on screen. */
  headingBox(axis: 'rows' | 'cols', line: number): Box | undefined;
  /** The cell under a point, or the nearest in view beyond an edge: for drags. */
  cellNear(x: number, y: number): { row: number; col: number } | undefined;
  /** Scrolls the rows and columns that scroll. */
  scrollBy(dx: number, dy: number): void;
  /** Calls `fn` after the grid scrolls, resizes or lays out again, until the function returned is called. */
  subscribe(fn: () => void): () => void;
}

export interface SurfaceProps {
  store: WorkbookStore;
  view: SheetView;
  /** Changes whenever cell data or styles change, so the grid redraws. */
  dataVersion: number;
  /** Changes whenever schema issues change, so markers redraw. */
  schemaVersion: number;
  selection: Selection;
  /** Given the grid's handle once it can answer, and undefined once it can't. */
  onHandle: (handle: GridHandle | undefined) => void;
}

/** A move once editing ends: across and down, by a cell. */
export type EditMove = readonly [-1 | 0 | 1, -1 | 0 | 1];

/** What the cell editors are given. */
export interface CellEditorProps {
  /** The text being edited: the cell's, or what was typed to open the editor. */
  text: string;
  onText: (text: string) => void;
  /** Ends editing, entering `text` (or, undefined, leaving the cell as it was), then moves. */
  finish: (text: string | undefined, move: EditMove) => void;
  /** Opened by typing, so `text` is what was typed. */
  typed: boolean;
  /** Opened with its text selected, rather than the caret at its end. */
  selectAll: boolean;
  readOnly: boolean;
  /** The cell wraps its text, so the editor does too, at the cell's width. */
  wrap: boolean;
  /** The cell's box. */
  box: Box;
  /** The values a list cell allows. */
  options?: readonly string[];
  /** A class for popups the editor opens, so clicks in them don't end editing. */
  popupClass: string;
}

/**
 * The element cell editors and their popups open in, one for the page,
 * inside the editor's style scope.
 */
export const PORTAL_ID = 'wbe-portal';
