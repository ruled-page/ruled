import { msg, type Text } from '../i18n/message';
import { colToLetters, MAX_COLS, rangeToA1, toA1 } from './refs';
import type { CellChange, LayoutOp, Rect, SheetInfo } from './types';
import { elements, escapeXml, firstElement, parseAttrs } from './xml';

const sameRect = (a: Rect, b: Rect) => a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;

/** Sets each entry of `values`, returning what each was before. */
function setEach<T>(values: Record<number, T>, get: (k: number) => T, set: (k: number, v: T) => void): Record<number, T> {
  const before: Record<number, T> = {};
  for (const [k, v] of Object.entries(values)) {
    const n = Number(k);
    before[n] = get(n);
    set(n, v);
  }
  return before;
}

/**
 * One sheet's layout as edited: the size and visibility set for each row and
 * column changed, and its frozen panes and merged cells once changed, over
 * the layout the file had.
 */
export class SheetLayout {
  readonly colWidths = new Map<number, number | null>();
  readonly rowHeights = new Map<number, number | null>();
  readonly hiddenCols = new Map<number, boolean>();
  readonly hiddenRows = new Map<number, boolean>();
  frozen?: { rows: number; cols: number };
  merges?: Rect[];
  /** The file's layout, copied, as what changes are measured against. */
  private readonly file: Pick<SheetInfo, 'colWidths' | 'rowHeights' | 'frozenRows' | 'frozenCols' | 'merges'>;
  private readonly fileHidden: Record<'rows' | 'cols', Set<number>>;

  constructor(info: SheetInfo) {
    this.file = { colWidths: { ...info.colWidths }, rowHeights: { ...info.rowHeights }, frozenRows: info.frozenRows, frozenCols: info.frozenCols, merges: [...info.merges] };
    this.fileHidden = { rows: new Set(info.hiddenRows), cols: new Set(info.hiddenCols) };
  }

  get changed(): boolean {
    return this.colWidths.size + this.rowHeights.size + this.hiddenCols.size + this.hiddenRows.size > 0 || !!this.frozen || !!this.merges;
  }

  colWidth(col: number): number | null {
    return this.colWidths.has(col) ? this.colWidths.get(col)! : (this.file.colWidths[col] ?? null);
  }

  rowHeight(row: number): number | null {
    return this.rowHeights.has(row) ? this.rowHeights.get(row)! : (this.file.rowHeights[row] ?? null);
  }

  hidden(axis: 'rows' | 'cols', line: number): boolean {
    const set = axis === 'rows' ? this.hiddenRows : this.hiddenCols;
    return set.has(line) ? set.get(line)! : this.fileHidden[axis].has(line);
  }

  get freeze(): { rows: number; cols: number } {
    return this.frozen ?? { rows: this.file.frozenRows, cols: this.file.frozenCols };
  }

  get mergeList(): Rect[] {
    return this.merges ?? this.file.merges;
  }

  /** Makes the change, returning the change that undoes it. */
  apply(op: LayoutOp): LayoutOp {
    switch (op.kind) {
      case 'colWidths':
        return { kind: 'colWidths', widths: setEach(op.widths, (c) => this.colWidth(c), (c, px) => this.colWidths.set(c, px)) };
      case 'rowHeights':
        return { kind: 'rowHeights', heights: setEach(op.heights, (r) => this.rowHeight(r), (r, px) => this.rowHeights.set(r, px)) };
      case 'hidden': {
        const set = op.axis === 'rows' ? this.hiddenRows : this.hiddenCols;
        return { kind: 'hidden', axis: op.axis, lines: setEach(op.lines, (l) => this.hidden(op.axis, l), (l, h) => set.set(l, h)) };
      }
      case 'freeze': {
        const was = this.freeze;
        this.frozen = { rows: op.rows, cols: op.cols };
        return { kind: 'freeze', ...was };
      }
      case 'merges': {
        const before = this.mergeList;
        const removed = before.filter((m) => op.remove.some((r) => sameRect(r, m)));
        const kept = before.filter((m) => !removed.includes(m));
        const added = op.add.filter((a) => !kept.some((m) => sameRect(m, a)));
        this.merges = [...kept, ...added];
        return { kind: 'merges', add: removed, remove: added };
      }
    }
  }

  /** How the layout differs from the file's, for the list of changes. */
  notes(sheet: string): CellChange[] {
    const out: CellChange[] = [];
    const note = (ref: string, text: Text) => out.push({ sheet, ref, before: '', after: '', note: text });
    const widths = [...this.colWidths].filter(([c, px]) => px !== (this.file.colWidths[c] ?? null)).map(([c]) => c);
    if (widths.length) note(lineRefs(widths, 'cols'), msg('core.change.columnWidths', { count: widths.length }));
    const heights = [...this.rowHeights].filter(([r, px]) => px !== (this.file.rowHeights[r] ?? null)).map(([r]) => r);
    if (heights.length) note(lineRefs(heights, 'rows'), msg('core.change.rowHeights', { count: heights.length }));
    for (const axis of ['rows', 'cols'] as const) {
      const hid: number[] = [];
      const shown: number[] = [];
      for (const [line, hidden] of axis === 'rows' ? this.hiddenRows : this.hiddenCols) {
        if (hidden !== this.fileHidden[axis].has(line)) (hidden ? hid : shown).push(line);
      }
      if (hid.length) note(lineRefs(hid, axis), msg(axis === 'rows' ? 'core.change.hidRows' : 'core.change.hidColumns', { count: hid.length }));
      if (shown.length) note(lineRefs(shown, axis), msg(axis === 'rows' ? 'core.change.shownRows' : 'core.change.shownColumns', { count: shown.length }));
    }
    const f = this.frozen;
    if (f && (f.rows !== this.file.frozenRows || f.cols !== this.file.frozenCols)) {
      if (f.rows || f.cols) note(toA1(f.rows, f.cols), msg('core.change.froze', { cell: toA1(f.rows, f.cols) }));
      else note('', msg('core.change.unfroze'));
    }
    if (this.merges) {
      for (const m of this.merges) if (!this.file.merges.some((x) => sameRect(x, m))) note(rangeToA1(m), msg('core.change.merged'));
      for (const m of this.file.merges) if (!this.merges.some((x) => sameRect(x, m))) note(rangeToA1(m), msg('core.change.unmerged'));
    }
    return out;
  }
}

/** Rows or columns as Excel names them, in runs: "3:5, 9:9" or "B:D, F:F", the first ten runs only. */
function lineRefs(lines: number[], axis: 'rows' | 'cols'): string {
  const name = (n: number) => (axis === 'rows' ? String(n + 1) : colToLetters(n));
  const sorted = [...lines].sort((a, b) => a - b);
  const runs: string[] = [];
  for (let i = 0; i < sorted.length; ) {
    let j = i;
    while (j + 1 < sorted.length && sorted[j + 1] === sorted[j] + 1) j++;
    runs.push(`${name(sorted[i])}:${name(sorted[j])}`);
    i = j + 1;
  }
  return runs.length > 10 ? `${runs.slice(0, 10).join(', ')}, …` : runs.join(', ');
}

// ---------------------------------------------------------------------------
// Writing the layout into a worksheet part

/** Points, as row heights are stored, from pixels. */
const points = (px: number) => Math.round(px * 75) / 100;
/** Characters of the default font, as column widths are stored, from pixels; in 256ths, as Excel writes them. */
const characters = (px: number) => Math.round((px / 7) * 256) / 256;

/** The <row> attributes to set (null: remove) for each row whose height or visibility changed. */
export function rowAttributes(layout: SheetLayout): Map<number, Record<string, string | null>> {
  const out = new Map<number, Record<string, string | null>>();
  const set = (row: number, attrs: Record<string, string | null>) => out.set(row, { ...out.get(row), ...attrs });
  for (const [row, px] of layout.rowHeights) set(row, px === null ? { ht: null, customHeight: null } : { ht: String(points(px)), customHeight: '1' });
  for (const [row, hidden] of layout.hiddenRows) set(row, { hidden: hidden ? '1' : null });
  return out;
}

/** The worksheet with its <cols>, frozen pane and <mergeCells> as the layout has them. */
export function writeLayout(xml: string, p: string, layout: SheetLayout): string {
  if (layout.colWidths.size || layout.hiddenCols.size) xml = writeCols(xml, p, layout);
  if (layout.frozen) xml = writePane(xml, p, layout.frozen);
  if (layout.merges) xml = writeMerges(xml, p, layout.merges);
  return xml;
}

function writeCols(xml: string, p: string, layout: SheetLayout): string {
  const data = firstElement(xml, 'sheetData');
  if (!data) return xml;
  const cols = firstElement(xml, 'cols', 0, data.start);
  // Each column's attributes, from the file's ranges; ranges that overlap one before them are invalid and left out.
  const at: (Record<string, string> | undefined)[] = new Array(MAX_COLS);
  let last = -1;
  for (const c of cols?.inner ? elements(cols.inner, 'col') : []) {
    const { min, max, ...rest } = parseAttrs(c.attrs);
    const a = Number(min) - 1;
    const b = Math.min(Number(max) - 1, MAX_COLS - 1);
    if (!(a > last && b >= a)) continue;
    for (let i = a; i <= b; i++) at[i] = rest;
    last = b;
  }
  const changed = new Set([...layout.colWidths.keys(), ...layout.hiddenCols.keys()]);
  for (const col of changed) {
    if (col < 0 || col >= MAX_COLS) continue;
    const attrs = { ...at[col] };
    if (layout.colWidths.has(col)) {
      const px = layout.colWidths.get(col)!;
      delete attrs.width;
      delete attrs.customWidth;
      delete attrs.bestFit;
      if (px !== null) Object.assign(attrs, { width: String(characters(px)), customWidth: '1' });
    }
    if (layout.hiddenCols.has(col)) {
      delete attrs.hidden;
      if (layout.hiddenCols.get(col)) attrs.hidden = '1';
    }
    at[col] = attrs;
  }
  // Back into runs of columns with the same attributes.
  const out: string[] = [];
  for (let i = 0; i < MAX_COLS; ) {
    const attrs = at[i];
    const key = attrs && Object.keys(attrs).length ? JSON.stringify(attrs) : '';
    let j = i;
    while (j + 1 < MAX_COLS && (at[j + 1] === attrs || (key && at[j + 1] && JSON.stringify(at[j + 1]) === key))) j++;
    if (key) {
      const rest = Object.entries(attrs!).map(([k, v]) => ` ${k}="${escapeXml(v)}"`).join('');
      out.push(`<${p}col min="${i + 1}" max="${j + 1}"${rest}/>`);
    }
    i = j + 1;
  }
  const el = out.length ? `<${p}cols>${out.join('')}</${p}cols>` : '';
  if (cols) return xml.slice(0, cols.start) + el + xml.slice(cols.end);
  return xml.slice(0, data.start) + el + xml.slice(data.start);
}

function writePane(xml: string, p: string, { rows, cols }: { rows: number; cols: number }): string {
  const data = firstElement(xml, 'sheetData');
  if (!data) return xml;
  const pane =
    rows || cols
      ? `<${p}pane${cols ? ` xSplit="${cols}"` : ''}${rows ? ` ySplit="${rows}"` : ''} topLeftCell="${toA1(rows, cols)}" activePane="${rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight'}" state="frozen"/>`
      : '';
  const views = firstElement(xml, 'sheetViews', 0, data.start);
  const view = views?.inner ? firstElement(views.inner, 'sheetView') : undefined;
  if (!views || !view) {
    if (!pane) return xml;
    // After <sheetPr> and <dimension>, the only elements that come before it.
    const before = [firstElement(xml, 'dimension', 0, data.start), firstElement(xml, 'sheetPr', 0, data.start)].find((e) => e);
    const at = before ? before.end : xml.indexOf('>', xml.indexOf('worksheet')) + 1;
    return xml.slice(0, at) + `<${p}sheetViews><${p}sheetView workbookViewId="0">${pane}</${p}sheetView></${p}sheetViews>` + xml.slice(at);
  }
  // The old pane goes, with selections in panes that may no longer exist; the top-left pane's selection stays.
  let inner = view.inner ?? '';
  for (const name of ['pane', 'selection']) {
    for (const e of elements(inner, name).reverse()) {
      if (name === 'pane' || /\spane="/.test(e.attrs)) inner = inner.slice(0, e.start) + inner.slice(e.end);
    }
  }
  const start = views.start + 1 + views.qname.length + views.attrs.length + 1 + view.start;
  const next = `<${view.qname}${view.attrs}>${pane}${inner}</${view.qname}>`;
  return xml.slice(0, start) + next + xml.slice(start + (view.end - view.start));
}

function writeMerges(xml: string, p: string, merges: Rect[]): string {
  const data = firstElement(xml, 'sheetData');
  if (!data) return xml;
  const el = merges.length ? `<${p}mergeCells count="${merges.length}">${merges.map((m) => `<${p}mergeCell ref="${rangeToA1(m)}"/>`).join('')}</${p}mergeCells>` : '';
  const existing = firstElement(xml, 'mergeCells', data.end);
  if (existing) return xml.slice(0, existing.start) + el + xml.slice(existing.end);
  if (!el) return xml;
  // After the last of the elements that must come before it.
  let at = data.end;
  for (const name of ['sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews']) {
    const e = firstElement(xml, name, data.end);
    if (e && e.end > at) at = e.end;
  }
  return xml.slice(0, at) + el + xml.slice(at);
}
