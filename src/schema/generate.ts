import { BUILT_IN, createI18n, translator, type Translator } from '../i18n';
import type { LocalizedMap, MessageKey, MessageParams } from '../i18n/message';
import { toA1 } from '@sitepen/formula-engine/refs';
import { VT, type DisplayStyle } from '../core/types';
import { isDateStyle, type CellRecord } from '../state/format';
import type { SheetView } from '../state/sheetView';
import { listOptions, listValidationAt } from '../state/validation';
import { cellText, type Pos } from './binding';
import type { CellSchema, FieldSchema, FieldType, SheetSchema, TableSchema, WorkbookSchema } from './types';

/** How far into a sheet to look for a table's header row. */
const SCAN_ROWS = 100;
const SCAN_COLS = 60;
/** At most this many distinct texts makes a column an enum. */
const ENUM_MAX = 8;

function isEmpty(rec: CellRecord | null): boolean {
  return !rec || (rec.t === VT.None && rec.f === undefined);
}

function toKey(label: string, used: Set<string>): string {
  const words = label
    .replace(/%/g, ' pct ')
    .replace(/[^A-Za-z0-9]+/g, ' ')
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  let key = words.map((w, i) => (i === 0 ? w.toLowerCase() : w[0].toUpperCase() + w.slice(1).toLowerCase())).join('') || 'field';
  if (/^\d/.test(key)) key = `f${key}`;
  let unique = key;
  for (let n = 2; used.has(unique); n++) unique = `${key}${n}`;
  used.add(unique);
  return unique;
}

let builtIn: [string, Translator][] | undefined;

/** A message in every built-in language, for text the generator writes into a schema. */
function inEveryLanguage(key: MessageKey, params: MessageParams): LocalizedMap {
  builtIn ??= Object.keys(BUILT_IN).map((language) => [language, translator(createI18n(language))]);
  return Object.fromEntries(builtIn.map(([language, t]) => [language, t(key, params)]));
}

/** "SKU-1001" → ^SKU-\d{4}$ when every value has the same shape. */
function inferPattern(texts: string[]): { pattern: string; hint: LocalizedMap } | undefined {
  if (texts.length < 3) return undefined;
  const shape = (s: string) =>
    s.replace(/[A-Z]+|[a-z]+|\d+|./g, (m) => (/^\d+$/.test(m) ? `\\d{${m.length}}` : /^[A-Za-z]+$/.test(m) ? m : m.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
  // Letters must be identical to count as the same shape; digits only need the same length.
  const first = shape(texts[0]);
  if (!/\\d/.test(first) || !texts.every((t) => shape(t) === first)) return undefined;
  return { pattern: `^${first}$`, hint: inEveryLanguage('schema.generate.patternHint', { example: texts[0] }) };
}

function decimalsOf(n: number): number {
  const s = String(parseFloat(n.toPrecision(15)));
  const dot = s.indexOf('.');
  return dot < 0 || s.includes('e') ? 0 : s.length - dot - 1;
}

/**
 * Guesses a field's definition from the cells under its label. `listValues`
 * are the choices of an Excel list validation on the column, if it has one.
 */
export function inferField(label: string, cells: (CellRecord | null)[], styles: DisplayStyle[], used: Set<string>, listValues?: readonly string[]): FieldSchema {
  const filled = cells.filter((c): c is CellRecord => !isEmpty(c));
  const field: FieldSchema = { key: toKey(label, used), label, type: 'text' };
  if (listValues?.length) {
    // The workbook already says what's allowed.
    field.type = 'enum';
    field.values = [...listValues];
    if (filled.length && filled.length === cells.length) field.required = true;
    return field;
  }
  if (!filled.length) return field;
  const formulas = filled.filter((c) => c.f !== undefined).length;
  if (formulas === filled.length) field.formula = 'require';
  if (filled.length === cells.length) field.required = true;

  const numbers = filled.filter((c) => c.t === VT.Number);
  const strings = filled.filter((c) => c.t === VT.String);
  const bools = filled.filter((c) => c.t === VT.Bool);
  let type: FieldType = 'text';
  if (numbers.length === filled.length) {
    const style = (c: CellRecord) => styles[c.style] ?? styles[0];
    const dates = numbers.filter((c) => isDateStyle(style(c))).length;
    const percents = numbers.filter((c) => /%/.test(style(c).numFmt)).length;
    if (dates > numbers.length / 2) type = 'date';
    else if (percents > numbers.length / 2) type = 'percent';
    else if (numbers.every((c) => Number.isInteger(c.n ?? 0))) type = 'integer';
    else type = 'number';
    if (type === 'number' && !formulas) field.decimals = Math.max(...numbers.map((c) => decimalsOf(c.n ?? 0)));
    if (type !== 'date' && numbers.every((c) => (c.n ?? 0) >= 0)) field.min = 0;
  } else if (bools.length === filled.length) {
    type = 'boolean';
  } else if (strings.length === filled.length) {
    const texts = strings.map((c) => c.s ?? '');
    const distinct = [...new Set(texts)];
    if (distinct.length <= ENUM_MAX && texts.length >= distinct.length * 2 && distinct.every((d) => d.length <= 40)) {
      type = 'enum';
      field.values = distinct.sort();
    } else {
      const p = inferPattern(texts);
      if (p) {
        field.pattern = p.pattern;
        field.patternHint = p.hint;
      }
      if (distinct.length === texts.length && texts.length >= 3 && (p || distinct.every((d) => d.length <= 24))) field.unique = true;
    }
  }
  if (type === 'integer' && numbers.length >= 3 && new Set(numbers.map((c) => c.n)).size === numbers.length && !formulas) {
    // Distinct whole numbers in the first column are usually identifiers.
    field.unique = true;
  }
  field.type = type;
  return field;
}

interface Candidate {
  at: Pos;
  width: number;
  records: number;
}

/** Finds the largest block that looks like a header row with data under it, outside `taken` cells. */
function findTable(view: SheetView, taken: (row: number, col: number) => boolean = () => false): Candidate | undefined {
  const maxRow = Math.min(view.info.maxRow, SCAN_ROWS);
  const maxCol = Math.min(view.info.maxCol, SCAN_COLS);
  // Headings are typed text: a formula's text result is data, not a heading.
  const header = (r: number, c: number) => {
    const rec = view.get(r, c);
    return rec?.t === VT.String && rec.f === undefined && !view.mergeAt(r, c) && !taken(r, c);
  };
  let best: Candidate | undefined;
  for (let r = 0; r <= maxRow; r++) {
    let c = 0;
    while (c <= maxCol) {
      if (!header(r, c)) {
        c++;
        continue;
      }
      let end = c;
      while (end + 1 <= maxCol && header(r, end + 1)) end++;
      if (end > c) {
        let records = 0;
        for (let rr = r + 1; rr <= view.info.maxRow + 1; rr++) {
          let any = false;
          let blocked = false;
          for (let cc = c; cc <= end && !any; cc++) {
            blocked ||= taken(rr, cc);
            any = !isEmpty(view.get(rr, cc));
          }
          if (!any || blocked) break;
          records++;
        }
        const score = (end - c + 1) * (records + 1);
        if (records > 0 && (!best || score > best.width * (best.records + 1))) best = { at: { row: r, col: c }, width: end - c + 1, records };
      }
      c = end + 1;
    }
  }
  return best;
}

/** Every table on a sheet: the best block first, then smaller ones that don't overlap it, top to bottom. */
function findTables(view: SheetView): Candidate[] {
  const found: Candidate[] = [];
  const taken = (row: number, col: number) => found.some((t) => row >= t.at.row && row <= t.at.row + t.records && col >= t.at.col && col < t.at.col + t.width);
  for (let k = 0; k < 8; k++) {
    const next = findTable(view, taken);
    // After the first, only blocks with a few rows of data count; two cells side by side is a label and its value.
    if (!next || (found.length && (next.records < 2 || next.width < 2))) break;
    found.push(next);
  }
  return found.sort((a, b) => a.at.row - b.at.row || a.at.col - b.at.col);
}

/**
 * Single cells outside the tables that read as "Label | value", like
 * "Report date | 2026-10-01". A text value counts when it is calculated or
 * its label ends with a colon; otherwise two texts side by side are more
 * likely headings.
 */
function findCells(view: SheetView, tables: Candidate[], styles: DisplayStyle[]): CellSchema[] {
  const inTable = (row: number, col: number) => tables.some((t) => row >= t.at.row && row <= t.at.row + t.records && col >= t.at.col && col < t.at.col + t.width);
  const out: CellSchema[] = [];
  const used = new Set<string>();
  const maxRow = Math.min(view.info.maxRow, SCAN_ROWS);
  const maxCol = Math.min(view.info.maxCol, SCAN_COLS);
  for (let r = 0; r <= maxRow && out.length < 40; r++) {
    for (let c = 0; c < maxCol; c++) {
      const label = view.get(r, c);
      if (label?.t !== VT.String || label.f !== undefined || inTable(r, c) || view.mergeAt(r, c) || inTable(r, c + 1)) continue;
      const value = view.get(r, c + 1);
      if (isEmpty(value) || view.mergeAt(r, c + 1)) continue;
      const text = (label.s ?? '').trim();
      if (!text || (value!.t === VT.String && value!.f === undefined && !text.endsWith(':'))) continue;
      const { label: _label, unique: _unique, ...field } = inferField(text.replace(/:$/, ''), [value], styles, used);
      out.push({ ...field, ref: toA1(r, c + 1), label: { ref: toA1(r, c), text } });
      c++;
    }
  }
  return out;
}

/**
 * Builds a table definition for the block whose top-left cell is `at`,
 * read in the given orientation. Used by the generator and by the
 * designer when someone changes a table's shape.
 */
export function inferTable(
  view: SheetView,
  styles: DisplayStyle[],
  key: string,
  at: Pos,
  orientation: 'rows' | 'columns',
  fixed: boolean,
  /** All sheets, to read Excel list validations whose values live elsewhere. */
  views: SheetView[] = [view],
): TableSchema {
  const byRow = orientation === 'rows';
  const place = (recordPos: number, fieldPos: number): Pos =>
    byRow ? { row: at.row + recordPos, col: at.col + fieldPos } : { row: at.row + fieldPos, col: at.col + recordPos };
  const labelOffset = fixed ? 1 : 0;

  const labels: string[] = [];
  for (let f = labelOffset; f < 500; f++) {
    const p = place(0, f);
    const text = cellText(view, p.row, p.col);
    if (!text) break;
    labels.push(text);
  }
  let records = 0;
  for (let r = 1; r < 1_048_576; r++) {
    let any = fixed && !!cellText(view, place(r, 0).row, place(r, 0).col);
    for (let f = 0; f < labels.length && !any; f++) {
      const p = place(r, labelOffset + f);
      any = !isEmpty(view.get(p.row, p.col));
    }
    if (!any) break;
    records++;
  }
  const used = new Set<string>();
  const fields = labels.map((label, f) => {
    const cells = Array.from({ length: records }, (_, r) => {
      const p = place(r + 1, labelOffset + f);
      return view.get(p.row, p.col);
    });
    const first = place(1, labelOffset + f);
    const dv = listValidationAt(view.info, first.row, first.col);
    const listValues = dv ? listOptions(dv, views, styles, false) : undefined;
    return inferField(label, cells, styles, used, listValues);
  });
  const table: TableSchema = { key, at: toA1(at.row, at.col), orientation, fields };
  if (fixed) {
    table.records = {
      labels: Array.from({ length: records }, (_, r) => {
        const p = place(r + 1, 0);
        return cellText(view, p.row, p.col);
      }),
    };
    const corner = cellText(view, at.row, at.col);
    if (corner) table.corner = corner;
  } else {
    table.records = { min: 1 };
  }
  return table;
}

/** A first-draft schema describing the workbook's sheets as they are now. */
export function generateSchema(views: SheetView[], styles: DisplayStyle[], name: string): WorkbookSchema {
  const sheets: SheetSchema[] = [];
  for (const view of views) {
    if (view.info.kind !== 'worksheet' || view.info.state !== 'visible') continue;
    const found = findTables(view);
    const cells = findCells(view, found, styles);
    if (!found.length && !cells.length) continue;
    const used = new Set<string>();
    const tables = found.map((t, i) => tableFor(view, styles, views, t, toKey(found.length > 1 ? `${view.info.name} ${i + 1}` : view.info.name, used)));
    const sheet: SheetSchema = { name: view.info.name };
    if (tables.length) sheet.tables = tables;
    if (cells.length) sheet.cells = cells;
    sheets.push(sheet);
  }
  const id = name.toLowerCase().replace(/\.[a-z]+$/, '').replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'schema';
  return { id, name: name.replace(/\.[a-z]+$/i, ''), sheets };
}

/** A table definition for a block the generator found, read as a fixed grid when it looks like one. */
function tableFor(view: SheetView, styles: DisplayStyle[], views: SheetView[], found: Candidate, key: string): TableSchema {
  // A first column of distinct labels over all-numeric data reads best as a fixed grid.
  const firstCol: (CellRecord | null)[] = [];
  let numericRest = 0;
  let cellsRest = 0;
  for (let r = 1; r <= found.records; r++) {
    firstCol.push(view.get(found.at.row + r, found.at.col));
    for (let c = 1; c < found.width; c++) {
      const rec = view.get(found.at.row + r, found.at.col + c);
      if (isEmpty(rec)) continue;
      cellsRest++;
      if (rec!.t === VT.Number) numericRest++;
    }
  }
  const labelTexts = firstCol.map((c) => (c?.t === VT.String ? c.s ?? '' : ''));
  const fixed =
    found.width > 1 &&
    found.records >= 2 &&
    labelTexts.every(Boolean) &&
    new Set(labelTexts).size === labelTexts.length &&
    cellsRest > 0 &&
    numericRest / cellsRest >= 0.9;
  return inferTable(view, styles, key, found.at, 'rows', fixed, views);
}
