import { parseStructRef, stringifyTokens, tokenize, type Token } from '@borgar/fx';
import { colToLetters, inRect } from '@sitepen/formula-engine/refs';
import type { Rect } from './types';

/** Where an Excel table sits, as structured references need it. */
export interface TablePlace {
  name: string;
  sheet: number;
  sheetName: string;
  rect: Rect;
  headerRows: number;
  totalsRows: number;
  columns: string[];
}

/** A sheet name as a reference writes it: quoted unless it is a plain word. */
function sheetPrefix(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? `${name}!` : `'${name.replace(/'/g, "''")}'!`;
}

/** The rows a structured reference's sections cover, or undefined when they aren't there. */
function rowsOf(t: TablePlace, sections: string[], row: number): [number, number] | undefined {
  const top = t.rect.r1 + t.headerRows;
  const bottom = t.rect.r2 - t.totalsRows;
  const want = new Set(sections.map((s) => s.toLowerCase()));
  if (want.has('this row')) return row >= top && row <= bottom ? [row, row] : undefined;
  if (want.has('all')) return [t.rect.r1, t.rect.r2];
  const headers = want.has('headers');
  const totals = want.has('totals');
  const data = want.has('data') || (!headers && !totals);
  if ((headers && !t.headerRows) || (totals && !t.totalsRows)) return undefined;
  const from = headers ? t.rect.r1 : data ? top : bottom + 1;
  const to = totals ? t.rect.r2 : data ? bottom : t.rect.r1 + t.headerRows - 1;
  return from <= to ? [from, to] : undefined;
}

/**
 * The formula with each table reference (Orders[Total], [@Qty],
 * Orders[[#Totals],[Qty]]) written as the cells it stands for, for a
 * formula in `at`, so that the calculation engine, which knows no tables,
 * can calculate it. Undefined when a reference names a table or column that
 * isn't there, a row outside the table, or a table in another workbook.
 */
export function resolveTableRefs(formula: string, at: { sheet: number; row: number; col: number }, tables: TablePlace[]): string | undefined {
  if (formula.indexOf('[') === -1) return formula;
  let tokens: Token[];
  try {
    tokens = tokenize(formula, { mergeRefs: true, allowTernary: false });
  } catch {
    return undefined;
  }
  if (!tokens.some((t) => t.type === 'structured')) return formula;
  const out: Token[] = [];
  for (const t of tokens) {
    if (t.type !== 'structured') {
      out.push(t);
      continue;
    }
    const ref = parseStructRef(t.value);
    if (!ref || ref.context?.length) return undefined;
    const table = ref.table
      ? tables.find((x) => x.name.toLowerCase() === ref.table!.toLowerCase())
      : tables.find((x) => x.sheet === at.sheet && inRect(x.rect, at.row, at.col));
    if (!table) return undefined;
    const index = (name: string) => table.columns.findIndex((c) => c.toLowerCase() === name.toLowerCase());
    const cols = ref.columns ?? [];
    const first = cols.length ? index(cols[0]) : 0;
    const last = cols.length ? index(cols[cols.length - 1]) : table.columns.length - 1;
    if (first < 0 || last < 0) return undefined;
    const rows = rowsOf(table, ref.sections ?? [], at.row);
    if (!rows) return undefined;
    const c1 = table.rect.c1 + Math.min(first, last);
    const c2 = table.rect.c1 + Math.max(first, last);
    const cell = (r: number, c: number) => `$${colToLetters(c)}$${r + 1}`;
    const range = rows[0] === rows[1] && c1 === c2 ? cell(rows[0], c1) : `${cell(rows[0], c1)}:${cell(rows[1], c2)}`;
    out.push({ ...t, type: 'range', value: sheetPrefix(table.sheetName) + range });
  }
  return stringifyTokens(out);
}
