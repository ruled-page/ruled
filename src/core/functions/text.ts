import { ArraySize, CellError, EmptyValue, ErrorType, FunctionArgumentType, SimpleRangeValue, type ImplementedFunctions } from 'hyperformula';
import {
  ExcelPlugin,
  array,
  arraySize,
  calcError,
  error,
  excelRound,
  grid,
  literalBoolean,
  literalNumber,
  numberText,
  omitted,
  plain,
  valueText,
  type Call,
  type Node,
  type State,
  type Value,
} from './common';

/** The longest text a cell can hold. */
const MAX_TEXT = 32767;

const matchesAt = (text: string, i: number, d: string, insensitive: boolean) => {
  const piece = text.slice(i, i + d.length);
  return insensitive ? piece.toLowerCase() === d.toLowerCase() : piece === d;
};

/** The texts an argument lists as delimiters: one text, or an array of them. */
function delimiters(v: unknown, keepEmpty: boolean): string[] | CellError {
  if (v === undefined) return [];
  const out: string[] = [];
  for (const x of grid(v).flat()) {
    if (x instanceof CellError) return x;
    const s = valueText(x);
    if (s || keepEmpty) out.push(s);
  }
  // Where two delimiters match at one place, the longer one is taken.
  return out.sort((a, b) => b.length - a.length);
}

function splitOn(text: string, delims: string[], insensitive: boolean): string[] {
  if (!delims.length) return [text];
  const parts: string[] = [];
  let start = 0;
  for (let i = 0; i < text.length; ) {
    const d = delims.find((x) => matchesAt(text, i, x, insensitive));
    if (d) {
      parts.push(text.slice(start, i));
      i += d.length;
      start = i;
    } else {
      i++;
    }
  }
  parts.push(text.slice(start));
  return parts;
}

/** TEXTSPLIT's table: rows split first, then each row into columns. */
function splitText(text: string, cols: string[], rows: string[], ignoreEmpty: boolean, insensitive: boolean): string[][] {
  const table = splitOn(text, rows, insensitive).map((line) => splitOn(line, cols, insensitive));
  return ignoreEmpty ? table.map((r) => r.filter(Boolean)).filter((r) => r.length) : table;
}

/**
 * Where the nth delimiter is, counting from the end when n is negative, as
 * [start, end]. With matchEnd the end of the text (or its start, counting
 * backwards) counts as one more delimiter. An empty delimiter matches at once.
 */
function findDelimiter(text: string, delims: string[], n: number, insensitive: boolean, matchEnd: boolean): [number, number] | undefined {
  if (delims.includes('')) return n > 0 ? [0, 0] : [text.length, text.length];
  let count = 0;
  if (n > 0) {
    for (let i = 0; i < text.length; ) {
      const d = delims.find((x) => matchesAt(text, i, x, insensitive));
      if (!d) {
        i++;
        continue;
      }
      if (++count === n) return [i, i + d.length];
      i += d.length;
    }
    return matchEnd && count === n - 1 ? [text.length, text.length] : undefined;
  }
  for (let j = text.length; j > 0; ) {
    const d = delims.find((x) => x.length <= j && matchesAt(text, j - x.length, x, insensitive));
    if (!d) {
      j--;
      continue;
    }
    if (++count === -n) return [j - d.length, j];
    j -= d.length;
  }
  return matchEnd && count === -n - 1 ? [0, 0] : undefined;
}

/** The text an argument is written as, when it is a literal. */
function literalText(node: Node): string | undefined {
  while (node.type === 'PARENTHESES') node = node.expression!;
  if (node.type === 'STRING') return node.value as string;
  if (node.type === 'FUNCTION_CALL' && !node.args?.length && (node.procedureName === 'TRUE' || node.procedureName === 'FALSE')) return node.procedureName;
  const n = literalNumber(node);
  return n === undefined ? undefined : numberText(n);
}

function literalDelimiters(node: Node | undefined): string[] | undefined {
  if (omitted(node)) return [];
  const items = node!.type === 'ARRAY' ? (node!.args as unknown as Node[][]).flat() : [node!];
  const texts = items.map(literalText);
  return texts.every((t): t is string => t !== undefined) ? texts.filter(Boolean).sort((a, b) => b.length - a.length) : undefined;
}

/** Thousands separators in the whole-number part of a formatted number. */
const groupThousands = (digits: string) => digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',');

const quote = (s: string) => `"${s.replace(/"/g, '""')}"`;

export class TextFunctions extends ExcelPlugin {
  static implementedFunctions: ImplementedFunctions = {
    CONCAT: { method: 'concat', parameters: [{ argumentType: FunctionArgumentType.SCALAR }], repeatLastArgs: 1, expandRanges: true },
    NUMBERVALUE: {
      method: 'numbervalue',
      parameters: [{ argumentType: FunctionArgumentType.SCALAR }, { argumentType: FunctionArgumentType.STRING, optionalArg: true }, { argumentType: FunctionArgumentType.STRING, optionalArg: true }],
    },
    TEXTSPLIT: {
      method: 'textsplit',
      sizeOfResultArrayMethod: 'textsplitSize',
      vectorizationForbidden: true,
      parameters: [
        { argumentType: FunctionArgumentType.ANY },
        { argumentType: FunctionArgumentType.ANY, optionalArg: true },
        { argumentType: FunctionArgumentType.ANY, optionalArg: true },
        { argumentType: FunctionArgumentType.BOOLEAN, defaultValue: false, emptyAsDefault: true },
        { argumentType: FunctionArgumentType.NUMBER, defaultValue: 0, emptyAsDefault: true },
        { argumentType: FunctionArgumentType.SCALAR, optionalArg: true },
      ],
    },
    TEXTBEFORE: { method: 'textbefore', parameters: textAroundParameters() },
    TEXTAFTER: { method: 'textafter', parameters: textAroundParameters() },
    VALUETOTEXT: {
      method: 'valuetotext',
      parameters: [{ argumentType: FunctionArgumentType.SCALAR }, { argumentType: FunctionArgumentType.NUMBER, defaultValue: 0, emptyAsDefault: true }],
    },
    ARRAYTOTEXT: {
      method: 'arraytotext',
      vectorizationForbidden: true,
      parameters: [{ argumentType: FunctionArgumentType.ANY }, { argumentType: FunctionArgumentType.NUMBER, defaultValue: 0, emptyAsDefault: true }],
    },
    FIXED: {
      method: 'fixed',
      parameters: [
        { argumentType: FunctionArgumentType.NUMBER },
        { argumentType: FunctionArgumentType.NUMBER, defaultValue: 2, emptyAsDefault: true },
        { argumentType: FunctionArgumentType.BOOLEAN, defaultValue: false, emptyAsDefault: true },
      ],
    },
  };

  concat(ast: Call, state: State) {
    return this.run(ast, state, 'CONCAT', (...values: Value[]) => {
      let out = '';
      for (const v of values) {
        if (v instanceof CellError) return v;
        out += valueText(v);
      }
      return out.length > MAX_TEXT ? error() : out;
    });
  }

  numbervalue(ast: Call, state: State) {
    return this.run(ast, state, 'NUMBERVALUE', (text: Value, decimal?: string, group?: string) => {
      if (text instanceof CellError || typeof text === 'number') return text;
      if (typeof text === 'boolean' || decimal === '' || group === '') return error();
      const dec = decimal?.[0] ?? '.';
      // A group separator left out is ",", or "." when "," is the decimal separator: "3,14".
      const grp = group?.[0] ?? (dec === ',' ? '.' : ',');
      if (dec === grp) return error();
      // Spaces are ignored anywhere, and each trailing % divides by 100.
      let s = text === EmptyValue ? '' : text.replace(/\s/g, '');
      let percents = 0;
      while (s.endsWith('%')) {
        s = s.slice(0, -1);
        percents++;
      }
      if (!s) return percents ? error() : 0;
      const at = s.indexOf(dec);
      if (at !== -1 && s.indexOf(dec, at + 1) !== -1) return error();
      const whole = at === -1 ? s : s.slice(0, at);
      const fraction = at === -1 ? '' : s.slice(at + 1);
      // Group separators only count before the decimal separator.
      if (fraction.includes(grp)) return error();
      const literal = whole.split(grp).join('') + (at === -1 ? '' : `.${fraction}`);
      if (!/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(literal)) return error();
      return Number(literal) / 100 ** percents;
    });
  }

  textsplit(ast: Call, state: State) {
    return this.run(ast, state, 'TEXTSPLIT', (text: unknown, col: unknown, row: unknown, ignoreEmpty: boolean, mode: number, pad?: Value) => {
      const cols = delimiters(col, false);
      if (cols instanceof CellError) return cols;
      const rows = delimiters(row, false);
      if (rows instanceof CellError) return rows;
      const m = Math.trunc(mode);
      if ((!cols.length && !rows.length) || (m !== 0 && m !== 1)) return error();
      const split = (t: string) => splitText(t, cols, rows, ignoreEmpty, m === 1);
      if (text instanceof SimpleRangeValue) {
        // A range of texts gives each one's first part, as Excel does.
        return array(grid(text).map((r) => r.map((v) => (v instanceof CellError ? v : (split(valueText(v))[0]?.[0] ?? calcError())))));
      }
      const parts = split(valueText(plain(text) as Exclude<Value, CellError>));
      if (!parts.length) return calcError();
      const width = Math.max(...parts.map((r) => r.length));
      const filler = pad === undefined ? error(ErrorType.NA) : pad;
      return array(parts.map((r) => [...r, ...Array<Value>(width - r.length).fill(filler)]));
    });
  }

  textsplitSize(ast: Call, state: State) {
    const [text, col, row, ignore, mode] = ast.args;
    const size = this.sizeOf(text, state);
    if (!size) return ArraySize.error();
    if (size.width * size.height > 1) return arraySize(size.width, size.height);
    const t = literalText(text);
    const cols = literalDelimiters(col);
    const rows = literalDelimiters(row);
    const ig = omitted(ignore) ? false : literalBoolean(ignore);
    const m = omitted(mode) ? 0 : literalNumber(mode);
    // Arguments that will give an error give a single cell.
    if (t === undefined || !cols || !rows || ig === undefined || m === undefined || (!cols.length && !rows.length) || (Math.trunc(m) !== 0 && Math.trunc(m) !== 1)) return ArraySize.error();
    const parts = splitText(t, cols, rows, ig, Math.trunc(m) === 1);
    return parts.length ? arraySize(Math.max(...parts.map((r) => r.length)), parts.length) : ArraySize.error();
  }

  textbefore(ast: Call, state: State) {
    return this.run(ast, state, 'TEXTBEFORE', this.textAround(false));
  }

  textafter(ast: Call, state: State) {
    return this.run(ast, state, 'TEXTAFTER', this.textAround(true));
  }

  private textAround(after: boolean) {
    return (text: Value, delimiter: unknown, instance: number, mode: number, matchEnd: number, notFound?: Value) => {
      if (text instanceof CellError) return text;
      const t = valueText(text);
      const delims = delimiters(delimiter, true);
      if (delims instanceof CellError) return delims;
      const n = Math.trunc(instance);
      const m = Math.trunc(mode);
      const end = Math.trunc(matchEnd);
      if (n === 0 || Math.abs(n) > t.length || (m !== 0 && m !== 1) || (end !== 0 && end !== 1)) return error();
      const at = findDelimiter(t, delims, n, m === 1, end === 1);
      if (!at) return notFound === undefined ? error(ErrorType.NA) : notFound;
      return after ? t.slice(at[1]) : t.slice(0, at[0]);
    };
  }

  valuetotext(ast: Call, state: State) {
    return this.run(ast, state, 'VALUETOTEXT', (v: Value, format: number) => {
      const f = Math.trunc(format);
      if (f !== 0 && f !== 1) return error();
      return this.asText(v, f === 1);
    });
  }

  arraytotext(ast: Call, state: State) {
    return this.run(ast, state, 'ARRAYTOTEXT', (v: unknown, format: number) => {
      const f = Math.trunc(format);
      if (f !== 0 && f !== 1) return error();
      const rows = grid(v);
      if (f === 0) return rows.flat().map((x) => this.asText(x, false)).join(', ');
      return `{${rows.map((r) => r.map((x) => this.asText(x, true)).join(',')).join(';')}}`;
    });
  }

  /** A value as text, concisely or (strict) as it would be typed in a formula, with text in quotes. */
  private asText(v: Value, strict: boolean): string {
    if (v instanceof CellError) return this.errorText(v);
    return strict && typeof v === 'string' ? quote(v) : valueText(v);
  }

  fixed(ast: Call, state: State) {
    return this.run(ast, state, 'FIXED', (n: number, decimals: number, noCommas: boolean) => {
      const d = Math.trunc(decimals);
      if (d > 127) return error();
      const r = excelRound(n, d, 'half');
      const places = Math.max(0, d);
      let [whole, fraction = ''] = Math.abs(r).toFixed(Math.min(places, 100)).split('.');
      fraction = fraction.padEnd(places, '0');
      if (!noCommas) whole = groupThousands(whole);
      return (r < 0 ? '-' : '') + whole + (places ? `.${fraction}` : '');
    });
  }
}

function textAroundParameters() {
  return [
    { argumentType: FunctionArgumentType.SCALAR },
    { argumentType: FunctionArgumentType.ANY },
    { argumentType: FunctionArgumentType.NUMBER, defaultValue: 1, emptyAsDefault: true },
    { argumentType: FunctionArgumentType.NUMBER, defaultValue: 0, emptyAsDefault: true },
    { argumentType: FunctionArgumentType.NUMBER, defaultValue: 0, emptyAsDefault: true },
    { argumentType: FunctionArgumentType.SCALAR, optionalArg: true },
  ];
}
