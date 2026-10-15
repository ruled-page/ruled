/**
 * Finds the functions Google Sheets couldn't export to Excel. Google writes
 * such a cell's formula as `IFERROR(__xludf.DUMMYFUNCTION("<its formula>"), <last value>)`,
 * so Excel, and this editor, show the last value.
 */

/** Sheets functions Excel doesn't have, by the name Google gives them. */
const GOOGLE_ONLY = new Set(
  (
    'ADD AI ARRAYFORMULA ARRAY_CONSTRAIN AVERAGE.WEIGHTED CONTINUE COUNTUNIQUE COUNTUNIQUEIFS DETECTLANGUAGE DIVIDE EPOCHTODATE EQ ' +
    'FLATTEN GOOGLEFINANCE GOOGLETRANSLATE GT GTE IMAGE IMPORTDATA IMPORTFEED IMPORTHTML IMPORTRANGE IMPORTXML ISBETWEEN ISDATE ' +
    'ISEMAIL ISURL JOIN LT LTE MARGINOFERROR MINUS MULTIPLY NE POW QUERY REGEXEXTRACT REGEXMATCH REGEXREPLACE SORTN SPARKLINE ' +
    'SPLIT TO_DATE TO_DOLLARS TO_PERCENT TO_PURE_NUMBER TO_TEXT UMINUS UNARY_PERCENT UPLUS'
  ).split(' '),
);

const ENTITIES: Record<string, string> = { quot: '"', amp: '&', lt: '<', gt: '>', apos: "'" };
const unescapeXml = (text: string) =>
  text.replace(/&(quot|amp|lt|gt|apos|#\d+|#x[0-9a-f]+);/gi, (_, e: string) =>
    e[0] === '#' ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : Number(e.slice(1))) : ENTITIES[e.toLowerCase()],
  );

const DUMMY = '__xludf.DUMMYFUNCTION("';

/** The formulas inside a formula's DUMMYFUNCTION calls, unquoted. */
function dummies(formula: string): string[] {
  const found: string[] = [];
  for (let at = formula.indexOf(DUMMY); at >= 0; at = formula.indexOf(DUMMY, at)) {
    let i = at + DUMMY.length;
    let inner = '';
    while (i < formula.length) {
      if (formula[i] === '"') {
        if (formula[i + 1] !== '"') break;
        i++;
      }
      inner += formula[i++];
    }
    found.push(inner);
    at = i;
  }
  return found;
}

/** The Google-only functions a Sheets formula calls; its first function when none is known to be. */
export function googleFunctions(formula: string): string[] {
  const code = formula.replace(/"(?:[^"]|"")*"/g, '""');
  const names = [...code.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)].map((m) => m[1].toUpperCase());
  const only = [...new Set(names.filter((n) => GOOGLE_ONLY.has(n)))];
  return only.length ? only : names.slice(0, 1);
}

/** Cells by the Google-only function they used, from a worksheet's XML. */
export function scanSheet(xml: string, counts: Map<string, number>): void {
  if (!xml.includes('__xludf.DUMMYFUNCTION')) return;
  for (const [, text] of xml.matchAll(/<f(?:\s[^>]*)?>([^<]*)<\/f>/g)) {
    const formula = unescapeXml(text);
    const names = new Set(dummies(formula).flatMap(googleFunctions));
    for (const name of names) counts.set(name, (counts.get(name) ?? 0) + 1);
  }
}

/** The entries of a zip whose names pass `want`, inflated. Enough for the .xlsx files Google exports: no ZIP64, no encryption. */
async function readZip(bytes: Uint8Array, want: (name: string) => boolean): Promise<Map<string, string>> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = bytes.byteLength - 22;
  while (end >= 0 && view.getUint32(end, true) !== 0x06054b50) end--;
  if (end < 0) throw new Error('Not a zip file.');
  const count = view.getUint16(end + 10, true);
  let at = view.getUint32(end + 16, true);
  const out = new Map<string, string>();
  const decoder = new TextDecoder();
  for (let n = 0; n < count && view.getUint32(at, true) === 0x02014b50; n++) {
    const method = view.getUint16(at + 10, true);
    const size = view.getUint32(at + 20, true);
    const nameLength = view.getUint16(at + 28, true);
    const local = view.getUint32(at + 42, true);
    const name = decoder.decode(bytes.subarray(at + 46, at + 46 + nameLength));
    at += 46 + nameLength + view.getUint16(at + 30, true) + view.getUint16(at + 32, true);
    if (!want(name)) continue;
    const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
    const data = bytes.subarray(start, start + size);
    if (method === 0) out.set(name, decoder.decode(data));
    else if (method === 8) out.set(name, await new Response(new Blob([data as Uint8Array<ArrayBuffer>]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).text());
  }
  return out;
}

/** The Google-only functions an exported .xlsx used, with how many cells used each, most used first. */
export async function droppedFunctions(bytes: ArrayBuffer | Uint8Array): Promise<{ name: string; cells: number }[]> {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const counts = new Map<string, number>();
  for (const xml of (await readZip(data, (name) => /^xl\/worksheets\/[^/]+\.xml$/.test(name))).values()) scanSheet(xml, counts);
  return [...counts].map(([name, cells]) => ({ name, cells })).sort((a, b) => b.cells - a.cells || a.name.localeCompare(b.name));
}
