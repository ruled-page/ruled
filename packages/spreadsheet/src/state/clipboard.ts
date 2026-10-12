/**
 * Parses the text/plain flavour Excel and Google Sheets put on the
 * clipboard: tab-separated cells, one row per line. Cells containing a tab,
 * newline or quote are wrapped in quotes with inner quotes doubled.
 */
export function parseClipboardText(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          cell += '"';
          i++;
        } else {
          quoted = false;
        }
      } else {
        cell += ch;
      }
    } else if (ch === '"' && cell === '' && text.indexOf('"', i + 1) !== -1) {
      quoted = true;
    } else if (ch === '\t') {
      row.push(cell);
      cell = '';
    } else if (ch === '\r' || ch === '\n') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
    } else {
      cell += ch;
    }
  }
  if (cell !== '' || row.length) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
}

/** An attribute's value in an HTML tag, quoted or not. */
function htmlAttr(tag: string, name: string): string | undefined {
  const m = new RegExp(`\\s${name}(?:\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s>]+)))?`, 'i').exec(tag);
  return m ? (m[1] ?? m[2] ?? m[3] ?? '') : undefined;
}

/**
 * The exact numbers in the HTML Excel puts on the clipboard, cell for cell as
 * parseClipboardText lays out its text: the text holds what each cell shows
 * (1,234.57), the HTML its value as well (x:num="1234.5678"). Null for a cell
 * without one, and undefined for HTML that didn't come from Excel. Merged
 * cells take up their whole span, as they do in the text.
 */
export function clipboardNumbers(html: string): (number | null)[][] | undefined {
  const lower = html.toLowerCase();
  if (lower.indexOf('x:num') === -1) return undefined;
  const rows: (number | null)[][] = [];
  /** Rows still covered below, by column, by a cell spanning several rows. */
  const covered: number[] = [];
  let at = 0;
  for (;;) {
    const tr = lower.indexOf('<tr', at);
    if (tr === -1) break;
    const end = lower.indexOf('</tr', tr);
    const stop = end === -1 ? lower.length : end;
    const row: (number | null)[] = [];
    let col = 0;
    const skipCovered = () => {
      while (covered[col] > 0) {
        covered[col]--;
        row[col++] = null;
      }
    };
    let cell = lower.indexOf('<t', tr + 3);
    while (cell !== -1 && cell < stop) {
      const close = lower.indexOf('>', cell);
      if (close === -1) break;
      // <td> and <th>, not <thead> or <tbody>.
      if (/^<t[dh][\s>]/.test(lower.slice(cell, cell + 4))) {
        skipCovered();
        const tag = html.slice(cell, close);
        const span = Math.max(1, Number(htmlAttr(tag, 'colspan')) || 1);
        const down = Math.max(1, Number(htmlAttr(tag, 'rowspan')) || 1);
        const raw = htmlAttr(tag, 'x:num');
        const n = raw ? Number(raw) : NaN;
        for (let k = 0; k < span; k++) {
          row[col + k] = k === 0 && Number.isFinite(n) ? n : null;
          if (down > 1) covered[col + k] = down - 1;
        }
        col += span;
      }
      cell = lower.indexOf('<t', close);
    }
    // Columns after the row's last cell that a cell above still covers.
    for (; col < covered.length; col++) {
      if (covered[col] > 0) covered[col]--;
      row[col] = null;
    }
    rows.push(row);
    at = stop + 1;
  }
  return rows;
}

/** Rows as columns and columns as rows; short rows count as padded with empty text. */
export function transposeGrid(values: readonly (readonly string[])[]): string[][] {
  const width = Math.max(0, ...values.map((r) => r.length));
  return Array.from({ length: width }, (_, c) => values.map((r) => r[c] ?? ''));
}
