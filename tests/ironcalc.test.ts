import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { toA1 } from '../src/core/refs';
import { WorkbookSession } from '../src/core/session';
import { english } from '../src/i18n';

/**
 * IronCalc's calculation workbooks (tests/ironcalc, see its README), graded
 * against the values Excel saved in them. Disagreements are held to the list
 * in tests/ironcalc/known.json, which only shrinks.
 */

const dir = join(__dirname, 'ironcalc');
const knownFile = join(dir, 'known.json');

function workbooks(d = dir): string[] {
  return readdirSync(d).flatMap((f) => {
    const p = join(d, f);
    return statSync(p).isDirectory() ? workbooks(p) : f.endsWith('.xlsx') ? [p] : [];
  });
}

/** Whether Excel saved the workbook, so its values are Excel's answers. */
function savedByExcel(bytes: Uint8Array): boolean {
  const app = unzipSync(bytes, { filter: (f) => f.name === 'docProps/app.xml' })['docProps/app.xml'];
  return !!app && /<Application>[^<]*Excel[^<]*<\/Application>/.test(strFromU8(app));
}

/** The cells, as Sheet!A1, whose results differ from the values Excel saved. */
function disagreements(bytes: Uint8Array, name: string): string[] {
  const s = WorkbookSession.open(bytes, name, 'en-US', { recalc: 'on' });
  const report = s.ensureEngine();
  const out: string[] = [];
  for (const [key, reason] of report.cellReasons) {
    if (!english.text(reason).startsWith('This editor calculated')) continue;
    const [sheet, row, col] = key.split(':').map(Number);
    out.push(`${s.model.sheets[sheet].info.name}!${toA1(row, col)}`);
  }
  return out.sort();
}

describe("IronCalc's workbooks", () => {
  it('give the answers Excel saved in them, apart from the known disagreements', { timeout: 60_000 }, () => {
    const found: Record<string, string[]> = {};
    for (const file of workbooks().sort()) {
      const bytes = new Uint8Array(readFileSync(file));
      if (!savedByExcel(bytes)) continue;
      const name = relative(dir, file);
      const cells = disagreements(bytes, name);
      if (cells.length) found[name] = cells;
    }
    if (process.env.UPDATE_IRONCALC) writeFileSync(knownFile, `${JSON.stringify(found, null, 2)}\n`);
    const known = JSON.parse(readFileSync(knownFile, 'utf8')) as Record<string, string[]>;
    expect(found).toEqual(known);
  });
});
