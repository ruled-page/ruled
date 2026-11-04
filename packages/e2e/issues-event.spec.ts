import ExcelJS from 'exceljs';
import { expect, test, type Page } from '@playwright/test';
import { goTo, serverUrl, setCell } from './helpers';

/** <workbook-editor> fires issueschange with the handle's issues, and only when they change. */

type Issue = { sheet?: string; cell?: string; rule: string; severity: string; message: string };
type Editor = HTMLElement & { issues: readonly Issue[]; open(bytes: Uint8Array, name: string): Promise<boolean>; registerSchema(schema: unknown): void };

async function workbook(): Promise<number[]> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Lines');
  ws.addRow(['Item', 'Qty']);
  ws.addRow(['Bolt', 4]);
  ws.addRow(['Nut', 8]);
  return [...new Uint8Array(await wb.xlsx.writeBuffer())];
}

const schema = {
  id: 'lines',
  name: { en: 'Order lines' },
  sheets: [{ name: 'Lines', tables: [{ key: 'lines', at: 'A1', fields: [{ key: 'item', label: 'Item', type: 'text' }, { key: 'qty', label: 'Qty', type: 'integer', min: 0 }] }] }],
};

/** The issueschange events seen so far, each as its detail and whether it is the element's issues. */
const events = (page: Page) => page.evaluate(() => (window as unknown as { seen: { detail: Issue[]; same: boolean }[] }).seen);

test('issueschange follows the issues and not the selection', async ({ page }) => {
  await page.goto(`${serverUrl('lib')}/examples/html/`);
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  await page.evaluate(
    async ({ bytes, schema }) => {
      const el = document.getElementById('editor') as Editor;
      el.registerSchema(schema);
      await el.open(new Uint8Array(bytes), 'lines.xlsx');
      const w = window as unknown as { seen: { detail: Issue[]; same: boolean }[] };
      w.seen = [];
      el.addEventListener('issueschange', (e) => w.seen.push({ detail: [...(e as CustomEvent<Issue[]>).detail], same: (e as CustomEvent).detail === el.issues }));
    },
    { bytes: await workbook(), schema },
  );
  await expect(page.locator('.wbe-file-name')).toHaveText('lines.xlsx');

  await setCell(page, 'B2', '-1');
  await expect.poll(async () => (await events(page)).at(-1)?.detail.some((i) => i.sheet === 'Lines' && i.cell === 'B2' && i.severity === 'error')).toBe(true);
  expect((await events(page)).every((e) => e.same)).toBe(true);

  // Moving about fires no issueschange.
  const count = (await events(page)).length;
  await goTo(page, 'A3');
  await goTo(page, 'B3');
  expect((await events(page)).length).toBe(count);

  await setCell(page, 'B2', '5');
  await expect.poll(async () => (await events(page)).at(-1)?.detail.some((i) => i.cell === 'B2')).toBe(false);
});
