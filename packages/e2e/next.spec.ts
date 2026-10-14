import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { expect, test, type Page } from '@playwright/test';
import { ordersWorkbook, readWorkbook, serverUrl, setCell, shown } from './helpers';

/**
 * The editor in a Next.js App Router app, installed from the packed library
 * (scripts/next-app.mjs): a Server Component reads a workbook from disk and
 * renders <Workbook> with its bytes, and a Server Action receives saves.
 * Run with E2E_NEXT_APP=1.
 */
const workbooks = fileURLToPath(new URL('./next-app/workbooks/', import.meta.url));
const saved = join(workbooks, 'saved.xlsx');

// The tests share the app's files.
test.describe.configure({ mode: 'serial' });

test.beforeEach(async () => {
  mkdirSync(workbooks, { recursive: true });
  writeFileSync(join(workbooks, 'orders.xlsx'), await ordersWorkbook());
  rmSync(saved, { force: true });
});

/** Errors in the page as it runs, a difference between the server's render and the browser's among them. */
function errorsIn(page: Page): string[] {
  const errors: string[] = [];
  page.on('console', (m) => m.type() === 'error' && errors.push(m.text()));
  page.on('pageerror', (e) => errors.push(e.message));
  return errors;
}

test('a Server Component renders the editor with a workbook from the server, and a Server Action receives it saved', async ({ page, request }) => {
  // Rendered on the server: the editor's box, until the browser makes the editor.
  const first = await request.get(serverUrl('nextApp'));
  expect(first.status()).toBe(200);
  expect(await first.text()).toContain('aria-busy="true"');

  const errors = errorsIn(page);
  await page.goto(serverUrl('nextApp'));
  await expect(shown(page, 'D2')).toHaveText('6');
  await setCell(page, 'B2', '10');
  await expect(shown(page, 'D2')).toHaveText('30');
  await page.getByRole('button', { name: 'Save', exact: true }).click();
  await expect(page.getByText('Saved orders.xlsx.')).toBeVisible();
  const wb = await readWorkbook(readFileSync(saved));
  expect(wb.getWorksheet('Lines')!.getCell('B2').value).toBe(10);
  expect(errors).toEqual([]);
});

test("a Server Action's answer reaches the editor: a refusal is shown, and the changes kept", async ({ page }) => {
  await page.goto(serverUrl('nextApp'));
  await expect(shown(page, 'D2')).toHaveText('6');
  await setCell(page, 'A5', 'refuse');
  await page.getByRole('button', { name: 'Save', exact: true }).click();
  await expect(page.getByText('Not saved: The server refused this workbook.')).toBeVisible();
  expect(existsSync(saved)).toBe(false);
  await expect(shown(page, 'A5')).toHaveText('refuse');
});
