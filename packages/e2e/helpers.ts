import { expect, type Locator, type Page } from '@playwright/test';
import ExcelJS from 'exceljs';
import { parseA1 } from '@ruled/formula/refs';

/** Where the test servers are: the app's port, then the library's, the Angular example's and the host app's (see playwright.config.ts). */
export const serverUrl = (which: 'app' | 'lib' | 'angular' | 'hostApp' | 'portal') =>
  `http://localhost:${Number(process.env.E2E_PORT ?? 4173) + { app: 0, lib: 1, angular: 2, hostApp: 3, portal: 4 }[which]}`;

export const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/** Opens one of the demo's samples by its title. */
export async function openSample(page: Page, title: string): Promise<void> {
  await page.goto('/');
  await page
    .locator('.samples-list button')
    .filter({ has: page.locator('strong', { hasText: new RegExp(`^${title}$`) }) })
    .click();
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
}

/** Opens a workbook through the editor's file picker, as a user choosing a file would. */
export async function openFile(page: Page, name: string, bytes: Uint8Array): Promise<void> {
  await page.goto('/');
  await page.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({ name, mimeType: XLSX, buffer: Buffer.from(bytes) });
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
}

/**
 * The key that moves the focus to the next control: Tab, or Option+Tab in
 * Safari on a Mac, where Tab passes over buttons unless that is set.
 */
export function tabKey(page: Page, back = false): string {
  const mac = process.platform === 'darwin' && page.context().browser()?.browserType().name() === 'webkit';
  return `${back ? 'Shift+' : ''}${mac ? 'Alt+' : ''}Tab`;
}

/** Selects a cell or range by typing it into the name box; the grid takes the focus. */
export async function goTo(page: Page, ref: string): Promise<void> {
  const box = page.getByRole('combobox', { name: 'Go to cell' });
  await box.click();
  await box.fill(ref);
  await box.press('Enter');
  await expect(page.locator('canvas[data-testid="data-grid-canvas"]:focus-within')).toHaveCount(1);
  await expect(box).toHaveValue(ref.replace(/^.*!/, '').toUpperCase());
}

/** Types into a cell through the formula bar. */
export async function setCell(page: Page, ref: string, text: string): Promise<void> {
  await goTo(page, ref);
  const input = page.getByRole('combobox', { name: 'Cell contents' });
  await input.fill(text);
  await input.press('Enter');
}

/**
 * A cell as the grid shows it, from the table the grid keeps for screen
 * readers (the grid itself is drawn on a canvas). Covers the visible part of
 * sheets without hidden columns or frozen rows.
 */
export function shown(page: Page, ref: string): Locator {
  const at = parseA1(ref);
  if (!at) throw new Error(`Not a cell: ${ref}`);
  // Column 0 is the row numbers.
  return page.locator(`[data-testid="glide-cell-${at.col + 1}-${at.row}"]`);
}

/** A small workbook with formulas across two sheets. */
export async function ordersWorkbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const lines = wb.addWorksheet('Lines');
  lines.addRow(['Item', 'Qty', 'Price', 'Total']);
  lines.addRow(['Bolts', 2, 3, { formula: 'B2*C2', result: 6 }]);
  lines.addRow(['Nuts', 4, 0.5, { formula: 'B3*C3', result: 2 }]);
  const summary = wb.addWorksheet('Summary');
  summary.addRow(['Grand total', { formula: 'SUM(Lines!D2:D3)', result: 8 }]);
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

/** Reads a saved workbook back, as another program would. */
export async function readWorkbook(bytes: Buffer): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes as unknown as ArrayBuffer);
  return wb;
}
