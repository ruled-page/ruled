import AxeBuilder from '@axe-core/playwright';
import { expect, test, type Page } from '@playwright/test';
import ExcelJS from 'exceljs';
import { de } from '@ruled/sheet/internal/i18n/index.ts';
import { goTo, openFile, shown, XLSX } from './helpers';

/** "Data": A1 = 1, B1 = 31 January 2026 as a date, E1 = "Low". */
async function fillWorkbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet('Data');
  data.getCell('A1').value = 1;
  data.getCell('B1').value = new Date(Date.UTC(2026, 0, 31));
  data.getCell('B1').numFmt = 'yyyy-mm-dd';
  data.getColumn(2).width = 14;
  data.getCell('E1').value = 'Low';
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

const grid = (page: Page) => page.locator('canvas[data-testid="data-grid-canvas"]:focus-within');

/** Opens the cell menu on the selection and picks an item, without the mouse. */
async function fromMenu(page: Page, item: string): Promise<void> {
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: item }).focus();
  await page.keyboard.press('Enter');
}

async function violations(page: Page): Promise<string[]> {
  await page.waitForTimeout(300);
  const { violations } = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa']).analyze();
  return violations.flatMap((v) => v.nodes.map((n) => `${v.id}: ${n.target.join(' ')}`));
}

test.describe('in English', () => {
  test.beforeEach(async ({ page }) => {
    await openFile(page, 'fill.xlsx', await fillWorkbook());
  });

  test('fills a series from the cell menu with the keyboard, undone in one step', async ({ page }) => {
    await goTo(page, 'A1:A5');
    await fromMenu(page, 'Fill series…');
    const dialog = page.getByRole('dialog', { name: 'Fill series' });
    await expect(dialog.getByRole('radio', { name: 'Columns' })).toBeFocused();
    await expect(dialog.getByRole('radio', { name: 'Columns' })).toBeChecked();
    await expect(dialog.getByRole('radio', { name: 'Month' })).toBeDisabled();
    expect(await violations(page)).toEqual([]);
    const step = dialog.getByRole('textbox', { name: 'Step value' });
    await step.fill('2');
    await step.press('Enter');
    await expect(dialog).toBeHidden();
    await expect(grid(page)).toHaveCount(1);
    await expect(shown(page, 'A5')).toHaveText('9');
    await page.keyboard.press('ControlOrMeta+z');
    await expect(shown(page, 'A2')).toHaveText('');
    await expect(shown(page, 'A5')).toHaveText('');
  });

  test('says what is wrong with a step, and runs a date series on from one cell to its stop date', async ({ page }) => {
    await goTo(page, 'B1');
    await fromMenu(page, 'Fill series…');
    const dialog = page.getByRole('dialog', { name: 'Fill series' });
    // A date starts a series of dates.
    await expect(dialog.getByRole('radio', { name: 'Date', exact: true })).toBeChecked();
    await dialog.getByRole('radio', { name: 'Month' }).check();
    const step = dialog.getByRole('textbox', { name: 'Step value' });
    await step.fill('1.5');
    await dialog.getByRole('button', { name: 'Fill', exact: true }).click();
    await expect(dialog.getByText('Weekdays, months and years go in whole steps: enter a whole number.')).toBeVisible();
    await expect(step).toBeFocused();
    await expect(step).toHaveAttribute('aria-invalid', 'true');
    await step.fill('1');
    await dialog.getByRole('button', { name: 'Fill', exact: true }).click();
    const stop = dialog.getByRole('textbox', { name: 'Stop value' });
    await expect(stop).toBeFocused();
    await expect(dialog.getByText('With one cell selected, enter a stop value', { exact: false })).toBeVisible();
    await stop.fill('4/15/2026');
    await stop.press('Enter');
    await expect(dialog).toBeHidden();
    await expect(shown(page, 'B3')).toHaveText('2026-03-31');
    await expect(shown(page, 'B4')).toHaveText('');
  });

  test('adds a custom list that AutoFill continues, and keeps it in the browser', async ({ page }) => {
    await goTo(page, 'E1:E4');
    await fromMenu(page, 'Fill series…');
    const series = page.getByRole('dialog', { name: 'Fill series' });
    const opener = series.getByRole('button', { name: 'Custom lists…' });
    await opener.focus();
    await page.keyboard.press('Enter');
    const lists = page.getByRole('dialog', { name: 'Custom lists' });
    await expect(lists.getByRole('listbox', { name: 'Lists' })).toBeFocused();
    await expect(lists.getByRole('option', { name: 'Sun, Mon, Tue, Wed, Thu, Fri, Sat' })).toBeAttached();
    expect(await violations(page)).toEqual([]);
    await lists.getByRole('textbox', { name: 'List entries' }).fill('Low\nMedium\nHigh');
    await lists.getByRole('button', { name: 'Add' }).click();
    await expect(lists.getByRole('status')).toHaveText('List added.');
    await expect(lists.getByRole('listbox', { name: 'Lists' })).toHaveValue('own:0');
    await page.keyboard.press('Escape');
    await expect(lists).toBeHidden();
    await expect(opener).toBeFocused();

    await series.getByRole('radio', { name: 'AutoFill' }).check();
    await expect(series.getByRole('textbox', { name: 'Step value' })).toBeDisabled();
    await series.getByRole('button', { name: 'Fill', exact: true }).click();
    await expect(shown(page, 'E2')).toHaveText('Medium');
    await expect(shown(page, 'E4')).toHaveText('Low');

    await openFile(page, 'fill.xlsx', await fillWorkbook());
    await goTo(page, 'E1:E4');
    await fromMenu(page, 'Fill series…');
    await page.getByRole('button', { name: 'Custom lists…' }).click();
    await expect(page.getByRole('option', { name: 'Low, Medium, High' })).toBeAttached();
  });
});

test.describe('in German', () => {
  test.use({ locale: 'de-DE' });

  test('reads steps and stop values as German writes them', async ({ page }) => {
    await page.goto('/');
    await page.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({ name: 'fill.xlsx', mimeType: XLSX, buffer: Buffer.from(await fillWorkbook()) });
    const box = page.getByRole('combobox', { name: de['formulaBar.goTo'] });
    await box.fill('A1:A5');
    await box.press('Enter');
    await expect(grid(page)).toHaveCount(1);
    await fromMenu(page, de['fill.menu.series']);
    const dialog = page.getByRole('dialog', { name: 'Reihe ausfüllen' });
    const step = dialog.getByRole('textbox', { name: 'Inkrement' });
    await step.fill('abc');
    await step.press('Enter');
    await expect(dialog.getByText('Das Inkrement muss eine Zahl sein, z. B. 1,5.')).toBeVisible();
    await step.fill('0,5');
    await dialog.getByRole('textbox', { name: 'Endwert' }).fill('2,5');
    await step.press('Enter');
    await expect(dialog).toBeHidden();
    await expect(shown(page, 'A3')).toHaveText('2');
    // Shown as German writes it, too.
    await expect(shown(page, 'A4')).toHaveText('2,5');
    await expect(shown(page, 'A5')).toHaveText('');
  });
});
