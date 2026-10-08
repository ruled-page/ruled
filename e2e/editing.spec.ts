import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { goTo, openFile, openSample, ordersWorkbook, readWorkbook, setCell, shown } from './helpers';

test.describe('editing a workbook', () => {
  test.beforeEach(async ({ page }) => {
    await openFile(page, 'orders.xlsx', await ordersWorkbook());
  });

  test('recalculates formulas on this sheet and others', async ({ page }) => {
    await expect(shown(page, 'D2')).toHaveText('6');
    await setCell(page, 'B2', '5');
    await expect(shown(page, 'D2')).toHaveText('15');

    await page.getByRole('navigation', { name: 'Sheets' }).getByRole('button', { name: 'Summary' }).click();
    await expect(page.getByRole('button', { name: 'Summary' })).toHaveAttribute('aria-current', 'true');
    await expect(shown(page, 'B1')).toHaveText('17');
  });

  test('types straight into the grid', async ({ page }) => {
    await goTo(page, 'C3');
    await page.keyboard.type('1.25', { delay: 30 });
    await page.keyboard.press('Enter');
    await expect(shown(page, 'D3')).toHaveText('5');
    // Enter moves down, as in Excel.
    await expect(page.getByRole('textbox', { name: 'Go to cell' })).toHaveValue('C4');
  });

  test('pastes from Excel with the exact values behind the rounded text', async ({ page, browserName }) => {
    test.skip(browserName === 'firefox', 'Firefox gives a paste event made by a script its own empty clipboard data.');
    await goTo(page, 'C2');
    // What Excel puts on the clipboard for a price shown as 0.33 but worth a third.
    await page.evaluate(() => {
      const data = new DataTransfer();
      data.setData('text/plain', '0.33\r\n');
      data.setData('text/html', '<table><tr><td x:num="0.333333333333333">0.33</td></tr></table>');
      document.activeElement!.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }));
    });
    // Two at a third each (0.66 had the rounded text been used), in a column wide enough for six digits.
    await expect(shown(page, 'D2')).toHaveText('0.666667');
    await expect(page.getByRole('textbox', { name: 'Cell contents' })).toHaveValue('0.333333333333333');
  });

  test('copies the text and an escaped HTML table with exact values', async ({ page }) => {
    await setCell(page, 'A2', '<b onmouseover="alert(1)">Bolts</b>');
    await setCell(page, 'C2', '=1/3');
    await goTo(page, 'A2:C2');
    const copied = await page.evaluate(() => {
      let html = '';
      let text = '';
      // Reads what the editor put on the event once it has handled it.
      window.addEventListener('copy', (e) => ((html = e.clipboardData!.getData('text/html')), (text = e.clipboardData!.getData('text/plain'))), { once: true });
      document.activeElement!.dispatchEvent(new ClipboardEvent('copy', { clipboardData: new DataTransfer(), bubbles: true, cancelable: true }));
      return { html, text };
    });
    expect(copied.html).not.toContain('<b ');
    expect(copied.html).toContain('&#60;b onmouseover=&#34;alert(1)&#34;&#62;Bolts&#60;/b&#62;');
    expect(copied.html).toContain('x:num="0.333333333333333"');
    expect(copied.text.split('\t')[1]).toBe('2');
  });

  test('undoes and redoes', async ({ page }) => {
    await setCell(page, 'B3', '10');
    await expect(shown(page, 'D3')).toHaveText('5');
    await page.getByRole('button', { name: 'Undo' }).click();
    await expect(shown(page, 'D3')).toHaveText('2');
    await page.getByRole('button', { name: 'Redo' }).click();
    await expect(shown(page, 'D3')).toHaveText('5');
  });

  test('saves an .xlsx with the edits, formulas and recalculated values', async ({ page }) => {
    await setCell(page, 'B2', '5');
    await setCell(page, 'A3', 'Washers');
    await expect(shown(page, 'D2')).toHaveText('15');

    const download = page.waitForEvent('download');
    await page.getByRole('button', { name: /Save as Excel/ }).click();
    const file = await download;
    expect(file.suggestedFilename()).toBe('orders.xlsx');

    const wb = await readWorkbook(await readFile(await file.path()));
    const lines = wb.getWorksheet('Lines')!;
    expect(lines.getCell('B2').value).toBe(5);
    expect(lines.getCell('A3').value).toBe('Washers');
    expect(lines.getCell('D2').value).toMatchObject({ formula: 'B2*C2', result: 15 });
    expect(wb.getWorksheet('Summary')!.getCell('B1').value).toMatchObject({ formula: 'SUM(Lines!D2:D3)', result: 17 });
  });

  test('goes to a cell on another sheet from the name box', async ({ page }) => {
    await goTo(page, 'Summary!B1');
    await expect(page.getByRole('button', { name: 'Summary' })).toHaveAttribute('aria-current', 'true');
    await expect(page.getByRole('textbox', { name: 'Cell contents' })).toHaveValue('=SUM(Lines!D2:D3)');
  });

  test('says when the name box holds something that is not a reference', async ({ page }) => {
    const box = page.getByRole('textbox', { name: 'Go to cell' });
    await box.click();
    await box.fill('Nowhere!A1');
    await box.press('Enter');
    await expect(page.getByText('"Nowhere!A1" isn\'t a cell or range on a sheet in this workbook.')).toBeVisible();
  });
});

test('refuses files that are not workbooks', async ({ page }) => {
  await page.goto('/');
  await page.locator('input[type="file"][accept*=".xlsx"]').setInputFiles({ name: 'notes.txt', mimeType: 'text/plain', buffer: Buffer.from('hello') });
  await expect(page.getByText('Only Excel workbooks (.xlsx or .xlsm) can be opened.')).toBeVisible();
});

test('opens the file picker from the keyboard', async ({ page }) => {
  await page.goto('/');
  const chooser = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: /Choose a workbook/ }).focus();
  await page.keyboard.press('Enter');
  expect((await chooser).isMultiple()).toBe(false);
});

test('switching theme restyles the editor and keeps the workbook open', async ({ page }) => {
  await openSample(page, 'Formatting');
  const root = page.locator('.wbe').first();
  await page.getByRole('combobox', { name: 'Theme' }).selectOption('Dark');
  await expect(root).toHaveCSS('color-scheme', 'dark');
  expect(await root.evaluate((el) => getComputedStyle(el).getPropertyValue('--wbe-grid-bg').trim())).toBe('#1b1f23');
  await expect(page.getByRole('button', { name: 'Borders & text' })).toBeVisible();
  await page.getByRole('combobox', { name: 'Theme' }).selectOption('Default');
  await expect(root).toHaveCSS('color-scheme', 'light');
});
