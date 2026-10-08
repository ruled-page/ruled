import AxeBuilder from '@axe-core/playwright';
import { readFile } from 'node:fs/promises';
import { expect, test, type Page } from '@playwright/test';
import ExcelJS from 'exceljs';
import { goTo, openFile, readWorkbook, shown } from './helpers';

/** "Data": A1 links to a web page, A2 to Other!B3, A3 to a script; "Other" is plain. */
async function linksWorkbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet('Data');
  data.getCell('A1').value = { text: 'Example', hyperlink: 'https://example.com/a', tooltip: 'Example site' };
  data.getCell('A2').value = { text: 'To Other', hyperlink: '#Other!B3' };
  data.getCell('A3').value = { text: 'Run me', hyperlink: 'javascript:alert(1)' };
  data.getCell('B1').value = 'Plain';
  wb.addWorksheet('Other').getCell('B3').value = 'Here';
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

/** Records the addresses the page asks to open in a new tab, opening none: no test goes to another site. */
async function catchOpens(page: Page): Promise<() => Promise<unknown[][]>> {
  await page.addInitScript(() => {
    const opened: unknown[][] = [];
    (window as unknown as { opened: unknown[][] }).opened = opened;
    window.open = (...args: unknown[]) => {
      opened.push(args);
      return null;
    };
  });
  return () => page.evaluate(() => (window as unknown as { opened: unknown[][] }).opened);
}

async function violations(page: Page): Promise<string[]> {
  await page.waitForTimeout(300);
  const { violations } = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa']).analyze();
  return violations.flatMap((v) => v.nodes.map((n) => `${v.id}: ${n.target.join(' ')}`));
}

let opened: () => Promise<unknown[][]>;
test.beforeEach(async ({ page }) => {
  opened = await catchOpens(page);
  await openFile(page, 'links.xlsx', await linksWorkbook());
});

test('follows a link to another sheet with the keyboard, and screen readers hear the link', async ({ page }) => {
  await expect(shown(page, 'A1')).toHaveText('Example, link to https://example.com/a');
  await goTo(page, 'A2');
  await expect(page.locator('.wbe-cell-link')).toHaveText('Other!B3');
  await page.keyboard.press('ControlOrMeta+Enter');
  await expect(page.getByRole('tab', { name: 'Other' })).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByRole('textbox', { name: 'Go to cell' })).toHaveValue('B3');
  // The grid has the focus, so typing goes on from there.
  await expect(page.locator('canvas[data-testid="data-grid-canvas"]:focus-within')).toHaveCount(1);
  expect(await opened()).toEqual([]);
});

test('refuses a link to a script, saying why, and opens nothing', async ({ page }) => {
  await goTo(page, 'A3');
  await page.keyboard.press('ControlOrMeta+Enter');
  await expect(page.getByText("This link can't be opened here: “javascript:alert(1)”", { exact: false })).toBeVisible();
  await expect(page.getByRole('alertdialog')).toHaveCount(0);
  expect(await opened()).toEqual([]);
});

test('asks before opening a web page, showing its address, then opens it in a new tab', async ({ page }) => {
  // Ctrl+click (⌘-click on a Mac) on the cell; rows are 20 pixels under a 24 pixel heading, after 40 pixels of row numbers.
  const box = (await page.locator('canvas[data-testid="data-grid-canvas"]').first().boundingBox())!;
  await page.keyboard.down('ControlOrMeta');
  await page.mouse.click(box.x + 40 + 30, box.y + 24 + 10);
  await page.keyboard.up('ControlOrMeta');
  const confirm = page.getByRole('alertdialog', { name: 'Open this link?' });
  await expect(confirm).toBeVisible();
  await confirm.getByRole('button', { name: 'Cancel' }).click();
  expect(await opened()).toEqual([]);
  // The formula bar shows the address, and follows it too.
  await page.getByRole('button', { name: 'Open link to https://example.com/a' }).click();
  await expect(confirm).toContainText('https://example.com/a');
  await confirm.getByRole('button', { name: 'Open link' }).click();
  expect(await opened()).toEqual([['https://example.com/a', '_blank', 'noopener,noreferrer']]);
});

test('adds a link from the dialog, with the keyboard, and saves it', async ({ page }) => {
  await goTo(page, 'B1');
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Insert link…' }).focus();
  await page.keyboard.press('Enter');
  const dialog = page.getByRole('dialog', { name: 'Insert link' });
  await expect(dialog.getByRole('textbox', { name: 'Address' })).toBeFocused();
  // An address that can't be opened is refused, saying why next to the field.
  await page.keyboard.type('javascript:alert(1)');
  await page.keyboard.press('Enter');
  const address = dialog.getByRole('textbox', { name: 'Address' });
  await expect(address).toHaveAttribute('aria-invalid', 'true');
  await expect(address).toHaveAccessibleDescription(/Links can only go to web pages/);
  await address.fill('www.example.com/new');
  await dialog.getByRole('textbox', { name: 'Text to display' }).fill('New page');
  await dialog.getByRole('textbox', { name: 'Screen tip' }).fill('Our new page');
  await page.keyboard.press('Enter');
  await expect(dialog).toBeHidden();
  await expect(shown(page, 'B1')).toHaveText('New page, link to https://www.example.com/new');
  await expect(page.getByRole('button', { name: '2 changes' })).toBeVisible();

  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: /Save as Excel/ }).click();
  const wb = await readWorkbook(await readFile(await (await download).path()));
  expect(wb.getWorksheet('Data')!.getCell('B1').value).toEqual({ text: 'New page', hyperlink: 'https://www.example.com/new' });
});

test('links to a place in the workbook from the dialog, and removes a link from the menu', async ({ page }) => {
  await goTo(page, 'B1');
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Insert link…' }).click();
  const dialog = page.getByRole('dialog', { name: 'Insert link' });
  await dialog.getByRole('radio', { name: 'Place in this workbook' }).check();
  await dialog.getByRole('combobox', { name: 'Sheet or name' }).selectOption('Other');
  await dialog.getByRole('textbox', { name: 'Cell' }).fill('B3');
  await dialog.getByRole('button', { name: 'Insert' }).click();
  await expect(page.locator('.wbe-cell-link')).toHaveText('Other!B3');

  await goTo(page, 'A1');
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Link', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Remove link' }).click();
  await expect(shown(page, 'A1')).toHaveText('Example');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(shown(page, 'A1')).toHaveText('Example, link to https://example.com/a');
});

for (const theme of ['Default', 'Dark']) {
  test(`the link dialog and the question before opening are accessible, ${theme} theme`, async ({ page }) => {
    await page.getByRole('combobox', { name: 'Theme' }).selectOption(theme);
    await goTo(page, 'A1');
    await page.keyboard.press('Shift+F10');
    await page.getByRole('menuitem', { name: 'Link', exact: true }).click();
    await page.getByRole('menuitem', { name: 'Edit link…' }).click();
    const dialog = page.getByRole('dialog', { name: 'Edit link' });
    await expect(dialog.getByRole('textbox', { name: 'Address' })).toHaveValue('https://example.com/a');
    expect(await violations(page)).toEqual([]);
    await page.keyboard.press('Escape');
    await expect(dialog).toBeHidden();
    await page.keyboard.press('ControlOrMeta+Enter');
    await expect(page.getByRole('alertdialog', { name: 'Open this link?' })).toBeVisible();
    expect(await violations(page)).toEqual([]);
  });
}
