import { expect, test } from '@playwright/test';
import { goTo, openFile, ordersWorkbook, shown } from './helpers';

/** Glide's grid starts editing on typing only for the printable characters of ASCII. */
test.skip(process.env.E2E_GRID !== 'ruled', "Glide's grid doesn't start editing on other characters");

test.beforeEach(async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
});

test('typing accented letters into a cell edits it, the first letter kept', async ({ page }) => {
  await goTo(page, 'A5');
  await page.keyboard.type('Ärger über Öl', { delay: 20 });
  await page.keyboard.press('Enter');
  await expect(shown(page, 'A5')).toHaveText('Ärger über Öl');
});

test("an input method's composition in a cell edits it, with all of what was composed", async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'The composition is made through the DevTools protocol');
  await goTo(page, 'A5');
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Input.imeSetComposition', { text: 'に', selectionStart: 1, selectionEnd: 1 });
  await cdp.send('Input.imeSetComposition', { text: 'にほん', selectionStart: 3, selectionEnd: 3 });
  await cdp.send('Input.insertText', { text: '日本' });
  await expect(page.getByRole('textbox', { name: 'Contents of A5' })).toHaveValue('日本');
  await page.keyboard.press('Enter');
  await expect(shown(page, 'A5')).toHaveText('日本');
});
