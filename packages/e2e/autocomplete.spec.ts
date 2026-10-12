import { expect, test } from '@playwright/test';
import { goTo, openFile, ordersWorkbook, shown } from './helpers';

test.beforeEach(async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
});

test("finishes text from the cells above as it's typed, in their case, as Excel's AutoComplete does", async ({ page }) => {
  const entry = page.locator('.wbe-cell-entry-input');
  const selected = () => entry.evaluate((el: HTMLTextAreaElement) => [el.selectionStart, el.selectionEnd]);
  await goTo(page, 'A4');
  await page.keyboard.type('b');
  // The rest of "Bolts" is added and selected, so typing on replaces it.
  await expect(entry).toHaveValue('bolts');
  await expect.poll(selected).toEqual([1, 5]);
  await page.keyboard.type('O');
  await expect(entry).toHaveValue('bOlts');
  await expect.poll(selected).toEqual([2, 5]);
  await page.keyboard.press('Enter');
  await expect(shown(page, 'A4')).toHaveText('Bolts');

  // Backspace takes away what was added.
  await goTo(page, 'A5');
  await page.keyboard.type('n');
  await expect(entry).toHaveValue('nuts');
  // Held a moment, as a person's key would be: an Enter straight after a Backspace can be lost under load.
  await page.keyboard.press('Backspace', { delay: 50 });
  await expect(entry).toHaveValue('n');
  await page.keyboard.press('Enter');
  await expect(shown(page, 'A5')).toHaveText('n');
});

test('finishes text typed into the formula bar too', async ({ page }) => {
  await goTo(page, 'A4');
  const bar = page.getByRole('combobox', { name: 'Cell contents' });
  await bar.click();
  await bar.pressSequentially('nu');
  await expect(bar).toHaveValue('nuts');
  await bar.press('Enter');
  await expect(shown(page, 'A4')).toHaveText('Nuts');
});
