import { expect, test, type Page } from '@playwright/test';
import { grid as gridSurface, openFile, openSample, ordersWorkbook, shown } from './helpers';

const nameBox = (page: Page) => page.getByRole('combobox', { name: 'Go to cell' });

async function goTo(page: Page, text: string): Promise<void> {
  await nameBox(page).click();
  await nameBox(page).fill(text);
  await nameBox(page).press('Enter');
}

test("the sheet goes on past Z and row 500 as it is scrolled, to Excel's last cell", async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
  const grid = gridSurface(page).first();
  const box = (await grid.boundingBox())!;
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  // Scrolling right shows columns beyond the 26 an empty sheet starts with: a click there selects one.
  for (let i = 0; i < 8; i++) {
    await page.mouse.wheel(3000, 0);
    await page.waitForTimeout(50);
  }
  await page.mouse.click(box.x + box.width / 2, box.y + 60);
  await expect(nameBox(page)).toHaveValue(/^[B-Z][A-Z]\d+$/);
  await goTo(page, 'XFD1048576');
  await expect(nameBox(page)).toHaveValue('XFD1048576');
  await expect(shown(page, 'XFD1048576')).toBeAttached();
});

test('the name box lists names with where they go, and won’t open a hidden sheet', async ({ page }) => {
  await openSample(page, 'Sales report');
  await nameBox(page).click();
  await page.getByRole('button', { name: 'Show names' }).click();
  const option = page.getByRole('listbox', { name: 'Names' }).getByRole('option');
  await expect(option).toHaveText(['TaxRate Settings!$B$1']);
  await option.click();
  await expect(page.getByText('“Settings” is a hidden sheet. Unhide it to go there.').first()).toBeVisible();
  await expect(page.getByRole('tab', { name: 'Summary', selected: true })).toBeVisible();
});
