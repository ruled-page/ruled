import AxeBuilder from '@axe-core/playwright';
import { expect, test, type Page } from '@playwright/test';
import { goTo, openFile, openSample, ordersWorkbook, setCell } from './helpers';

/** WCAG 2.1 A and AA problems on the page as it is now, one line each. */
async function violations(page: Page): Promise<string[]> {
  // Let transitions finish so colours are measured as they end up.
  await page.waitForTimeout(300);
  const { violations } = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa']).analyze();
  return violations.flatMap((v) => v.nodes.map((n) => `${v.id} (${v.impact}): ${n.target.join(' ')}: ${n.failureSummary?.split('\n')[1]?.trim() ?? v.help}`));
}

async function useTheme(page: Page, name: string): Promise<void> {
  await page.getByRole('combobox', { name: 'Theme' }).selectOption(name);
}

for (const theme of ['Default', 'Dark']) {
  test.describe(`accessibility, ${theme} theme`, () => {
    test('start screen', async ({ page }) => {
      await page.goto('/');
      await useTheme(page, theme);
      expect(await violations(page)).toEqual([]);
    });

    test('a workbook with a schema, its checks and changes', async ({ page }) => {
      await openSample(page, 'Price list');
      await useTheme(page, theme);
      expect(await violations(page)).toEqual([]);
      await setCell(page, 'B11', 'Cable glands');
      await page.getByRole('tab', { name: /Changes/ }).click();
      expect(await violations(page)).toEqual([]);
    });

    test('a plain workbook and the report on what it holds', async ({ page }) => {
      await openSample(page, 'Sales report');
      await useTheme(page, theme);
      expect(await violations(page)).toEqual([]);
      await page.getByRole('button', { name: /^What's in this file/ }).click();
      expect(await violations(page)).toEqual([]);
    });

    test('menus, popups and dialogs', async ({ page }) => {
      // Scans a dozen menus, popups and dialogs, one after another.
      test.slow();
      await openSample(page, 'Price list');
      await useTheme(page, theme);
      await goTo(page, 'B3');
      await page.keyboard.press('Shift+F10');
      await expect(page.getByRole('menu')).toBeVisible();
      expect(await violations(page)).toEqual([]);
      await page.getByRole('menuitem', { name: 'Filter by values…' }).click();
      await expect(page.getByRole('dialog', { name: 'Filter by values' })).toBeVisible();
      expect(await violations(page)).toEqual([]);
      await page.keyboard.press('Escape');

      await goTo(page, 'C3');
      await page.keyboard.press('Enter');
      await expect(page.getByRole('listbox')).toBeVisible();
      expect(await violations(page)).toEqual([]);
      await page.keyboard.press('Escape');

      await page.keyboard.press('Control+f');
      await expect(page.getByRole('search')).toBeVisible();
      expect(await violations(page)).toEqual([]);
      await page.keyboard.press('Escape');

      await page.getByRole('button', { name: 'Fill colour' }).click();
      await expect(page.getByRole('dialog', { name: 'Fill colour' })).toBeVisible();
      expect(await violations(page)).toEqual([]);
      await page.keyboard.press('Escape');

      await page.keyboard.press('Control+/');
      await expect(page.getByRole('dialog', { name: 'Keyboard shortcuts' })).toBeVisible();
      expect(await violations(page)).toEqual([]);
      await page.keyboard.press('Escape');

      await goTo(page, 'B3');
      await page.keyboard.press('Shift+F10');
      await page.getByRole('menuitem', { name: 'Columns' }).click();
      await expect(page.getByRole('menu', { name: 'Columns' })).toBeVisible();
      expect(await violations(page)).toEqual([]);
      await page.getByRole('menuitem', { name: 'Column width…' }).click();
      await expect(page.getByRole('dialog', { name: 'Width of column B' })).toBeVisible();
      expect(await violations(page)).toEqual([]);
      await page.keyboard.press('Escape');
    });

    test('a confirmation', async ({ page }) => {
      await openFile(page, 'orders.xlsx', await ordersWorkbook());
      await useTheme(page, theme);
      // Clearing formulas asks first.
      await goTo(page, 'D2:D3');
      await page.keyboard.press('Delete');
      await expect(page.getByRole('alertdialog')).toBeVisible();
      expect(await violations(page)).toEqual([]);
    });

    test('the list of functions and argument hints, in the formula bar and in a cell', async ({ page }) => {
      await openFile(page, 'orders.xlsx', await ordersWorkbook());
      await useTheme(page, theme);
      await goTo(page, 'E2');
      await page.getByRole('combobox', { name: 'Cell contents' }).click();
      await page.keyboard.type('=B2*C2+SU');
      await expect(page.getByRole('listbox', { name: 'Functions' })).toBeVisible();
      expect(await violations(page)).toEqual([]);
      await page.keyboard.type('M(B2,');
      await expect(page.getByRole('tooltip')).toBeVisible();
      expect(await violations(page)).toEqual([]);
      await page.keyboard.press('Escape');

      // The same, typed in the cell.
      await goTo(page, 'E3');
      await page.keyboard.type('=B2*C2+SU', { delay: 20 });
      await expect(page.getByRole('listbox', { name: 'Functions' })).toBeVisible();
      expect(await violations(page)).toEqual([]);
      await page.keyboard.type('M(B2,', { delay: 20 });
      await expect(page.getByRole('tooltip')).toBeVisible();
      expect(await violations(page)).toEqual([]);
    });

    test('the schema designer', async ({ page }) => {
      await openSample(page, 'Price list');
      await useTheme(page, theme);
      await page.getByRole('button', { name: 'Schemas' }).click();
      await expect(page.getByRole('dialog')).toBeVisible();
      expect(await violations(page)).toEqual([]);
    });
  });
}

test('dialogs close with Escape and give the focus back', async ({ page }) => {
  await openSample(page, 'Price list');
  const opener = page.getByRole('button', { name: 'Schemas' });
  await opener.click();
  const dialog = page.getByRole('dialog');
  await expect(dialog).toBeVisible();
  await expect(dialog.locator(':focus')).toHaveCount(1);
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(opener).toBeFocused();
});
