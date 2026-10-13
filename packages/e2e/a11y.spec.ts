import AxeBuilder from '@axe-core/playwright';
import ExcelJS from 'exceljs';
import { readFileSync } from 'node:fs';
import { expect, test, type Page } from '@playwright/test';
import { goTo, openFile, openSample, ordersWorkbook, setCell } from './helpers';

/** WCAG 2.2 A and AA problems on the page as it is now, one line each. */
async function violations(page: Page): Promise<string[]> {
  // Let transitions and animations finish so colours are measured as they end up, however slow the browser.
  await page.waitForTimeout(300);
  await page.evaluate(() => Promise.all(document.getAnimations().filter((a) => a.effect?.getTiming().iterations !== Infinity).map((a) => a.finished.catch(() => undefined))));
  const { violations } = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa']).analyze();
  return violations.flatMap((v) => v.nodes.map((n) => `${v.id} (${v.impact}): ${n.target.join(' ')}: ${n.failureSummary?.split('\n')[1]?.trim() ?? v.help}`));
}

async function useTheme(page: Page, name: string): Promise<void> {
  await page.getByRole('combobox', { name: 'Theme' }).selectOption(name);
}

for (const theme of ['Default', 'Indigo', 'Paper', 'Dark', 'High contrast']) {
  test.describe(`accessibility, ${theme} theme`, () => {
    // A scan of the whole page is slow, slowest in Firefox on CI's machines.
    test.describe.configure({ timeout: 90_000 });
    // The themes differ only in colour, whose contrast axe works out alike in every browser:
    // Chromium checks them all, Firefox and WebKit the default and the dark.
    test.skip(({ browserName }) => browserName !== 'chromium' && theme !== 'Default' && theme !== 'Dark', 'Chromium checks the other themes');
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
      await expect(page.getByRole('dialog', { name: /^Show rows where/ })).toBeVisible();
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

    test("a cell's dropdown with more values than it shows at once", async ({ page }) => {
      const wb = new ExcelJS.Workbook();
      const values = Array.from({ length: 30 }, (_, i) => `Item ${i + 1}`);
      wb.addWorksheet('Sheet1').getCell('A1').dataValidation = { type: 'list', allowBlank: true, formulae: [`"${values.join(',')}"`] };
      await openFile(page, 'list.xlsx', new Uint8Array(await wb.xlsx.writeBuffer()));
      await useTheme(page, theme);
      await goTo(page, 'A1');
      await page.keyboard.press('Enter');
      const list = page.getByRole('listbox', { name: 'Values' });
      await expect(list).toBeVisible();
      expect(await violations(page)).toEqual([]);
      // The keyboard reaches every value, keeping the highlighted one in view, and the wheel scrolls the list.
      await page.keyboard.press('ArrowUp');
      await expect(list.getByRole('option', { name: 'Item 30' })).toHaveAttribute('aria-selected', 'true');
      await expect(list.getByRole('option', { name: 'Item 30' })).toBeInViewport();
      await expect(list.getByRole('option', { name: 'Item 1', exact: true })).not.toBeInViewport();
      await list.hover();
      await page.mouse.wheel(0, -2000);
      await expect(list.getByRole('option', { name: 'Item 1', exact: true })).toBeInViewport();
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

    test('a note, a comment thread, and both read-only', async ({ page }) => {
      await openFile(page, 'comments.xlsx', new Uint8Array(readFileSync(new URL('../spreadsheet/tests/corpus/comments.xlsx', import.meta.url))));
      await useTheme(page, theme);
      for (const [ref, name] of [
        ['Review!C3', 'Comments on C3'],
        ['B2', 'Note on B2'],
        ['Plain!A1', 'New note on A1'],
        ['Locked!A1', 'Note on A1'],
      ]) {
        await goTo(page, ref);
        await page.keyboard.press('Shift+F2');
        await expect(page.getByRole('dialog', { name })).toBeVisible();
        expect(await violations(page)).toEqual([]);
        await page.keyboard.press('Escape');
      }
    });

    test('the name box’s list of names, and the name manager and name dialogs', async ({ page }) => {
      await openFile(page, 'orders.xlsx', await ordersWorkbook());
      await useTheme(page, theme);
      await goTo(page, 'D2:D3');
      const box = page.getByRole('combobox', { name: 'Go to cell' });
      await box.click();
      await box.fill('Totals');
      await box.press('Enter');
      await expect(box).toHaveValue('Totals');
      await box.click();
      await box.press('Alt+ArrowDown');
      await expect(page.getByRole('listbox', { name: 'Names' })).toBeVisible();
      expect(await violations(page)).toEqual([]);
      await box.press('Escape');
      await box.press('Escape');

      await goTo(page, 'B2');
      await page.keyboard.press('Shift+F10');
      await page.getByRole('menuitem', { name: 'Name manager…' }).click();
      await expect(page.getByRole('dialog', { name: 'Name manager' })).toBeVisible();
      expect(await violations(page)).toEqual([]);
      await page.getByRole('button', { name: 'New…' }).click();
      const dialog = page.getByRole('dialog', { name: 'New name' });
      await dialog.getByRole('textbox', { name: 'Name' }).fill('A1');
      await dialog.getByRole('textbox', { name: 'Refers to' }).fill('=Nowhere!A1');
      await dialog.getByRole('button', { name: 'OK' }).click();
      await expect(dialog.getByText('This workbook has no sheet called "Nowhere".')).toBeVisible();
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
