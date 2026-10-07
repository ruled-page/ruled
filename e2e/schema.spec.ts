import { expect, test } from '@playwright/test';
import { goTo, openSample, setCell, shown } from './helpers';

test.describe('a workbook checked by a schema', () => {
  test.beforeEach(async ({ page }) => {
    await openSample(page, 'Price list');
  });

  test('counts problems and clears them as they are fixed', async ({ page }) => {
    const chip = page.getByRole('button', { name: /^Price list/ });
    await expect(chip).toContainText('4 errors');

    await setCell(page, 'B11', 'Cable glands');
    await expect(chip).toContainText('3 errors');
    await setCell(page, 'C12', 'Software');
    await expect(chip).toContainText('2 errors');
  });

  test('keeps what was typed and flags it when it is the wrong type', async ({ page }) => {
    const chip = page.getByRole('button', { name: /^Price list/ });
    await expect(chip).toContainText('4 errors');
    await setCell(page, 'D3', 'twelve');
    // Screen readers hear the problem with the value; the grid marks it in red.
    await expect(shown(page, 'D3')).toHaveText('twelve, error: Expected a number.');
    // The price inc. tax formula that uses it fails too.
    await expect(chip).toContainText('6 errors');
    const panel = page.getByRole('tabpanel');
    await expect(panel.getByText('Expected a number.')).toBeVisible();
    await expect(panel.getByText('The formula returns #VALUE!.')).toBeVisible();
    await goTo(page, 'D3');
    await expect(page.locator('.cell-note')).toContainText('Expected a number.');
  });

  test('refuses a formula where the schema wants a typed value', async ({ page }) => {
    await setCell(page, 'D3', '=1+2');
    await expect(page.getByText('Formulas aren\'t allowed in “Unit price”. Type a value instead.')).toBeVisible();
    await expect(shown(page, 'D3')).toHaveText('7.25');
  });

  test('lists the problems in the checks panel and goes to each one', async ({ page }) => {
    // A workbook with a schema opens with its checks showing.
    const panel = page.getByRole('tabpanel');
    await expect(page.getByRole('tab', { name: 'Checks' })).toHaveAttribute('aria-selected', 'true');
    await panel.getByText('Required.').click();
    await expect(page.getByRole('textbox', { name: 'Go to cell' })).toHaveValue('B11');
  });

  test('moves between the panel tabs with the arrow keys', async ({ page }) => {
    await page.getByRole('tab', { name: 'Checks' }).focus();
    await page.keyboard.press('ArrowRight');
    await expect(page.getByRole('tab', { name: 'Changes' })).toBeFocused();
    await expect(page.getByRole('tab', { name: 'Changes' })).toHaveAttribute('aria-selected', 'true');
    await page.keyboard.press('End');
    await expect(page.getByRole('tab', { name: "What's in this file" })).toHaveAttribute('aria-selected', 'true');
  });
});
