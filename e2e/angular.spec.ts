import { expect, test } from '@playwright/test';
import { setCell } from './helpers';

/**
 * The editor in an Angular app (examples/angular): Angular binds the
 * element's options, listens to its events and shows what a save hands it.
 * Runs with E2E_ANGULAR=1, after `npm ci` in examples/angular.
 */
test.use({ baseURL: 'http://localhost:4175' });

test('mounts in Angular, takes its options, and hands Angular the values on save', async ({ page }) => {
  await page.goto('/');
  await expect(page.getByRole('navigation', { name: 'Sheets' })).toBeVisible();
  // The schema came from Angular's [options] binding, set before the editor had loaded.
  const chip = page.getByRole('button', { name: /^Forecast \(from Angular\)/ });
  await expect(chip).toContainText('1 error');

  await setCell(page, 'E3', '30000');
  await expect(page.getByRole('status').filter({ hasText: 'change' })).toHaveText('1 change, not saved');
  await expect(chip).toContainText('All checks pass');

  // Angular's own button calls the element's save(); the save event reaches Angular.
  await page.getByRole('button', { name: 'Save from Angular' }).click();
  const received = page.getByRole('region', { name: 'What Angular received' });
  await expect(received.getByRole('row', { name: /2027-Q1/ })).toContainText('30,000');
  await expect(received).toContainText('Forecast!E3');
  await expect(page.getByRole('status').filter({ hasText: 'change' })).toHaveText('1 change, saved');

  // Angular switches the editor's language.
  await page.getByRole('combobox', { name: 'Language' }).selectOption('de');
  await expect(page.getByRole('button', { name: /^Prognose \(aus Angular\)/ })).toContainText('Alle Prüfungen bestanden');
});
