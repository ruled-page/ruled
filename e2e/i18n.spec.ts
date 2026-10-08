import AxeBuilder from '@axe-core/playwright';
import { expect, test, type Page } from '@playwright/test';
import { de as demoDe } from '../src/demo/locales/de';
import { de } from '../src/i18n';
import { openSample, setCell } from './helpers';

/** A German message as the editor shows it, with its count filled in. */
const german = (key: keyof typeof de, count?: number) => de[key].replace('{{count, number}}', String(count));

async function useGerman(page: Page): Promise<void> {
  await page.getByRole('combobox', { name: 'Language' }).selectOption({ label: 'Deutsch' });
}

test('switches the demo and the editor to German, and remembers it', async ({ page }) => {
  await page.goto('/');
  await useGerman(page);
  await expect(page.locator('html')).toHaveAttribute('lang', 'de');
  await expect(page.locator('.wbe').first()).toHaveAttribute('lang', 'de');
  await expect(page.getByRole('heading', { name: demoDe['welcome.heading'] })).toBeVisible();
  await expect(page.getByRole('button', { name: de['toolbar.shortcuts'] })).toBeVisible();

  await page.reload();
  await expect(page.getByRole('combobox', { name: demoDe['picker.language'] })).toHaveValue('de');
  await expect(page.getByRole('button', { name: de['toolbar.shortcuts'] })).toBeVisible();
});

test('translates what is already on screen when the language changes', async ({ page }) => {
  await openSample(page, 'Price list');
  const panel = page.getByRole('tabpanel');
  await expect(panel.getByText('Required.')).toBeVisible();
  await setCell(page, 'D3', 'twelve');
  await expect(panel.getByText('Expected a number.')).toBeVisible();

  await useGerman(page);
  // Problems found before the switch, the schema's own text, and the editor's labels.
  await expect(panel.getByText(de['schema.check.required'])).toBeVisible();
  await expect(panel.getByText(de['schema.check.expectedNumber'])).toBeVisible();
  await expect(page.getByRole('tab', { name: de['panel.checks'] })).toHaveAttribute('aria-selected', 'true');
  await expect(page.locator('.wbe-chip-status')).toContainText(german('schemaChip.errors_other', 6));
  await expect(page.getByRole('button', { name: german('toolbar.changes_one', 1) })).toBeVisible();
  // What screen readers read is German too.
  await expect(page.getByRole('combobox', { name: de['formulaBar.goTo'] })).toBeVisible();

  await page.getByRole('combobox', { name: demoDe['picker.language'] }).selectOption({ label: 'English' });
  await expect(panel.getByText('Required.')).toBeVisible();
});

test('the German editor meets the same accessibility checks', async ({ page }) => {
  await openSample(page, 'Price list');
  await useGerman(page);
  await page.waitForTimeout(300);
  const { violations } = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa']).analyze();
  expect(violations.map((v) => `${v.id}: ${v.nodes.map((n) => n.target.join(' ')).join(', ')}`)).toEqual([]);
});
