import { expect, test, type Page } from '@playwright/test';
import { openSample } from './helpers';

/** The toolbar's buttons, lists and toggles that are cut off or scrolled out of it. */
async function cutOff(page: Page): Promise<string[]> {
  return page.evaluate(() => {
    const bar = document.querySelector('.wbe-toolbar')!.getBoundingClientRect();
    return [...document.querySelectorAll('.wbe-toolbar :is(button, select, label)')]
      .filter((el) => {
        const r = el.getBoundingClientRect();
        return r.width > 0 && (r.right > bar.right + 0.5 || r.left < bar.left - 0.5);
      })
      .map((el) => el.getAttribute('aria-label') ?? el.textContent ?? '');
  });
}

const label = (page: Page, text: string) => page.locator('.wbe-toolbar .wbe-label', { hasText: new RegExp(`^${text}$`) });

for (const [width, folded] of [
  [1920, false],
  [1440, true],
  [1100, true],
  [800, true],
] as const) {
  test(`the toolbar fits ${width} pixels, its labels folding into icons only when they must`, async ({ page }) => {
    await page.setViewportSize({ width, height: 800 });
    await openSample(page, 'Sales report');
    await expect(page.getByRole('button', { name: 'Save as Excel' })).toBeInViewport({ ratio: 1 });
    expect(await cutOff(page)).toEqual([]);
    // Folded labels stay for screen readers.
    await expect(page.getByRole('button', { name: 'Open', exact: true })).toBeVisible();
    if (folded) await expect(label(page, 'Open')).toHaveCSS('position', 'absolute');
    else await expect(label(page, 'Open')).toHaveCSS('position', 'static');
  });
}

test('the toolbar fits in German, whose labels are longer, and folds again when the language changes', async ({ page }) => {
  await page.setViewportSize({ width: 1440, height: 800 });
  await openSample(page, 'Sales report');
  await page.getByRole('combobox', { name: 'Language' }).selectOption('de');
  await expect(page.getByRole('button', { name: 'Als Excel speichern' })).toBeInViewport({ ratio: 1 });
  expect(await cutOff(page)).toEqual([]);
});
