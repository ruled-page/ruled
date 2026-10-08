import { readFileSync } from 'node:fs';
import { expect, test, type Locator, type Page } from '@playwright/test';
import { strFromU8, unzipSync } from 'fflate';
import { goTo as goToCell, openFile } from './helpers';

/** Everything here is done without the mouse. */

const comments = () => new Uint8Array(readFileSync(new URL('../tests/corpus/comments.xlsx', import.meta.url)));
const grid = (page: Page) => page.locator('canvas[data-testid="data-grid-canvas"]:focus-within');
const focused = (el: Locator) => el.evaluate((node) => node === document.activeElement || node.contains(document.activeElement));

/** Presses a key until the element, or something in it, has the focus. */
async function reach(page: Page, el: Locator, key: string): Promise<void> {
  for (let i = 0; i < 30 && !(await focused(el)); i++) await page.keyboard.press(key);
  expect(await focused(el), `${key} reaches ${el}`).toBe(true);
}

/** Goes to a cell from the keyboard: into the editor with Tab, to the name box with F6, and the cell's reference. */
async function goTo(page: Page, ref: string): Promise<void> {
  await reach(page, page.locator('.wbe-root'), 'Tab');
  const box = page.getByRole('textbox', { name: 'Go to cell' });
  await reach(page, box, 'F6');
  await page.keyboard.type(ref);
  await page.keyboard.press('Enter');
  await expect(grid(page)).toHaveCount(1);
}

test('the Changes panel lists changed notes and threads, and undo takes them back', async ({ page }) => {
  await openFile(page, 'comments.xlsx', comments());
  await goToCell(page, 'Review!B2');
  await page.keyboard.press('Shift+F2');
  await page.getByRole('textbox', { name: 'Note' }).fill('Raised.');
  await page.getByRole('button', { name: 'Save note' }).click();
  await goToCell(page, 'C3');
  await page.keyboard.press('Shift+F2');
  await page.getByRole('button', { name: 'Resolve thread' }).click();
  await page.getByRole('button', { name: 'Close' }).click();
  await page.getByRole('button', { name: /2 changes/ }).click();
  const rows = page.locator('.wbe-changes-table tbody tr');
  await expect(rows).toHaveCount(2);
  await expect(rows.first()).toContainText('Agreed with finance.\nRevisit in January.');
  await expect(rows.first()).toContainText('Note edited');
  await expect(rows.last()).toContainText('Comment thread: resolved');
  await expect(page.locator('.wbe-report-lede')).toHaveText('2 cells differ from the file as it was opened.');
  await expect(page.locator('.wbe-history')).toContainText('Resolve comment thread');

  await page.getByRole('button', { name: 'Undo' }).click();
  await expect(rows).toHaveCount(1);
  await page.getByRole('button', { name: 'Undo' }).click();
  await expect(page.locator('.wbe-report-lede')).toHaveText('Nothing has changed since the file was opened.');
});

test.describe('notes and comments with only a keyboard', () => {
  test.beforeEach(async ({ page }) => {
    await openFile(page, 'comments.xlsx', comments());
  });

  test('Shift+F2 adds a note, which shows in the formula bar and is counted as a change', async ({ page }) => {
    await goTo(page, 'Plain!B2');
    await page.keyboard.press('Shift+F2');
    const popup = page.getByRole('dialog', { name: 'New note on B2' });
    await expect(popup).toBeVisible();
    await expect(popup.getByRole('textbox', { name: 'Note' })).toBeFocused();
    await page.keyboard.type('Counted twice.');
    await reach(page, popup.getByRole('button', { name: 'Save note' }), 'Tab');
    await page.keyboard.press('Enter');
    await expect(popup).toBeHidden();
    await expect(grid(page)).toHaveCount(1);
    await expect(page.locator('.wbe-cell-note')).toHaveText('Note — Counted twice.');
    await expect(page.getByRole('button', { name: /1 change/ })).toBeVisible();

    // Again to change it, with Ctrl+Enter from the text.
    await page.keyboard.press('Shift+F2');
    await expect(page.getByRole('dialog', { name: 'Note on B2' })).toBeVisible();
    await page.keyboard.type(' Fixed.');
    await page.keyboard.press('Control+Enter');
    await expect(page.locator('.wbe-cell-note')).toHaveText('Note — Counted twice. Fixed.');
  });

  test('Shift+F2 opens a thread to reply to and resolve, and the saved file has the reply', async ({ page }) => {
    await goTo(page, 'Review!C3');
    await page.keyboard.press('Shift+F2');
    const popup = page.getByRole('dialog', { name: 'Comments on C3' });
    await expect(popup.getByRole('listitem')).toHaveCount(2);
    await expect(popup.getByRole('textbox', { name: 'Reply' })).toBeFocused();
    await page.keyboard.type('Thanks, Bob.');
    await page.keyboard.press('Tab');
    await expect(popup.getByRole('button', { name: 'Reply', exact: true })).toBeFocused();
    await page.keyboard.press('Enter');
    await expect(popup.getByRole('listitem')).toHaveCount(3);
    const mine = popup.getByRole('listitem').last();
    await expect(mine.locator('.wbe-comment-who')).toHaveText('Anonymous');
    await expect(mine.locator('.wbe-comment-text')).toHaveText('Thanks, Bob.');
    // Only the user's own comments can be changed.
    await expect(mine.getByRole('button', { name: 'Edit' })).toBeVisible();
    await expect(popup.getByRole('listitem').first().getByRole('button')).toHaveCount(0);
    await expect(popup.getByRole('textbox', { name: 'Reply' })).toHaveValue('');

    await reach(page, popup.getByRole('button', { name: 'Resolve thread' }), 'Tab');
    await page.keyboard.press('Enter');
    await expect(popup.getByText('Resolved', { exact: true })).toBeVisible();
    await expect(popup.getByRole('textbox', { name: 'Reply' })).toHaveCount(0);
    await page.keyboard.press('Escape');
    await expect(popup).toBeHidden();
    await expect(grid(page)).toHaveCount(1);
    await expect(page.locator('.wbe-cell-note')).toHaveText('Comment — Ann Analyst: Does this include the Q3 adjustment? (2 replies)');

    const download = page.waitForEvent('download');
    await page.keyboard.press('Control+s');
    const files = unzipSync(new Uint8Array(readFileSync(await (await download).path())));
    const threads = strFromU8(files['xl/threadedComments/threadedComment1.xml']);
    expect(threads).toContain('<text>Thanks, Bob.</text>');
    expect(threads).toMatch(/id="\{2F4A6C8E-1B3D-4E5F-8A9B-0C1D2E3F4A5B\}" done="1">/);
    expect(strFromU8(files['xl/persons/person.xml'])).toContain('displayName="Anonymous"');
  });

  test("a protected sheet's notes can be read but not changed", async ({ page }) => {
    await goTo(page, 'Locked!A1');
    await page.keyboard.press('Shift+F2');
    const popup = page.getByRole('dialog', { name: 'Note on A1' });
    await expect(popup).toContainText('This sheet is protected in Excel, which keeps its notes and comments as they are.');
    await expect(popup).toContainText('Ask before changing.');
    await expect(popup.getByRole('textbox')).toHaveCount(0);
    await expect(popup.getByRole('button', { name: 'Close' })).toBeFocused();
    await page.keyboard.press('Enter');
    await expect(popup).toBeHidden();

    // A cell with nothing to read says why.
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('Shift+F2');
    await expect(page.getByRole('dialog')).toHaveCount(0);
    await expect(page.locator('.wbe-toast')).toContainText('keeps its notes and comments');
  });

  test('the cell menu starts a comment thread', async ({ page }) => {
    await goTo(page, 'Plain!A2');
    await page.keyboard.press('Shift+F10');
    await reach(page, page.getByRole('menuitem', { name: 'New comment' }), 'ArrowDown');
    await page.keyboard.press('Enter');
    const popup = page.getByRole('dialog', { name: 'New comment on A2' });
    await expect(popup.getByRole('textbox', { name: 'Comment' })).toBeFocused();
    await page.keyboard.type('Which bolts?');
    await page.keyboard.press('Control+Enter');
    await expect(popup).toBeHidden();
    await expect(page.locator('.wbe-cell-note')).toHaveText('Comment — Anonymous: Which bolts?');
  });
});
