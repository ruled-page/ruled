import { readFile } from 'node:fs/promises';
import AxeBuilder from '@axe-core/playwright';
import { expect, test, type Page } from '@playwright/test';
import { goTo, openFile, ordersWorkbook, readWorkbook, shown, tabKey } from './helpers';

/** Everything here, but the dragging, is done without the mouse. */

const tab = (page: Page, name: string) => page.getByRole('tab', { name, exact: true });
const tabNames = (page: Page) => page.getByRole('tab').allTextContents();

/** Moves down the open menu to an item and chooses it, as the arrow keys and Enter do. */
async function choose(page: Page, item: string): Promise<void> {
  for (let i = 0; i < 12; i++) {
    const focused = page.locator('[role="menuitem"]:focus');
    if ((await focused.textContent())?.trim() === item) break;
    await page.keyboard.press('ArrowDown');
  }
  await expect(page.locator('[role="menuitem"]:focus')).toHaveText(item);
  await page.keyboard.press('Enter');
}

async function saveWithKeys(page: Page) {
  const download = page.waitForEvent('download');
  await page.keyboard.press('Control+s');
  return readWorkbook(await readFile(await (await download).path()));
}

test.describe('sheets, with only a keyboard', () => {
  test.beforeEach(async ({ page }) => {
    await openFile(page, 'orders.xlsx', await ordersWorkbook());
    await goTo(page, 'A1');
    await page.keyboard.press('F6');
    await expect(tab(page, 'Lines')).toBeFocused();
  });

  test('adds a sheet after the one shown, renames it, moves it and deletes it, and saves the result', async ({ page }) => {
    await page.keyboard.press(tabKey(page));
    await expect(page.getByRole('button', { name: 'New sheet' })).toBeFocused();
    await page.keyboard.press('Enter');
    await expect(tab(page, 'Sheet1')).toBeFocused();
    await expect(tab(page, 'Sheet1')).toHaveAttribute('aria-selected', 'true');
    expect(await tabNames(page)).toEqual(['Lines', 'Sheet1', 'Summary']);

    // F2 renames it; Excel's rules are checked as the name is typed.
    await page.keyboard.press('F2');
    const dialog = page.getByRole('dialog', { name: 'Rename “Sheet1”' });
    const name = dialog.getByRole('textbox', { name: 'Sheet name' });
    await expect(name).toBeFocused();
    await page.keyboard.type('Summary');
    await expect(name).toHaveAttribute('aria-invalid', 'true');
    await expect(name).toHaveAccessibleDescription(/There's already a sheet called “Summary”/);
    await page.keyboard.press('Enter');
    await expect(dialog).toBeVisible();
    await page.keyboard.press('ControlOrMeta+a');
    await page.keyboard.type('Notes 2026');
    await expect(name).toHaveAttribute('aria-invalid', 'false');
    await page.keyboard.press('Enter');
    await expect(dialog).toBeHidden();
    await expect(tab(page, 'Notes 2026')).toBeFocused();

    // Alt+Shift+arrows move it; so does its menu.
    await page.keyboard.press('Alt+Shift+ArrowRight');
    await expect.poll(() => tabNames(page)).toEqual(['Lines', 'Summary', 'Notes 2026']);
    await expect(tab(page, 'Notes 2026')).toBeFocused();
    await page.keyboard.press('Shift+F10');
    await expect(page.getByRole('menu', { name: 'Sheet “Notes 2026”' })).toBeVisible();
    await choose(page, 'Move left');
    await expect.poll(() => tabNames(page)).toEqual(['Lines', 'Notes 2026', 'Summary']);
    await expect(tab(page, 'Notes 2026')).toBeFocused();

    // A sheet with nothing on it goes without asking.
    await page.keyboard.press('Shift+F10');
    await choose(page, 'Delete');
    await expect.poll(() => tabNames(page)).toEqual(['Lines', 'Summary']);
    await expect(tab(page, 'Summary')).toBeFocused();

    // One with cells asks first; formulas that used it show #REF!.
    await page.keyboard.press('ArrowLeft');
    await expect(tab(page, 'Lines')).toBeFocused();
    await page.keyboard.press('Shift+F10');
    await choose(page, 'Delete');
    const confirm = page.getByRole('alertdialog');
    await expect(confirm).toContainText('Delete the sheet “Lines”?');
    // It starts on Cancel, so a stray Enter deletes nothing.
    await expect(confirm.getByRole('button', { name: 'Cancel' })).toBeFocused();
    await page.keyboard.press(tabKey(page));
    await expect(confirm.getByRole('button', { name: 'Delete sheet' })).toBeFocused();
    await page.keyboard.press('Enter');
    await expect.poll(() => tabNames(page)).toEqual(['Summary']);
    await expect(tab(page, 'Summary')).toBeFocused();
    await expect(shown(page, 'B1')).toHaveText('#REF!');

    const wb = await saveWithKeys(page);
    expect(wb.worksheets.map((w) => w.name)).toEqual(['Summary']);
    expect(wb.getWorksheet('Summary')!.getCell('B1').value).toMatchObject({ formula: 'SUM(#REF!)' });

    // Undo brings it back.
    await page.keyboard.press('Control+z');
    await expect.poll(() => tabNames(page)).toEqual(['Lines', 'Summary']);
  });

  test('hides a sheet and unhides it from the note in the tab bar', async ({ page }) => {
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('Shift+F10');
    await choose(page, 'Hide');
    await expect.poll(() => tabNames(page)).toEqual(['Lines']);
    await expect(tab(page, 'Lines')).toBeFocused();
    // The last visible sheet stays.
    await page.keyboard.press('Shift+F10');
    await choose(page, 'Hide');
    await expect(page.getByRole('alert').filter({ hasText: "A workbook must keep at least one visible sheet, so this one can't be hidden." })).toBeVisible();

    // Past the tabs come New sheet, Sheet menu and the note.
    await tab(page, 'Lines').focus();
    for (let i = 0; i < 3; i++) await page.keyboard.press(tabKey(page));
    const note = page.getByRole('button', { name: '1 hidden' });
    await expect(note).toBeFocused();
    await page.keyboard.press('Enter');
    const dialog = page.getByRole('dialog', { name: 'Unhide a sheet' });
    await expect(dialog.getByRole('radio', { name: 'Summary' })).toBeFocused();
    await page.keyboard.press('Enter');
    await expect.poll(() => tabNames(page)).toEqual(['Lines', 'Summary']);
    await expect(tab(page, 'Summary')).toHaveAttribute('aria-selected', 'true');
    await expect(tab(page, 'Summary')).toBeFocused();
  });

  test('inserts a sheet before the one whose menu is open, from the menu key', async ({ page }) => {
    await page.keyboard.press('ContextMenu');
    await choose(page, 'Insert sheet');
    await expect.poll(() => tabNames(page)).toEqual(['Sheet1', 'Lines', 'Summary']);
    await expect(tab(page, 'Sheet1')).toBeFocused();
    // Escape closes the menu and gives the focus back to the tab.
    await page.keyboard.press('Shift+F10');
    await page.keyboard.press('Escape');
    await expect(page.getByRole('menu')).toBeHidden();
    await expect(tab(page, 'Sheet1')).toBeFocused();
  });
});

test('the sheet menu opens from a button too, for touch screens', async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
  await page.getByRole('button', { name: 'Sheet menu' }).click();
  await expect(page.getByRole('menu', { name: 'Sheet “Lines”' })).toBeVisible();
  await page.getByRole('menuitem', { name: 'Move right' }).click();
  await expect.poll(() => tabNames(page)).toEqual(['Summary', 'Lines']);
});

test('dragging a tab moves its sheet', async ({ page }) => {
  await openFile(page, 'orders.xlsx', await ordersWorkbook());
  // Dropped on the far half of Summary's tab, so after it, moving in steps as a person's drag does, so
  // every browser sends the drag's events along the way.
  const from = (await tab(page, 'Lines').boundingBox())!;
  const box = (await tab(page, 'Summary').boundingBox())!;
  await page.mouse.move(from.x + from.width / 2, from.y + from.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2, { steps: 8 });
  await page.mouse.move(box.x + box.width - 4, box.y + box.height / 2, { steps: 4 });
  await page.mouse.up();
  await expect.poll(() => tabNames(page)).toEqual(['Summary', 'Lines']);
  await expect(tab(page, 'Lines')).toHaveAttribute('aria-selected', 'true');
});

for (const theme of ['Default', 'Dark']) {
  test(`the sheet menu and dialogs pass an accessibility scan, ${theme} theme`, async ({ page }) => {
    const violations = async () => {
      await page.waitForTimeout(300);
      const { violations } = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa']).analyze();
      return violations.flatMap((v) => v.nodes.map((n) => `${v.id}: ${n.target.join(' ')}`));
    };
    await openFile(page, 'orders.xlsx', await ordersWorkbook());
    await page.getByRole('combobox', { name: 'Theme' }).selectOption(theme);
    await tab(page, 'Summary').click({ button: 'right' });
    await expect(page.getByRole('menu', { name: 'Sheet “Summary”' })).toBeVisible();
    expect(await violations()).toEqual([]);
    await page.getByRole('menuitem', { name: 'Rename…' }).click();
    await page.getByRole('textbox', { name: 'Sheet name' }).fill('a/b');
    expect(await violations()).toEqual([]);
    await page.keyboard.press('Escape');
    await tab(page, 'Summary').click({ button: 'right' });
    await page.getByRole('menuitem', { name: 'Hide', exact: true }).click();
    await page.getByRole('button', { name: '1 hidden' }).click();
    await expect(page.getByRole('dialog', { name: 'Unhide a sheet' })).toBeVisible();
    expect(await violations()).toEqual([]);
  });
}
