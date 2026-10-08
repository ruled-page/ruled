import { expect, test, type Page } from '@playwright/test';
import { goTo, openFile, ordersWorkbook, shown } from './helpers';

/** Help while typing a formula, as Excel gives it: functions offered, argument hints, coloured references and F4. */

const formulaBar = (page: Page) => page.getByRole('combobox', { name: 'Cell contents' });

/** How many of the grid's pixels are drawn in this colour: the outline of a reference. */
async function gridPixels(page: Page, color: string): Promise<number> {
  return page.evaluate((css) => {
    const probe = document.createElement('canvas').getContext('2d')!;
    probe.fillStyle = css;
    probe.fillRect(0, 0, 1, 1);
    const [r, g, b] = probe.getImageData(0, 0, 1, 1).data;
    let n = 0;
    for (const canvas of document.querySelectorAll<HTMLCanvasElement>('canvas[data-testid="data-grid-canvas"]')) {
      const { data } = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height);
      for (let i = 0; i < data.length; i += 4) if (Math.abs(data[i] - r) + Math.abs(data[i + 1] - g) + Math.abs(data[i + 2] - b) < 6) n++;
    }
    return n;
  }, color);
}

/** A theme token's value, as the editor resolves it. */
const token = (page: Page, name: string) => page.evaluate((n) => getComputedStyle(document.querySelector('.wbe-root')!).getPropertyValue(n).trim(), name);

test.describe('typing a formula in the formula bar', () => {
  test.beforeEach(async ({ page }) => {
    await openFile(page, 'orders.xlsx', await ordersWorkbook());
    await goTo(page, 'E2');
    await formulaBar(page).click();
  });

  test('offers the functions a name could be, as a combobox; Tab inserts the highlighted one', async ({ page }) => {
    const input = formulaBar(page);
    await expect(input).toHaveAttribute('aria-expanded', 'false');
    await page.keyboard.type('=1+SU');
    const list = page.getByRole('listbox', { name: 'Functions' });
    await expect(list).toBeVisible();
    await expect(input).toHaveAttribute('aria-expanded', 'true');
    await expect(input).toHaveAttribute('aria-controls', (await list.getAttribute('id'))!);
    const options = list.getByRole('option');
    await expect(options.first()).toHaveText('SUBSTITUTE');
    await expect(options.first()).toHaveAttribute('aria-selected', 'true');
    await expect(input).toHaveAttribute('aria-activedescendant', (await options.first().getAttribute('id'))!);
    // What the highlighted function takes and does.
    await expect(page.locator('.wbe-fx-about')).toContainText('SUBSTITUTE(text, old_text, new_text, [instance_num])');

    await page.keyboard.type('M');
    await expect(options.first()).toHaveText('SUM');
    await page.keyboard.press('ArrowDown');
    await expect(options.nth(1)).toHaveAttribute('aria-selected', 'true');
    await expect(input).toHaveAttribute('aria-activedescendant', (await options.nth(1).getAttribute('id'))!);
    await page.keyboard.press('ArrowUp');
    await page.keyboard.press('Tab');
    await expect(input).toHaveValue('=1+SUM(');
    await expect(input).toBeFocused();
    await expect(list).toBeHidden();
    await page.keyboard.type('D2:D3)');
    await page.keyboard.press('Enter');
    await expect(shown(page, 'E2')).toHaveText('9');
  });

  test('Enter inserts the function too, and Escape closes the list without leaving the formula', async ({ page }) => {
    const input = formulaBar(page);
    await page.keyboard.type('=rou');
    await expect(page.getByRole('option', { name: 'ROUND', exact: true })).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(page.getByRole('listbox')).toBeHidden();
    await expect(input).toBeFocused();
    await expect(input).toHaveValue('=rou');
    await page.keyboard.type('n');
    await page.keyboard.press('Enter');
    await expect(input).toHaveValue('=ROUND(');
    await expect(input).toBeFocused();
  });

  test('keeps the highlighted function in view in a long list, and goes round from the top to the bottom', async ({ page }) => {
    await page.keyboard.type('=C');
    const options = page.getByRole('listbox').getByRole('option');
    await expect(options.first()).toBeVisible();
    expect(await options.count()).toBeGreaterThan(40);
    await expect(options.last()).not.toBeInViewport();
    await page.keyboard.press('ArrowUp');
    await expect(options.last()).toHaveAttribute('aria-selected', 'true');
    await expect(options.last()).toBeInViewport();
  });

  test('marks the functions it leaves for Excel to calculate, and inserts a function clicked', async ({ page }) => {
    await page.keyboard.type('=LA');
    await expect(page.getByRole('option', { name: 'LAMBDA Not calculated here' })).toBeVisible();
    await expect(page.getByRole('option', { name: 'LARGE', exact: true })).toBeVisible();
    await page.getByRole('option', { name: 'LARGE', exact: true }).click();
    await expect(formulaBar(page)).toHaveValue('=LARGE(');
    await expect(formulaBar(page)).toBeFocused();
  });

  test("shows the arguments of the function the caret is in, the one being typed in bold, and says it", async ({ page }) => {
    const input = formulaBar(page);
    await page.keyboard.type('=VLOOKUP(A2,');
    const hint = page.getByRole('tooltip');
    await expect(hint).toHaveText('VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])');
    await expect(hint.locator('b')).toHaveText('table_array');
    await expect(input).toHaveAttribute('aria-describedby', (await hint.getAttribute('id'))!);
    await expect(page.locator('.wbe-formula-bar [aria-live="polite"]')).toHaveText('VLOOKUP, argument 2: table_array');
    await page.keyboard.type('A2:D3,4,');
    await expect(hint.locator('b')).toHaveText('range_lookup');
    await page.keyboard.type('FALSE)');
    await expect(hint).toBeHidden();
    await page.keyboard.press('Enter');
    await expect(shown(page, 'E2')).toHaveText('6');
  });

  test('colours each reference, and outlines it on the grid in the same colour', async ({ page }) => {
    const blue = await token(page, '--wbe-formula-ref1');
    const red = await token(page, '--wbe-formula-ref2');
    expect(await gridPixels(page, blue)).toBe(0);
    await page.keyboard.type('=B2*C2+b2');
    const refs = page.locator('.wbe-formula-text .wbe-ref');
    await expect(refs).toHaveText(['B2', 'C2', 'b2']);
    const colors = await refs.evaluateAll((els) => els.map((el) => getComputedStyle(el).color));
    expect(colors[0]).toBe(colors[2]);
    expect(colors[0]).not.toBe(colors[1]);
    await expect.poll(() => gridPixels(page, blue)).toBeGreaterThan(20);
    await expect.poll(() => gridPixels(page, red)).toBeGreaterThan(20);
    // The field's own text is hidden under the coloured copy, which lines up with it.
    const field = await formulaBar(page).boundingBox();
    const copy = await page.locator('.wbe-formula-text > span').boundingBox();
    expect(Math.abs(copy!.x - field!.x - 10)).toBeLessThan(1);

    await page.keyboard.press('Escape');
    await expect.poll(() => gridPixels(page, blue)).toBe(0);
    await expect(page.locator('.wbe-formula-text')).toHaveCount(0);
  });

  test('F4 cycles the reference at the caret between relative and absolute', async ({ page }) => {
    const input = formulaBar(page);
    await page.keyboard.type('=SUM(B2:B3)*C2');
    const seen: string[] = [];
    for (let i = 0; i < 4; i++) {
      await page.keyboard.press('F4');
      seen.push(await input.inputValue());
    }
    expect(seen).toEqual(['=SUM(B2:B3)*$C$2', '=SUM(B2:B3)*C$2', '=SUM(B2:B3)*$C2', '=SUM(B2:B3)*C2']);
    for (let i = 0; i < 4; i++) await page.keyboard.press('ArrowLeft');
    await page.keyboard.press('F4');
    await expect(input).toHaveValue('=SUM($B$2:$B$3)*C2');
    // The caret is after the reference it changed.
    expect(await input.evaluate((el: HTMLInputElement) => el.selectionStart)).toBe('=SUM($B$2:$B$3'.length);
  });
});

test.describe('typing a formula in a cell', () => {
  const editor = (page: Page) => page.getByRole('textbox', { name: 'Contents of E2' });

  test.beforeEach(async ({ page }) => {
    await openFile(page, 'orders.xlsx', await ordersWorkbook());
    await goTo(page, 'E2');
  });

  test('offers the functions a name could be; Tab inserts one, Escape closes the list and a second Escape stops editing', async ({ page }) => {
    await page.keyboard.type('=SU', { delay: 20 });
    const cell = editor(page);
    await expect(cell).toBeFocused();
    const list = page.getByRole('listbox', { name: 'Functions' });
    await expect(list).toBeVisible();
    await expect(cell).toHaveAttribute('aria-autocomplete', 'list');
    await expect(cell).toHaveAttribute('aria-controls', (await list.getAttribute('id'))!);
    await page.keyboard.press('ArrowDown');
    const second = list.getByRole('option').nth(1);
    await expect(second).toHaveText('SUBTOTAL');
    await expect(cell).toHaveAttribute('aria-activedescendant', (await second.getAttribute('id'))!);
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('Tab');
    await expect(cell).toHaveValue('=SUM(');
    await expect(cell).toBeFocused();

    await page.keyboard.type('B2:B3)+ro', { delay: 20 });
    await expect(list).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(list).toBeHidden();
    await expect(cell).toHaveValue('=SUM(B2:B3)+ro');
    await page.keyboard.press('Escape');
    await expect(cell).toBeHidden();
    await expect(shown(page, 'E2')).toHaveText('');
  });

  test('Enter inserts a function rather than finishing the cell; then finishes it', async ({ page }) => {
    await page.keyboard.type('=1+ab', { delay: 20 });
    await page.keyboard.press('Enter');
    await expect(editor(page)).toHaveValue('=1+ABS(');
    await page.keyboard.type('-D3)', { delay: 20 });
    await page.keyboard.press('Enter');
    await expect(shown(page, 'E2')).toHaveText('3');
  });

  test('shows the arguments under the cell, colours references and outlines them, and F4 cycles one', async ({ page }) => {
    await page.keyboard.type('=ROUND(C3*B2', { delay: 20 });
    const hint = page.getByRole('tooltip');
    await expect(hint).toHaveText('ROUND(number, num_digits)');
    await expect(hint.locator('b')).toHaveText('number');
    await expect(editor(page)).toHaveAttribute('aria-describedby', (await hint.getAttribute('id'))!);
    await expect(page.locator('.wbe-cell-entry-text .wbe-ref')).toHaveText(['C3', 'B2']);
    const red = await token(page, '--wbe-formula-ref2');
    await expect.poll(() => gridPixels(page, red)).toBeGreaterThan(20);

    await page.keyboard.press('F4');
    await expect(editor(page)).toHaveValue('=ROUND(C3*$B$2');
    await page.keyboard.type(',1', { delay: 20 });
    await expect(hint.locator('b')).toHaveText('num_digits');
    await page.keyboard.type(')', { delay: 20 });
    await expect(hint).toBeHidden();
    await page.keyboard.press('Enter');
    await expect(shown(page, 'E2')).toHaveText('1');
    await expect.poll(() => gridPixels(page, red)).toBe(0);
    await goTo(page, 'E2');
    await expect(formulaBar(page)).toHaveValue('=ROUND(C3*$B$2,1)');
  });

  test('inserts a function clicked in the list, keeping the cell open', async ({ page }) => {
    await page.keyboard.type('=max', { delay: 20 });
    await page.getByRole('option', { name: 'MAXIFS', exact: true }).click();
    await expect(editor(page)).toHaveValue('=MAXIFS(');
    await expect(editor(page)).toBeFocused();
  });
});
