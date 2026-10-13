import { expect, test } from '@playwright/test';
import { goTo, openSample, setCell, shown } from './helpers';

test('a pivot table follows its source as it is edited, grows with a new item, and refreshes on demand', async ({ page }) => {
  await openSample(page, 'Pivot table');
  await page.getByRole('tab', { name: 'Pivot' }).click();
  await expect(shown(page, 'B5')).toHaveText(/^1,?315$/);
  const north = Number((await shown(page, 'B5').textContent())!.replace(/,/g, ''));

  // Data!C2 is one of North's rows: 100 more units there.
  await page.getByRole('tab', { name: 'Data' }).click();
  await goTo(page, 'C2');
  const units = Number(await page.getByRole('combobox', { name: 'Cell contents' }).inputValue());
  await setCell(page, 'C2', String(units + 100));
  await page.getByRole('tab', { name: 'Pivot' }).click();
  await expect(shown(page, 'B5')).toHaveText(new RegExp(`^${(north + 100).toLocaleString('en-US').replace(',', ',?')}$`));
  await expect(shown(page, 'A8')).toHaveText('Grand Total');

  // A region new to the data gets a row of its own.
  await page.getByRole('tab', { name: 'Data' }).click();
  await setCell(page, 'A3', 'Central');
  await page.getByRole('tab', { name: 'Pivot' }).click();
  await expect(shown(page, 'A8')).toHaveText('Central');
  await expect(shown(page, 'A9')).toHaveText('Grand Total');

  // Its cells stay read-only, and Refresh is in their menu.
  await goTo(page, 'B5');
  await page.keyboard.press('Shift+F10');
  await expect(page.getByRole('menuitem', { name: 'Refresh all' })).toBeVisible();
  await page.getByRole('menuitem', { name: /^Refresh/ }).first().click();
  await expect(shown(page, 'A9')).toHaveText('Grand Total');
});

test('the PivotTable Fields pane changes a pivot table’s fields, filters its items and summarises its values', async ({ page }) => {
  await openSample(page, 'Pivot table');
  await page.getByRole('tab', { name: 'Pivot' }).click();
  await goTo(page, 'B5');
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Show Field List' }).click();
  await expect(page.getByRole('tab', { name: 'PivotTable Fields' })).toHaveAttribute('aria-selected', 'true');
  const pane = page.getByRole('complementary');

  // Ticking a text field puts it down the rows, under the region.
  // It ticks once the pivot table is laid out again.
  await pane.getByRole('checkbox', { name: 'Product' }).click();
  await expect(pane.getByRole('checkbox', { name: 'Product' })).toBeChecked();
  await expect(shown(page, 'A5')).toHaveText('Gadget');

  // Its menu moves it across the columns.
  await pane.getByRole('button', { name: 'Product', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Move to Columns' }).click();
  await expect(shown(page, 'B4')).toHaveText('Gadget');
  await expect(shown(page, 'A5')).toHaveText('East');

  // The regions filtered: West left out of the rows and the totals.
  await pane.getByRole('button', { name: 'Filter Region' }).click();
  await pane.getByRole('checkbox', { name: 'West' }).uncheck();
  await pane.getByRole('button', { name: 'OK' }).click();
  await expect(shown(page, 'A8')).toHaveText('Grand Total');

  // The units averaged rather than summed.
  await pane.getByRole('button', { name: 'Sum of Units', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Summarize Values By' }).click();
  await page.getByRole('menuitemradio', { name: 'Average' }).click();
  await expect(shown(page, 'A3')).toHaveText('Average of Units');

  // Dragged to the filters, the region goes above the table, showing the regions it filters to.
  await pane.getByText('Region', { exact: true }).first().dragTo(pane.getByRole('region', { name: 'Filters' }));
  await expect(shown(page, 'A1')).toHaveText('Region');
  await expect(shown(page, 'B1')).toHaveText('(Multiple Items)');

  // Each change is a step of its own.
  await goTo(page, 'H20');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(shown(page, 'A1')).toHaveText('');
  await expect(shown(page, 'A5')).toHaveText('East');
});

test('Insert PivotTable puts an empty pivot table on a new sheet, ready for its fields', async ({ page }) => {
  await openSample(page, 'Pivot table');
  await page.getByRole('tab', { name: 'Data' }).click();
  await goTo(page, 'B5');
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menuitem', { name: 'Insert PivotTable…' }).click();
  const dialog = page.getByRole('dialog', { name: 'Insert PivotTable' });
  await expect(dialog.getByLabel('Table or range')).toHaveValue('Data!$A$1:$E$101');
  await dialog.getByRole('button', { name: 'OK' }).click();

  // A new sheet before the data, with the Fields pane open beside the empty pivot table.
  await expect(page.getByRole('tab', { name: 'Sheet1' })).toBeVisible();
  const pane = page.getByRole('complementary');
  await expect(pane.getByText('PivotTable1')).toBeVisible();
  await pane.getByRole('checkbox', { name: 'Region' }).click();
  await expect(pane.getByRole('checkbox', { name: 'Region' })).toBeChecked();
  await pane.getByRole('checkbox', { name: 'Units' }).click();
  await expect(shown(page, 'A3')).toHaveText('Row Labels');
  await expect(shown(page, 'A8')).toHaveText('Grand Total');
  await expect(shown(page, 'B8')).toHaveText(/^4,?554$/);
});
