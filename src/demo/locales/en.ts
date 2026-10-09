/**
 * The demo page's own text, in English. It lives in the host's i18next
 * instance (namespace "demo"), apart from the editor's messages.
 */
export const en = {
  'page.title': 'Workbook Editor POC',

  'picker.engine': 'Calculation engine',
  'picker.theme': 'Theme',
  'picker.language': 'Language',
  'theme.default': 'Default',
  'theme.indigo': 'Indigo',
  'theme.paper': 'Paper',
  'theme.dark': 'Dark',
  'theme.contrast': 'High contrast',

  'welcome.heading': 'Edit an Excel workbook in the browser',
  'welcome.examples': 'Other examples:',
  'welcome.exampleEmbedded': 'a supplier portal with the editor embedded',
  'welcome.exampleAngular': 'the element in an Angular app',
  'welcome.lede':
    'Open a .xlsx file, change cells, paste from Excel, drag to fill, and save it back as .xlsx. Formulas recalculate as you edit. Charts, pivot tables, macros and formatting are kept in the saved file, even the ones this editor can’t display.',
  'welcome.choose': '<strong>Choose a workbook</strong> or drop it anywhere on this page',
  'welcome.fileTypes': '.xlsx, .xlsm or .csv, processed entirely in your browser',
  'welcome.schemaSamples': 'Samples checked by a schema',
  'welcome.schemaCount_one': '{{count, number}} schema is registered.',
  'welcome.schemaCount_other': '{{count, number}} schemas are registered.',
  'welcome.manageSchemas': 'Manage, import or create schemas',
  'welcome.plainSamples': 'Plain workbooks',
  'welcome.sampleMissing': 'Sample {{file}} was not found. Run “npm run samples” to create it.',

  'sample.priceList.title': 'Price list',
  'sample.priceList.detail': 'Records as rows. Types, patterns, a pick list, an async catalogue check, and a lookup that fills new rows.',
  'sample.forecast.title': 'Quarterly forecast',
  'sample.forecast.detail': 'Records as columns, with calculated rows and a rule that quarters follow on from each other.',
  'sample.budget.title': 'Half-year budget',
  'sample.budget.detail': 'A fixed grid of categories by month, a sign-off cell and an async check against a budget service.',
  'sample.priceListOld.title': 'Price list, old layout',
  'sample.priceListOld.detail': 'One renamed column, so the schema doesn’t match and the file opens as a plain spreadsheet.',
  'sample.salesReport.title': 'Sales report',
  'sample.salesReport.detail': 'Charts, a chart sheet, a table, validation, notes, pictures, a hidden sheet, and LET, dynamic array and Ctrl+Shift+Enter formulas.',
  'sample.orders.title': 'Order log',
  'sample.orders.detail':
    'Excel’s own data validation: dropdown lists (inline, named, OFFSET and dependent INDIRECT) and number, date, text-length and custom-formula rules.',
  'sample.formatting.title': 'Formatting',
  'sample.formatting.detail': 'Borders, text running into empty cells, notes, conditional formatting, table styles and frozen rows.',
  'sample.pivot.title': 'Pivot table',
  'sample.pivot.detail': 'A pivot table plus filled-down formulas, which Excel stores in a compact shared form.',
  'sample.big.title': '100,000 rows',
  'sample.big.detail': '1.1 million cells and 300,000 formulas. Create it with npm run samples.',
};

export type DemoKey = keyof typeof en extends infer K ? (K extends `${infer Base}_${'one' | 'other'}` ? Base : K) : never;
