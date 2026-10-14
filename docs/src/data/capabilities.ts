/**
 * What the editor does with each part of a workbook, for the capabilities
 * page: edit it, show it, keep it in the file without showing it, or not.
 */
export type Support = 'edit' | 'show' | 'keep' | 'no';

export interface Capability {
  what: string;
  support: Support;
  note?: string;
  href?: string;
}

export const SUPPORT: Record<Support, { label: string; means: string }> = {
  edit: { label: 'Edit', means: 'Shown, and people can change it as in Excel.' },
  show: { label: 'Show', means: 'Shown, kept in the file, and changed only in part, as the note says.' },
  keep: { label: 'Keep', means: 'Not shown, but kept in the file as it was.' },
  no: { label: 'No', means: 'Not supported: the note says what happens instead.' },
};

/** `{functions}` and `{all}` are filled in with the engine's numbers. */
export const CAPABILITIES: { group: string; rows: Capability[] }[] = [
  {
    group: 'Files',
    rows: [
      { what: 'Excel workbooks (.xlsx, .xlsm)', support: 'edit', note: 'Saved back with every part the editor didn’t change as it was.', href: '/guides/opening-and-saving/' },
      { what: 'CSV, TSV and text files', support: 'edit', note: 'Open as a workbook of one sheet and save as .xlsx; any sheet downloads as CSV.', href: '/guides/csv-files/' },
      { what: 'Macros (VBA, Excel 4.0 macro sheets)', support: 'keep', note: 'Kept by default; the `macros` option can remove them when saving, or refuse such files.' },
      { what: 'ActiveX controls and embedded objects', support: 'keep', note: 'Or refused, with the `activeContent` option.' },
      { what: 'Links to other workbooks', support: 'keep', note: 'Their last values are shown and kept.' },
      { what: '.xls, .xlsb and .ods files', support: 'no', note: 'Refused, saying so.' },
      { what: 'Password-protected and sensitivity-labelled files', support: 'no', note: 'Refused, saying so.' },
    ],
  },
  {
    group: 'Cells and formatting',
    rows: [
      { what: 'Text, numbers and dates, typed in the user’s locale', support: 'edit' },
      { what: 'Number formats, built in and custom', support: 'edit' },
      { what: 'Fonts, colours, fills, borders, alignment, wrapping and indents', support: 'edit' },
      { what: 'Several fonts within one cell', support: 'show', note: 'Drawn in one font; all of them are kept.' },
      { what: 'Merged cells', support: 'edit' },
      { what: 'Column widths, row heights and AutoFit', support: 'edit' },
      { what: 'Hidden, grouped and frozen rows and columns', support: 'edit' },
      { what: 'Sheets that read right to left', support: 'show', note: 'Drawn from right to left, as the file sets them.' },
      { what: 'Conditional formatting', support: 'edit', note: 'Cell rules, top and bottom, averages, duplicates, colour scales, data bars and icon sets.' },
      { what: 'Data validation', support: 'edit', note: 'Lists and rules, as Excel’s Data Validation dialog sets them.' },
      { what: 'Excel 365 checkboxes', support: 'edit' },
      { what: 'Hyperlinks', support: 'edit' },
      { what: 'Notes and threaded comments', support: 'edit' },
      { what: 'Protected sheets', support: 'show', note: 'Their locked cells can’t be changed; the protection itself isn’t changed here.' },
    ],
  },
  {
    group: 'Formulas',
    rows: [
      { what: '{functions} of Excel’s {all} worksheet functions, calculated live', support: 'edit', note: 'The others keep Excel’s saved results.', href: '/reference/functions/' },
      { what: 'Dynamic arrays and their spills', support: 'edit' },
      { what: 'LET, LAMBDA, MAP, REDUCE, SCAN and names defined as LAMBDAs', support: 'edit' },
      { what: 'Defined names', support: 'edit', note: 'With Excel’s Name Manager.' },
      { what: 'Tables and structured references', support: 'edit', note: 'Used and kept; new tables aren’t made here.' },
      { what: 'Circular references and iterative calculation', support: 'edit' },
      { what: 'Names belonging to one sheet, written with it (Prices!Rate)', support: 'keep', note: 'Excel’s saved value is kept until Excel recalculates.' },
      { what: 'What-if data tables', support: 'keep', note: 'Excel’s saved values are kept.' },
    ],
  },
  {
    group: 'Working with data',
    rows: [
      { what: 'Sort', support: 'edit', note: 'By several levels, by a custom list, with case, or left to right.' },
      { what: 'Filters', support: 'edit', note: 'By values and by condition; filters by colour, icon or date period are kept.' },
      { what: 'Find and replace', support: 'edit' },
      { what: 'Fill handle, Fill Series, custom lists and AutoComplete', support: 'edit' },
      { what: 'Inserting and deleting rows, columns and cells, and moving cells', support: 'edit', note: 'References everywhere follow them, as in Excel.' },
      { what: 'Adding, renaming, moving, hiding and deleting sheets', support: 'edit' },
    ],
  },
  {
    group: 'Charts and pictures',
    rows: [
      { what: 'Column, bar, line, area, pie, doughnut and scatter charts', support: 'edit', note: 'Inserted from a selection, moved, resized, and their type, title, legend and data changed.' },
      { what: 'Radar, bubble, stock, 3-D and pie-of-pie charts', support: 'show', note: 'Drawn from the cells (3-D drawn flat), moved, resized and retitled.' },
      { what: 'Surface, waterfall, treemap, histogram, funnel and other newer charts', support: 'keep', note: 'Shown as a placeholder with the chart’s title.' },
      { what: 'Data labels, trendlines, error bars and gradient fills on charts', support: 'keep' },
      { what: 'Pictures and shapes', support: 'show', note: 'Shown as placeholders where they sit, and moved, resized and deleted.' },
      { what: 'Sparklines', support: 'keep' },
    ],
  },
  {
    group: 'Pivot tables',
    rows: [
      { what: 'Pivot tables', support: 'edit', note: 'Worked out live as their source changes, with the PivotTable Fields pane, filters and Show Values As.' },
      { what: 'Inserting a pivot table', support: 'edit', note: 'Over a range, a table or a name, on a new sheet or at a cell.' },
      { what: 'Pivot tables with grouped or calculated fields', support: 'show', note: 'As Excel saved them; Excel refreshes them on opening.' },
      { what: 'Slicers and timelines', support: 'keep' },
    ],
  },
  {
    group: 'Working together',
    rows: [
      { what: 'Undo, redo and the history of who changed what', support: 'edit' },
      { what: 'Reviewing every changed cell before saving', support: 'edit', note: 'With a schema’s `reviewBeforeSave`.' },
      { what: 'Several people editing at once', support: 'no', note: 'Your app can refuse a save that conflicts, with `SaveConflictError`.' },
      { what: 'Printing and PDF', support: 'no', note: 'Print settings are kept in the file.' },
    ],
  },
];

/** What the editor gives the app it's in. */
export const FOR_APPS: { what: string; note: string; href: string }[] = [
  { what: 'React, web component, Angular and plain HTML', note: 'A component and hook, a custom element, or a mount function.', href: '/start/install/' },
  { what: 'Schemas', note: 'Sheets, tables and cells described as data: types, bounds, patterns, lists, uniqueness and rules across records.', href: '/schemas/overview/' },
  { what: 'Async checks', note: 'Checks that call your services as people type, with cancellation, timeouts and saving that waits for them.', href: '/schemas/async-hooks/' },
  { what: 'Business logic', note: 'onChange fills cells from your systems in the same undo step; beforeSave has the last word.', href: '/schemas/business-logic/' },
  { what: 'Schema designer', note: 'Generate a schema from a workbook, edit it, and export or import it as JSON.', href: '/schemas/json-and-designer/' },
  { what: 'Saving to your app', note: 'The .xlsx, the values in your schema’s shape, every change and the history.', href: '/guides/opening-and-saving/' },
  { what: 'Drafts', note: 'Unsaved edits are kept in the browser and offered back.', href: '/guides/opening-and-saving/#drafts' },
  { what: 'Theming and icons', note: 'Design tokens for every colour, font, size and corner, and replacement icons.', href: '/guides/theming/' },
  { what: 'Languages and locales', note: 'English and German built in, your own added, dates and numbers in the user’s locale.', href: '/guides/languages/' },
  { what: 'Turning parts off', note: 'Hide opening, saving, the designer, the report, charts or the format bar.', href: '/guides/options-and-features/' },
  { what: 'Events and logging', note: 'What the editor does and how long it took, and when it stops working.', href: '/guides/events-and-handle/' },
  { what: 'Accessibility', note: 'Excel’s keys, screen reader support, forced colours and touch.', href: '/guides/keyboard-and-accessibility/' },
];
