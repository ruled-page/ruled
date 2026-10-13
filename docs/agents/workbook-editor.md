# Workbook Editor (`@sitepen/spreadsheet`)

A spreadsheet editor embedded in web apps. It opens real Excel files (.xlsx, .xlsm, CSV), checks them against schemas with sync and async hooks, and hands the saved .xlsx and its values back to the app. Full docs: {{SITE}}/llms-full.txt

## Embedding

- React: `const editor = useWorkbookEditor({ ...options, file })` (null on the first render; `file` is `{ name, bytes }`, opened when ready and again when it's another file), `{editor && <WorkbookEditor editor={editor} />}`, and `import '@sitepen/spreadsheet/style.css'`. The hook destroys the editor on unmount; theme, language, user, schemas and customLists follow `options` live; callbacks may be inline; other options are read once (change them with a new `key`). `useEditorState(editor)` gives `{ status, fileName, editedCount, unsaved }` as React state. Never make an editor with `useMemo` or during render: use the hook, or `createWorkbookEditor` in an effect with `destroy()` in its cleanup.
- Anything else: `import '@sitepen/spreadsheet/standalone'` (defines `<workbook-editor>`, styles and worker included). Set `element.options = {...}`, use `src="/file.xlsx"`, and `save-handler="event"` to get a `save` event whose `detail.waitUntil(promise)` keeps the save pending.
- Or `mountWorkbookEditor(element, options)` from the standalone build, which returns the handle.
- The editor fills its container: give the container (or the element) a height.
- Open a file with `await editor.open(arrayBuffer, 'name.xlsx')`. When the app chooses files, set `features: { open: false }`.
- Receive saves with `onSave: async (saved) => {...}`: `saved.bytes` (the .xlsx), `saved.values` (schema-shaped: `values[sheet].tables[tableKey]` is an array of records keyed by field key, `values[sheet].cells[key]`), `saved.changes`, `saved.schema` ({ id, errors, warnings }). Throw to show an error and keep the changes; throw `SaveConflictError` when the stored file changed.
- Change at runtime without closing the workbook: `setTheme`, `setLanguage`, `setUser`, `registerSchema`, `removeSchema`, `setCustomLists`. Other options need a new editor.

## Schemas

Write with `defineSchema({...})` for types. Shape: `{ id, name, version?, sheets: [{ name, tables?: [...], cells?: [...], outside?, validate? }], saveWithErrors?, hookTimeoutMs?, numbers?, validate?, onChange?, beforeSave? }`.

- Table: `{ key, at: 'A1', orientation?: 'rows' | 'columns', header?, records?: { min?, max?, labels? }, corner?, fields, checks?, rules? }`. `at` is the top-left cell including headings. `records.labels` makes a fixed grid.
- Field: `{ key, label, type, required?, unique?, min?, max?, decimals?, minLength?, maxLength?, pattern?, patternHint?, values?, readOnly?, formula?: 'allow' | 'forbid' | 'require', description?, validate? }`. `label` matches the heading text in the file, ignoring case and spaces (or a RegExp).
- Types: `text`, `number`, `integer`, `percent` (fractions: 12.5% is 0.125), `date` (bounds as 'YYYY-MM-DD'; hooks get a `Date` at UTC midnight), `boolean`, `enum` (with `values`).
- Single cells: `cells: [{ key, ref: 'B8', label?: { ref: 'A8', text: 'Approved by' }, type, ... }]`.
- Declarative table checks: `monotonic`, `firstEquals`, `sum`, `compare`, `requiredIf`, `uniqueTogether`.
- Cells outside tables are read-only unless the sheet says `outside: 'editable'`.
- Text for people can be a string or `{ en: '…', de: '…' }`.

## Hooks (code, may be async)

- Field `validate(value, { signal, language, values, record, table, sheet })`: return nothing when fine, a message for an error, or `{ message, severity: 'warning' }`. Not called for empty cells (use `required`). Pass `signal` to `fetch`: it aborts when the value changes again. Results are kept until the value changes.
- Table `rules: [(table, ctx) => IssueInput[]]`: issues are `{ record, field, message, severity? }`. Records have `index`, `label`, `values`, `ref(field)`. Prefer one batched request here over one per cell.
- Sheet `validate(sheet, ctx)` and workbook `validate(ctx)`: issues can use `ref: 'Sheet!B9'`. `ctx.sheet(name)?.table(key)`, `.cell(key)`.
- `onChange(ctx)`: after each edit; `ctx.changes` (each `{ sheet, ref, table?, record?, field?, before, after }`); `ctx.set({ sheet, table, record, field } | { sheet, ref }, value)` writes cells as part of the same undo step.
- `beforeSave(ctx)`: return nothing or `true` to allow, a message to stop the save (can't be overridden).
- A hook that throws is an error on its cell; one slower than `hookTimeoutMs` (default 10000) is reported and its result ignored. Saving waits for running checks. Errors block saving unless `saveWithErrors: 'confirm'`.
- Hooks aren't part of schema JSON (`schemaToJSON` / `schemaFromJSON`); attach them in code.

## Look and language

- `theme: { accent, surface, font, gridHeaderBg, ... }` (tokens are CSS custom properties `--wbe-*` on `.wbe`); `colorScheme: 'dark'` for dark themes.
- `language` ('en', 'de' built in; others via `translations`), `locale` for dates and numbers.
- `features`: `open`, `save`, `download`, `schemas`, `report`, `charts`, `format`, `csv`, `autoComplete`, each on unless `false`.
