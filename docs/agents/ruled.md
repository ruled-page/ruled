# Ruled (`@ruled/sheet`)

A spreadsheet editor embedded in web apps. It opens real Excel files (.xlsx, .xlsm, CSV), checks them against schemas with sync and async hooks, and hands the saved .xlsx and its values back to the app. Full docs: {{SITE}}/llms-full.txt

## Embedding

- React: `const editor = useWorkbookEditor({ ...options, file })` (null on the first render; `file` is `{ name, bytes }`, opened when ready and again when it's another file), `{editor && <WorkbookEditor editor={editor} />}`, and `import '@ruled/sheet/style.css'`. The hook destroys the editor on unmount; theme, language, user, schemas and customLists follow `options` live; callbacks may be inline; other options are read once (change them with a new `key`). `useEditorState(editor)` gives `{ status, fileName, editedCount, unsaved }` as React state. Never make an editor with `useMemo` or during render: use the hook, or `createWorkbookEditor` in an effect with `destroy()` in its cleanup.
- React Server Components (Next.js App Router): the package is `'use client'`. A Server Component renders `<Workbook file={{ name, bytes }} onSave={serverAction} {...plainOptions} />` and imports `@ruled/sheet/style.css` in the root layout. `onSave` gets `SavedWorkbookData` (no `cell`/`range`) and answers `{ error }` or `{ conflict: true | message }` rather than throwing. Options that are functions (onEvent, onError, onLinkClick, createWorker, schemas with hooks) go in the host's own `'use client'` component wrapping `<Workbook>`.
- Anything else: `import '@ruled/sheet/standalone'` (defines `<workbook-editor>`, styles and worker included). Set `element.options = {...}`, use `src="/file.xlsx"`, and `save-handler="event"` to get a `save` event whose `detail.waitUntil(promise)` keeps the save pending.
- Or `mountWorkbookEditor(element, options)` from the standalone build, which returns the handle.
- The editor fills its container: give the container (or the element) a height.
- With a Content Security Policy, allow `font-src data:` (the grid's fonts), `img-src blob:` (pictures in workbooks; refused, they show as placeholders) and the worker: `worker-src 'self'` for the React build, `worker-src blob: data:` and `style-src 'unsafe-inline'` for the standalone one.
- Password-protected workbooks need no code: the editor asks for the password, decrypts in the browser, saves the workbook encrypted with it again, and keeps no drafts of it. Node's `recalculate(bytes, { password })` takes one.
- Open a file with `await editor.open(arrayBuffer, 'name.xlsx')`. When the app chooses files, set `features: { open: false }`.
- Receive saves with `onSave: async (saved) => {...}`: `saved.bytes` (the .xlsx), `saved.values` (schema-shaped: `values[sheet].tables[tableKey]` is an array of records keyed by field key, `values[sheet].cells[key]`), `saved.changes`, `saved.schema` ({ id, errors, warnings }). Throw (or answer `{ error }`) to show an error and keep the changes; throw `SaveConflictError` (or answer `{ conflict: true }`) when the stored file changed.
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
- `features`: `open`, `save`, `download`, `schemas`, `report`, `charts`, `format`, `csv`, `autoComplete`, `sourceSave`, each on unless `false`.

## Plugins

Using: `plugins: [oneDrive({ clientId })]` in the options (the element's `options` property; read once, so a new editor to change them). Plugins are trusted code on the page. A source turns Open into a menu and adds "Save to <source>" and "Save a copy" entries beside the host's Save; `features.sourceSave: false` stops all writes to sources. `editor.origin` and `SavedWorkbook.origin` say where a workbook came from; `onEvent` hears `savedToSource` and `pluginError`. Docs: {{SITE}}/plugins/writing.md

Writing: a plain object, types only from `@ruled/sheet/plugin` (never a runtime import of `@ruled/sheet`; `@ruled/sheet` is a peer dependency).

- `{ id: 'com.example.thing' (reverse-DNS), apiVersion: 1, translations?: { en: {...}, de: {...} }, setup(ctx) { ...; return cleanup } }`. `setup` runs a moment after the editor is made.
- `ctx`: `editor` (the handle), `signal` (aborted on destroy), `on(event => …)`, `commands`, `files`, `ui`, `text(pluginText)`, `limits.maxFileBytes`, `storage`.
- Text (`PluginText`): a string, `{ en, de }`, or `{ key: 'plugin.<id>.<key>', params }` for a message from `translations` (placeholders `{{name}}`). Use `ctx.text()` for DOM you draw. Hosts reword `plugin.<id>.<key>` in their `translations`.
- `commands.register({ id, label, icon?, where: 'toolbar' | 'more' | 'cellMenu', run(ctx), enabled?(ctx), visible?(ctx), preload?() })`; `cellMenu` gets `{ selection: { sheet, active, ranges } }`. `run` is inside the click: open popups before any `await`, load SDKs in `preload`. `commands.refresh()` after plugin state that `enabled`/`visible` read changes.
- `files.registerSource({ id, label, icon?, preload?, pick(): Promise<RemoteFile | null>, save?(file, origin): Promise<FileOrigin>, saveAs?(file): Promise<FileOrigin | null>, fetch?(origin): Promise<RemoteFile | null> })`. `RemoteFile = { bytes, origin }`; `FileOrigin = { source (the source id), name, ref (JSON), version?, writable?, location? }`. `save` resolves with the new origin (new `version`); throw an error with `name: 'SaveConflictError'` when the file changed since. `null` means cancelled.
- `files.open(remoteFile)`, `files.fetch(origin)` (through whichever source the origin names), `files.current()` (`{ name, from: 'device' | 'host' | 'source', origin?, size, bytes() }`; read it on the `opened` event).
- `ui.toast(kind, text)`, `ui.confirm({ title, body, confirmLabel })`, `ui.dialog({ title, render(body, close), actions? })` (resolves the action id or null), `ui.registerPanel({ id, title, icon?, render(body) })`. `render` fills plain DOM, may return a cleanup, and runs again when the language changes. Style with `wbe-btn`, `wbe-input` and `--wbe-*` tokens only.
- Errors thrown from any of these are caught, shown naming the plugin and reported as `pluginError`; a throw in `setup` removes the plugin. Give errors a `text` (PluginText) to translate them. Your own DOM listeners aren't caught.
- `storage.get<T>(key)`, `set(key, json)`, `delete(key)`: `localStorage` under `ruled-plugin:<id>`, ~5 MB for the whole page, may silently keep nothing. No tokens, no large files.
