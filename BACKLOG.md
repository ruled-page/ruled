# Backlog

Work that is deliberately not being done yet, with what it would take. The
editor is a client-side proof of concept; nothing here should add a server.

## Needs a decision first

- **The calculation engine's licence.** HyperFormula is used under its GPLv3
  licence key, which a closed-source product can't do. Before production:
  buy its commercial licence and pass the key in through an option, or
  replace the engine. The key is set in `src/core/calc.ts` and
  `src/core/engineLimits.ts`.
- **The package's name and licence.** It is still called `fake-spreadsheet`,
  is marked private, and has no LICENSE file or changelog. Choose the name it
  will be published under and its licence, then add both, a changelog and
  versioned releases.

## Needs a person

- **Screen-reader testing.** The editor works from the keyboard and passes
  automated accessibility checks in Chromium, Firefox and WebKit, but nobody
  has used it with NVDA, JAWS or VoiceOver. A tester should go through
  opening a file, moving around the grid, editing, the checks panel and the
  dialogs with each.

## Excel features left out for now

- **Refreshing pivot tables, and slicers.** Pivot tables are kept and shown,
  and Excel refreshes them on opening when their source data changed.
  Refreshing them in the editor would need a pivot engine.
- **Sparklines.** Kept in the file and moved with their cells; not drawn.
- **.xls, .xlsb and .ods files.** Refused with a message. A converter such
  as SheetJS could open them as .xlsx, losing charts and pivot tables.
- **Printing and PDF.** Print settings are kept in the file; nothing prints.
  A print view could use the browser's own printing.
- **What-if data tables and iterative calculation.** Their cells keep
  Excel's saved values. The engine has no iterative calculation.
- **Several people editing at once.** Live editing needs a server. A host
  can still turn down a save that conflicts (`SaveConflictError`).
- **Password-protected and sensitivity-labelled files.** Refused with a
  message. Password-protected files could be decrypted in the browser;
  sensitivity-labelled ones can only be decrypted through Microsoft's
  service.
