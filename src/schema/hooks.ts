import type { WorkbookSchema } from './types';

/** A sheet's identity across copies of a schema: its name, as written. */
const sheetName = (name: unknown) => JSON.stringify(name);

/**
 * `next` with the code hooks of `from`, a schema with the same id: the
 * workbook's checks, business logic and last check before saving, and each
 * sheet's, table's, field's and cell's checks wherever the sheet name, table
 * key and field or cell key are still there. JSON can't carry code, so an
 * edited or imported copy of a host's schema would otherwise lose them.
 */
export function withHooksFrom(next: WorkbookSchema, from: WorkbookSchema): WorkbookSchema {
  return {
    ...next,
    validate: from.validate,
    onChange: from.onChange,
    beforeSave: from.beforeSave,
    sheets: next.sheets.map((sheet) => {
      const old = from.sheets.find((s) => sheetName(s.name) === sheetName(sheet.name));
      if (!old) return sheet;
      return {
        ...sheet,
        validate: old.validate,
        tables: sheet.tables?.map((table) => {
          const oldTable = old.tables?.find((t) => t.key === table.key);
          if (!oldTable) return table;
          return {
            ...table,
            rules: oldTable.rules,
            fields: table.fields.map((f) => ({ ...f, validate: oldTable.fields.find((o) => o.key === f.key)?.validate })),
          };
        }),
        cells: sheet.cells?.map((c) => ({ ...c, validate: old.cells?.find((o) => o.key === c.key)?.validate })),
      };
    }),
  };
}
