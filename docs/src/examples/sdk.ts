/**
 * What the docs' examples import as `@sitepen/spreadsheet`: the library's
 * types and defineSchema. The editor itself is the standalone build, which a
 * live example loads when it scrolls into view, so pages stay light.
 */
import type { WorkbookSchema } from '../../../packages/spreadsheet/src/schema/types';

export type * from '../../../packages/spreadsheet/src/schema/types';
export type { EditorOptions, EditorTheme, SavedWorkbook } from '../../../packages/spreadsheet/src/editor/index';
export type { WorkbookEditorElement } from '../../../packages/spreadsheet/src/editor/element';

export function defineSchema(schema: WorkbookSchema): WorkbookSchema {
  return schema;
}
