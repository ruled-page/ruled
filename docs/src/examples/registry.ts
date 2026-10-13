/**
 * The live examples the pages show, by name: the workbook each opens, and
 * the options the editor gets, as a host would pass them.
 */
import type { EditorOptions } from '@ruled/sheet';
import { budget } from './budget';
import { forecast } from './forecast';
import { priceList } from './price-list';

export interface LiveExample {
  /** A sample workbook in public/samples. */
  file: string;
  /** Other workbooks it can switch to. */
  files?: { file: string; label: string }[];
  /** Whether it can switch between themes. */
  themes?: boolean;
  /** The editor's options; `show` writes into the panel under the editor. */
  options?: (show: (text: string) => void) => EditorOptions;
}

export const EXAMPLES: Record<string, LiveExample> = {
  'price-list': { file: 'price-list.xlsx', options: () => ({ schemas: [priceList] }) },
  budget: { file: 'budget.xlsx', options: () => ({ schemas: [budget] }) },
  forecast: { file: 'forecast.xlsx', options: () => ({ schemas: [forecast] }) },
  saving: {
    file: 'forecast.xlsx',
    options: (show) => ({
      // The sample has errors on purpose; here they ask before saving rather than stop it.
      schemas: [{ ...forecast, saveWithErrors: 'confirm' }],
      saveLabel: 'Save to the app',
      features: { open: false },
      // The host's save: here it shows what it received rather than uploading it.
      onSave: async (saved) => {
        show(
          JSON.stringify(
            { fileName: saved.fileName, bytes: saved.bytes.length, editedCount: saved.editedCount, schema: saved.schema, values: saved.values, changes: saved.changes },
            null,
            2,
          ),
        );
      },
    }),
  },
  workbooks: {
    file: 'kitchen-sink.xlsx',
    files: [
      { file: 'kitchen-sink.xlsx', label: 'Kitchen sink' },
      { file: 'chart-gallery.xlsx', label: 'Charts' },
      { file: 'pivot.xlsx', label: 'Pivot table' },
      { file: 'orders.xlsx', label: 'Orders' },
    ],
  },
  theming: { file: 'budget.xlsx', themes: true, options: () => ({ schemas: [budget] }) },
};
