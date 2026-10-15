import type { RuledPlugin } from '@ruled/sheet/plugin';

/**
 * "New workbook" in the demo's Open menu: a source whose only file is a
 * blank one. It can't be saved back, so saving downloads it, or saves a
 * copy to another source.
 */
export function newWorkbook(): RuledPlugin {
  return {
    id: 'page.ruled.demo.new-workbook',
    apiVersion: 1,
    setup(ctx) {
      ctx.files.registerSource({
        id: 'new-workbook',
        label: { en: 'New workbook', de: 'Neue Arbeitsmappe' },
        // The zip writer loads only when a new workbook is made.
        pick: async () => ({
          bytes: (await import('./blankWorkbook')).blankWorkbook(ctx.text({ en: 'Sheet1', de: 'Tabelle1' })),
          origin: { source: 'new-workbook', name: ctx.text({ en: 'Book1.xlsx', de: 'Mappe1.xlsx' }), ref: null, writable: false },
        }),
      });
    },
  };
}
