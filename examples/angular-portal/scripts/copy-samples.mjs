/**
 * Copies the demo's forecast workbook into samples/, which the app serves:
 * Angular's CLI only serves assets from inside its own project.
 */
import { copyFileSync, mkdirSync } from 'node:fs';

mkdirSync(new URL('../samples/', import.meta.url), { recursive: true });
copyFileSync(new URL('../../demo/public/samples/forecast.xlsx', import.meta.url), new URL('../samples/forecast.xlsx', import.meta.url));
