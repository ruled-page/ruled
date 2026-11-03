import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { I18nextProvider } from 'react-i18next';
import { describe, expect, it } from 'vitest';
import { forecastSchema } from '@ruled/demo/schemas';
import { NAMESPACE } from '../src/i18n';
import { schemaFromJSON } from '../src/schema/json';
import type { StoreOptions, WorkbookStore } from '../src/state/store';
import { SidePanel } from '../src/ui/SidePanel';
import { sample } from './helpers';
import { opened, testStore } from './storeHarness';

const schema = { ...schemaFromJSON(JSON.stringify(forecastSchema)), saveWithErrors: 'confirm' as const };

async function forecast(features?: StoreOptions['features']): Promise<WorkbookStore> {
  const store = await testStore({ schemas: [schema], features });
  store.setNarrow(false);
  await opened(store, sample('forecast.xlsx'), 'forecast.xlsx');
  return store;
}

const markup = (store: WorkbookStore) => renderToStaticMarkup(createElement(I18nextProvider, { i18n: store.i18n, defaultNS: NAMESPACE }, createElement(SidePanel, { store })));

describe('the schema issues panel', () => {
  it('opens by itself on a wide editor when a workbook has errors', async () => {
    const store = await forecast();
    expect(store.schema?.summary().errors).toBeGreaterThan(0);
    expect(store.panel).toBe('issues');
    expect(markup(store)).toContain('Checks');
    store.dispose();
  });

  it('stays shut with the feature off, while the checks still run', async () => {
    const store = await forecast({ issuesPanel: false });
    expect(store.schema?.summary().errors).toBeGreaterThan(0);
    expect(store.schema?.issues().length).toBeGreaterThan(0);
    expect(store.panel).toBeNull();
    // Nothing the editor does opens it.
    store.set('panel', 'issues');
    expect(store.panel).toBeNull();
    const saving = store.save();
    for (let i = 0; i < 100 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
    expect(store.confirmRequest).toBeDefined();
    store.confirmRequest?.resolve('cancel');
    await saving;
    expect(store.panel).toBeNull();
    store.dispose();
  });

  it('shows the other panels without a Checks tab with the feature off', async () => {
    const store = await forecast({ issuesPanel: false });
    store.set('panel', 'changes');
    const html = markup(store);
    expect(html).toContain('Changes');
    expect(html).not.toContain(`>${'Checks'}<`);
    store.dispose();
  });

  it('draws nothing for the issues panel with the feature off, even if it was set directly', async () => {
    const store = await forecast({ issuesPanel: false });
    store.panel = 'issues';
    expect(markup(store)).toBe('');
    store.dispose();
  });
});
