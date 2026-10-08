import { Trans, useTranslation } from 'react-i18next';
import type { WorkbookStore } from '../state/store';
import { SITES } from '../sites';
import { Icon } from '../ui/icons';
import { DEMO_NS, i18next } from './i18n';
import type { DemoKey } from './locales/en';

interface Sample {
  file: string;
  title: DemoKey;
  detail: DemoKey;
}

export const SAMPLES: Sample[] = [
  { file: 'kitchen-sink.xlsx', title: 'sample.salesReport.title', detail: 'sample.salesReport.detail' },
  { file: 'orders.xlsx', title: 'sample.orders.title', detail: 'sample.orders.detail' },
  { file: 'formatting.xlsx', title: 'sample.formatting.title', detail: 'sample.formatting.detail' },
  { file: 'pivot.xlsx', title: 'sample.pivot.title', detail: 'sample.pivot.detail' },
  { file: 'big.xlsx', title: 'sample.big.title', detail: 'sample.big.detail' },
];

export const SCHEMA_SAMPLES: Sample[] = [
  { file: 'price-list.xlsx', title: 'sample.priceList.title', detail: 'sample.priceList.detail' },
  { file: 'forecast.xlsx', title: 'sample.forecast.title', detail: 'sample.forecast.detail' },
  { file: 'budget.xlsx', title: 'sample.budget.title', detail: 'sample.budget.detail' },
  { file: 'price-list-old.xlsx', title: 'sample.priceListOld.title', detail: 'sample.priceListOld.detail' },
];

export async function openSample(store: WorkbookStore, file: string): Promise<void> {
  const res = await fetch(`${import.meta.env.BASE_URL}samples/${file}`);
  const bytes = res.ok ? await res.arrayBuffer() : new ArrayBuffer(0);
  // Dev servers answer missing files with the app's HTML, so check for a zip signature.
  const head = new Uint8Array(bytes, 0, Math.min(2, bytes.byteLength));
  if (head[0] !== 0x50 || head[1] !== 0x4b) {
    // The demo's own message, so it comes from the demo's i18next rather than the editor's catalogue.
    store.toast('error', i18next.t('welcome.sampleMissing', { file }));
    return;
  }
  await store.open(file, bytes);
}

function SampleList({ store, samples, icon }: { store: WorkbookStore; samples: Sample[]; icon: 'schema' | 'sheet' }) {
  const { t } = useTranslation(DEMO_NS);
  return (
    <ul className="samples-list">
      {samples.map((s) => (
        <li key={s.file}>
          <button onClick={() => void openSample(store, s.file)}>
            <Icon name={icon} />
            <span>
              <strong>{t(s.title)}</strong>
              <span>{t(s.detail)}</span>
            </span>
          </button>
        </li>
      ))}
    </ul>
  );
}

export function Welcome({ store, onOpen }: { store: WorkbookStore; onOpen: () => void }) {
  const { t } = useTranslation(DEMO_NS);
  return (
    <div className="welcome">
      <div className="welcome-card">
        <h1>{t('welcome.heading')}</h1>
        <p className="lede">{t('welcome.lede')}</p>
        <p className="lede">
          {t('welcome.examples')} <a href={SITES.portal}>{t('welcome.exampleEmbedded')}</a>
          {SITES.angular && (
            <>
              {' · '}
              <a href={SITES.angular}>{t('welcome.exampleAngular')}</a>
            </>
          )}
        </p>
        <button className="wbe-dropzone" onClick={onOpen}>
          <Icon name="open" size={22} />
          <span>
            <Trans t={t} i18nKey="welcome.choose" components={{ strong: <strong /> }} />
          </span>
          <span className="wbe-muted">{t('welcome.fileTypes')}</span>
        </button>
        <h2>{t('welcome.schemaSamples')}</h2>
        <SampleList store={store} samples={SCHEMA_SAMPLES} icon="schema" />
        <p className="wbe-muted wbe-small schema-note">
          {t('welcome.schemaCount', { count: store.schemas.length })}{' '}
          <button className="wbe-link" onClick={() => store.set('designerOpen', true)}>
            {t('welcome.manageSchemas')}
          </button>
        </p>
        <h2>{t('welcome.plainSamples')}</h2>
        <SampleList store={store} samples={SAMPLES} icon="sheet" />
      </div>
    </div>
  );
}
