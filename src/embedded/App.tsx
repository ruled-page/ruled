import { useEffect, useRef, useState } from 'react';
import { useTranslation } from 'react-i18next';
import { Home } from './Home';
import { chooseLanguage, LANGUAGES, NS } from './i18n';
import { Review } from './Review';
import { SubmissionPage } from './SubmissionPage';

export type Route = { page: 'home' } | { page: 'review'; id?: string } | { page: 'submission'; id: string };

function parseRoute(hash: string): Route {
  const parts = hash.replace(/^#\/?/, '').split('/').filter(Boolean);
  if (parts[0] === 'review') return { page: 'review', ...(parts[1] ? { id: parts[1] } : {}) };
  if (parts[0] === 'submissions' && parts[1]) return { page: 'submission', id: parts[1] };
  return { page: 'home' };
}

export function go(path: string): void {
  location.hash = path;
}

function useRoute(): Route {
  const [route, setRoute] = useState(() => parseRoute(location.hash));
  useEffect(() => {
    const onChange = () => setRoute(parseRoute(location.hash));
    window.addEventListener('hashchange', onChange);
    return () => window.removeEventListener('hashchange', onChange);
  }, []);
  return route;
}

function LanguagePicker() {
  const { t, i18n } = useTranslation(NS);
  return (
    <label className="language">
      <span className="visually-hidden">{t('app.language')}</span>
      <select value={i18n.language} onChange={(e) => chooseLanguage(e.target.value)}>
        {LANGUAGES.map((l) => (
          <option key={l.code} value={l.code} lang={l.code}>
            {l.name}
          </option>
        ))}
      </select>
    </label>
  );
}

/** The procurement portal the editor is embedded in. */
export function App() {
  const { t } = useTranslation(NS);
  const route = useRoute();
  const main = useRef<HTMLElement>(null);

  // A new page takes the focus to its heading, as a page load would.
  const key = JSON.stringify(route);
  const first = useRef(true);
  useEffect(() => {
    if (first.current) {
      first.current = false;
      return;
    }
    main.current?.querySelector<HTMLElement>('h1')?.focus();
  }, [key]);

  return (
    <div className="portal">
      <button className="skip-link" onClick={() => main.current?.focus()}>
        {t('app.skip')}
      </button>
      <header className="portal-header">
        <a className="brand" href="#/">
          <svg viewBox="0 0 24 24" width="22" height="22" aria-hidden="true">
            <path d="M3 17c3 0 3-2 6-2s3 2 6 2 3-2 6-2M3 21c3 0 3-2 6-2s3 2 6 2 3-2 6-2M12 3v9M12 3l6 7H6z" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
          </svg>
          <span>{t('app.name')}</span>
        </a>
        <nav aria-label={t('app.nav')}>
          <a href="#/" aria-current={route.page !== 'home' ? undefined : 'page'}>
            {t('nav.priceLists')}
          </a>
          <a href="./">{t('nav.editorDemo')}</a>
        </nav>
        <LanguagePicker />
      </header>
      <main id="main" ref={main} tabIndex={-1} className={route.page === 'review' ? 'wide' : undefined}>
        {route.page === 'home' && <Home />}
        {route.page === 'review' && <Review key={route.id ?? 'new'} id={route.id} />}
        {route.page === 'submission' && <SubmissionPage key={route.id} id={route.id} />}
      </main>
    </div>
  );
}
