import { useEffect, useState } from 'react';
import { useTranslation } from 'react-i18next';
import { go } from './App';
import { useFormat } from './format';
import { NS } from './i18n';
import { CHANGE_THRESHOLD, priceChange, summarise } from './parse';
import { approveSubmission, deleteSubmission, getSubmission, type Submission } from './services';

/** A submitted price list as the portal shows it, built from the values the editor handed over. */
export function SubmissionPage({ id }: { id: string }) {
  const { t } = useTranslation(NS);
  const f = useFormat();
  const [submission, setSubmission] = useState<Submission | null | undefined>();
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [notice, setNotice] = useState('');

  useEffect(() => {
    void getSubmission(id).then((s) => setSubmission(s ?? null));
  }, [id]);

  if (submission === undefined) return <p className="muted">{t('home.loading')}</p>;
  if (submission === null) {
    return (
      <>
        <h1 tabIndex={-1}>{t('submission.missingTitle')}</h1>
        <p>
          <a href="#/">{t('review.back')}</a>
        </p>
      </>
    );
  }

  const s = submission;
  const sum = summarise(s.lines);
  const largest = Math.max(1, ...sum.byCategory.map((c) => c.orderValue));
  const download = () => {
    const url = URL.createObjectURL(new Blob([s.bytes as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = Object.assign(document.createElement('a'), { href: url, download: s.fileName });
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  return (
    <>
      <nav className="crumbs" aria-label={t('app.breadcrumbs')}>
        <a href="#/">{t('nav.priceLists')}</a>
        <span aria-hidden="true">›</span>
        <span aria-current="page">{s.fileName}</span>
      </nav>
      <div className="submission-head">
        <div>
          <h1 tabIndex={-1}>{s.fileName}</h1>
          <p className="muted">
            {t('submission.submittedAt', { when: f.moment(s.submittedAt) })}{' '}
            <span className={`status status-${s.status}`}>{t(`status.${s.status}`)}</span>
          </p>
        </div>
        <div className="buttons">
          {s.status === 'submitted' && (
            <button
              className="primary"
              onClick={async () => {
                setSubmission(await approveSubmission(s));
                setNotice(t('submission.approvedNotice'));
              }}
            >
              {t('submission.approve')}
            </button>
          )}
          <button onClick={() => go(`/review/${s.id}`)}>{t('submission.edit')}</button>
          <button onClick={download}>{t('submission.download')}</button>
          {confirmDelete ? (
            <span className="confirm" role="group" aria-label={t('submission.deleteQuestion')}>
              <span>{t('submission.deleteQuestion')}</span>
              <button
                className="danger"
                onClick={async () => {
                  await deleteSubmission(s.id);
                  go('/');
                }}
              >
                {t('submission.deleteYes')}
              </button>
              <button onClick={() => setConfirmDelete(false)}>{t('submission.deleteNo')}</button>
            </span>
          ) : (
            <button className="quiet" onClick={() => setConfirmDelete(true)}>
              {t('submission.delete')}
            </button>
          )}
        </div>
      </div>
      <p className="notice" role="status">
        {notice}
      </p>

      <section aria-label={t('submission.summary')} className="stats">
        <div>
          <span className="stat">{f.number(sum.products)}</span>
          <span>{t('submission.products', { count: sum.products })}</span>
        </div>
        <div>
          <span className="stat">{f.money(sum.orderValue, sum.currency)}</span>
          <span>{t('submission.orderValue')}</span>
        </div>
        <div>
          <span className="stat">{sum.averageChange === undefined ? '—' : f.change(sum.averageChange)}</span>
          <span>{t('submission.averageChange')}</span>
        </div>
        <div className={sum.bigChanges ? 'flag' : undefined}>
          <span className="stat">{f.number(sum.bigChanges)}</span>
          <span>{t('submission.bigChanges', { count: sum.bigChanges, threshold: f.change(CHANGE_THRESHOLD).replace('+', '') })}</span>
        </div>
        <div>
          <span className="stat">{f.number(sum.newProducts)}</span>
          <span>{t('submission.newProducts', { count: sum.newProducts })}</span>
        </div>
      </section>

      <section aria-labelledby="categories-heading">
        <h2 id="categories-heading">{t('submission.byCategory')}</h2>
        <ul className="bars">
          {sum.byCategory.map((c) => (
            <li key={c.category}>
              <span className="bar-label">{c.category}</span>
              <span className="bar" aria-hidden="true">
                <span style={{ width: `${(c.orderValue / largest) * 100}%` }} />
              </span>
              <span className="bar-value">
                {t('submission.categoryLine', { value: f.money(c.orderValue, sum.currency), count: c.products })}
              </span>
            </li>
          ))}
        </ul>
      </section>

      <section aria-labelledby="lines-heading">
        <h2 id="lines-heading">{t('submission.lines')}</h2>
        <p className="muted small">{t('submission.linesNote')}</p>
        <div className="table-scroll" tabIndex={0} role="region" aria-labelledby="lines-heading">
          <table className="list lines">
            <thead>
              <tr>
                <th scope="col">{t('col.sku')}</th>
                <th scope="col">{t('col.description')}</th>
                <th scope="col">{t('col.category')}</th>
                <th scope="col" className="num">
                  {t('col.price')}
                </th>
                <th scope="col" className="num">
                  {t('col.approved')}
                </th>
                <th scope="col" className="num">
                  {t('col.change')}
                </th>
                <th scope="col" className="num">
                  {t('col.minimumOrder')}
                </th>
                <th scope="col">{t('col.validFrom')}</th>
                <th scope="col" className="num">
                  {t('col.orderValue')}
                </th>
              </tr>
            </thead>
            <tbody>
              {s.lines.map((l, i) => {
                const change = priceChange(l);
                const big = change !== undefined && Math.abs(change) > CHANGE_THRESHOLD;
                return (
                  <tr key={i} className={big ? 'flag' : undefined}>
                    <th scope="row" className="mono">
                      {l.sku}
                    </th>
                    <td>{l.description}</td>
                    <td>
                      <span className="tag">{l.category}</span>
                    </td>
                    <td className="num">{f.money(l.price, l.currency)}</td>
                    <td className="num">{l.approved ? f.money(l.approved.price, l.approved.currency) : <span className="muted">{t('submission.new')}</span>}</td>
                    <td className={`num ${change === undefined ? '' : change > 0 ? 'up' : change < 0 ? 'down' : ''}`}>
                      {change === undefined ? '—' : (
                        <>
                          <span aria-hidden="true">{change > 0 ? '▲ ' : change < 0 ? '▼ ' : ''}</span>
                          {f.change(change)}
                          {big && <span className="visually-hidden"> {t('submission.flagged')}</span>}
                        </>
                      )}
                    </td>
                    <td className="num">{f.number(l.minimumOrder)}</td>
                    <td>{f.day(l.validFrom)}</td>
                    <td className="num">{f.money(l.orderValue, l.currency)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </section>
    </>
  );
}
