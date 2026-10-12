import { useEffect, useRef, useState } from 'react';
import { Trans, useTranslation } from 'react-i18next';
import { useFormat } from './format';
import { NS } from './i18n';
import { summarise } from './parse';
import { startReview } from './Review';
import { listSubmissions, type Submission } from './services';

const SAMPLES = `${import.meta.env.BASE_URL}samples/`;

export function Home() {
  const { t } = useTranslation(NS);
  const f = useFormat();
  const input = useRef<HTMLInputElement>(null);
  const [submissions, setSubmissions] = useState<Submission[] | undefined>();
  const [problem, setProblem] = useState('');

  useEffect(() => {
    void listSubmissions().then(setSubmissions);
  }, []);

  const trySample = async () => {
    const res = await fetch(`${SAMPLES}supplier-price-list.xlsx`);
    if (!res.ok) return setProblem(t('home.sampleMissing'));
    startReview('supplier-price-list.xlsx', await res.arrayBuffer());
  };

  return (
    <>
      <h1 tabIndex={-1}>{t('home.title')}</h1>
      <p className="lede">{t('home.lede')}</p>

      <section className="actions" aria-label={t('home.actions')}>
        <button className="primary" onClick={() => input.current?.click()}>
          {t('home.upload')}
        </button>
        <input
          ref={input}
          type="file"
          accept=".xlsx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          hidden
          onChange={async (e) => {
            const file = e.target.files?.[0];
            e.target.value = '';
            if (file) startReview(file.name, await file.arrayBuffer());
          }}
        />
        <button onClick={() => void trySample()}>{t('home.sample')}</button>
        <a className="button" href={`${SAMPLES}supplier-template.xlsx`} download>
          {t('home.template')}
        </a>
      </section>
      {problem && (
        <p className="problem" role="alert">
          {problem}
        </p>
      )}

      <section aria-labelledby="checks-heading" className="how">
        <h2 id="checks-heading">{t('home.howTitle')}</h2>
        <ul>
          <li>
            <Trans t={t} i18nKey="home.howTypes" components={{ b: <strong /> }} />
          </li>
          <li>
            <Trans t={t} i18nKey="home.howExcel" components={{ b: <strong /> }} />
          </li>
          <li>
            <Trans t={t} i18nKey="home.howBusiness" components={{ b: <strong /> }} />
          </li>
          <li>
            <Trans t={t} i18nKey="home.howSave" components={{ b: <strong /> }} />
          </li>
        </ul>
      </section>

      <section aria-labelledby="submitted-heading">
        <h2 id="submitted-heading">{t('home.submitted')}</h2>
        {submissions === undefined ? (
          <p className="muted">{t('home.loading')}</p>
        ) : submissions.length === 0 ? (
          <p className="muted">{t('home.none')}</p>
        ) : (
          <table className="list">
            <thead>
              <tr>
                <th scope="col">{t('home.colFile')}</th>
                <th scope="col">{t('home.colSubmitted')}</th>
                <th scope="col" className="num">
                  {t('home.colProducts')}
                </th>
                <th scope="col" className="num">
                  {t('home.colOrderValue')}
                </th>
                <th scope="col">{t('home.colStatus')}</th>
              </tr>
            </thead>
            <tbody>
              {submissions.map((s) => {
                const sum = summarise(s.lines);
                return (
                  <tr key={s.id}>
                    <th scope="row">
                      <a href={`#/submissions/${s.id}`}>{s.fileName}</a>
                    </th>
                    <td>{f.moment(s.submittedAt)}</td>
                    <td className="num">{f.number(sum.products)}</td>
                    <td className="num">{f.money(sum.orderValue, sum.currency)}</td>
                    <td>
                      <span className={`status status-${s.status}`}>{t(`status.${s.status}`)}</span>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </section>
    </>
  );
}
