import { useEffect, useState } from 'react';
import { useTranslation } from 'react-i18next';
import { createWorkbookEditor, WorkbookEditor, type EditorHandle, type EditorTheme } from '@sitepen/spreadsheet';
import { go } from './App';
import { NS } from './i18n';
import { readPriceLines } from './parse';
import { supplierPriceList } from './schema';
import { approvedPrices, getSubmission, newId, putSubmission, type Submission } from './services';

/** The portal's brand, passed to the editor as theme tokens. */
export const PORTAL_THEME: EditorTheme = {
  accent: '#0f766e',
  accentStrong: '#115e59',
  accentSoft: '#e6f4f2',
  accentBorder: '#99d5cc',
  accentTint: 'rgba(15, 118, 110, 0.12)',
  focusRing: 'rgba(15, 118, 110, 0.4)',
  fieldSoft: '#e9f5f3',
  fieldText: '#0b5a54',
  infoBannerBg: '#eef6f5',
  infoBannerBorder: '#cfe6e2',
  infoBannerText: '#0b4743',
  gridSelection: '#0f766e',
  gridSelectionFill: 'rgba(15, 118, 110, 0.1)',
  gridCutFill: 'rgba(15, 118, 110, 0.06)',
  gridHeaderSelectedBg: '#e6f4f2',
  gridHeaderSelectedText: '#0f766e',
  gridLabelFill: '#e9f5f3',
  radius: '8px',
};

/** A workbook chosen on the home page, waiting for the review page. */
let pending: { name: string; bytes: ArrayBuffer } | undefined;

export function startReview(name: string, bytes: ArrayBuffer): void {
  pending = { name, bytes };
  go('/review');
}

/**
 * The editor, embedded: the supplier checks and corrects the workbook, can
 * download a copy at any point, and submits it, which hands the portal the
 * file and its values. `id` edits a submission again.
 */
export function Review({ id }: { id?: string }) {
  const { t, i18n } = useTranslation(NS);
  const [editor, setEditor] = useState<EditorHandle | null>(null);
  const [fileName, setFileName] = useState('');
  const [missing, setMissing] = useState(false);

  useEffect(() => {
    let current = true;
    const editor = createWorkbookEditor({
      schemas: [supplierPriceList],
      language: i18n.language,
      theme: PORTAL_THEME,
      // The portal chooses the file, and suppliers need neither the schema designer
      // nor the report on what the file holds beyond what they fill in.
      features: { open: false, schemas: false, report: false },
      saveLabel: { en: 'Submit price list', de: 'Preisliste einreichen' },
      user: 'Supplier',
      // The host's save: read the values into the portal's own records and store them.
      onSave: async (saved) => {
        const submission: Submission = {
          id: id ?? newId(),
          fileName: saved.fileName,
          submittedAt: new Date().toISOString(),
          status: 'submitted',
          bytes: saved.bytes,
          lines: readPriceLines(saved, approvedPrices()),
          changedCells: saved.editedCount,
          warnings: saved.schema?.warnings ?? 0,
        };
        await putSubmission(submission);
        pending = undefined;
        // Once the editor has finished its save.
        setTimeout(() => go(`/submissions/${submission.id}`), 0);
      },
    });
    setEditor(editor);
    const source = id
      ? getSubmission(id).then((s) => s && { name: s.fileName, bytes: s.bytes.slice().buffer as ArrayBuffer })
      : Promise.resolve(pending && { ...pending, bytes: pending.bytes.slice(0) });
    void source.then((src) => {
      if (!current) return;
      if (!src) return setMissing(true);
      setFileName(src.name);
      // A file the editor can't open is shown as such in the editor.
      editor.open(src.bytes, src.name).catch(() => undefined);
    });
    return () => {
      current = false;
      editor.destroy();
    };
    // The language is passed on below, without re-creating the editor.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [id]);

  useEffect(() => {
    if (!editor) return;
    const follow = (lng: string) => editor.setLanguage(lng);
    i18n.on('languageChanged', follow);
    return () => i18n.off('languageChanged', follow);
  }, [editor, i18n]);

  if (missing) {
    return (
      <>
        <h1 tabIndex={-1}>{t('review.missingTitle')}</h1>
        <p>{t('review.missing')}</p>
        <p>
          <a href="#/">{t('review.back')}</a>
        </p>
      </>
    );
  }

  return (
    <>
      <nav className="crumbs" aria-label={t('app.breadcrumbs')}>
        <a href="#/">{t('nav.priceLists')}</a>
        <span aria-hidden="true">›</span>
        <span aria-current="page">{fileName || t('review.loading')}</span>
      </nav>
      <div className="review-head">
        <div>
          <h1 tabIndex={-1}>{t(id ? 'review.titleAgain' : 'review.title')}</h1>
          <p className="lede">{t('review.lede')}</p>
        </div>
        <ul className="legend" aria-label={t('review.legend')}>
          <li>
            <span className="mark mark-error" aria-hidden="true" />
            {t('review.legendError')}
          </li>
          <li>
            <span className="mark mark-warning" aria-hidden="true" />
            {t('review.legendWarning')}
          </li>
          <li>
            <span className="mark mark-pending" aria-hidden="true" />
            {t('review.legendPending')}
          </li>
        </ul>
      </div>
      <div className="editor-frame">{editor && <WorkbookEditor editor={editor} />}</div>
      <p className="muted small">{t('review.drafts')}</p>
    </>
  );
}
