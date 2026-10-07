import { describe, expect, it } from 'vitest';
import { BUILT_IN, createI18n, de, en, msg, translator } from '../src/i18n';

const t = (language: string, translations?: Parameters<typeof createI18n>[1]) => translator(createI18n(language, translations));

/** The {{placeholders}} a message uses. */
const placeholders = (message: string) => [...message.matchAll(/\{\{\s*(\w+)/g)].map((m) => m[1]).sort();

describe('translation', () => {
  it('picks plural forms and formats numbers for the language', () => {
    expect(t('en-US')('toolbar.changes', { count: 1 })).toBe('1 change');
    expect(t('en-US')('toolbar.changes', { count: 1234 })).toBe('1,234 changes');
    expect(t('de-DE')('toolbar.changes', { count: 1234 })).toBe('1.234 Änderungen');
  });

  it('falls back from a regional language to its base, then to English', () => {
    expect(t('de-CH')('toolbar.open')).toBe('Öffnen');
    expect(t('fr')('toolbar.open')).toBe('Open');
  });

  it('takes a host’s own languages and wording', () => {
    const fr = t('fr', { fr: { 'toolbar.open': 'Ouvrir' }, en: { 'toolbar.undo': 'Take back' } });
    expect(fr('toolbar.open')).toBe('Ouvrir');
    expect(fr('toolbar.redo')).toBe('Redo');
    expect(t('en', { en: { 'toolbar.undo': 'Take back' } })('toolbar.undo')).toBe('Take back');
  });

  it('translates messages handed over as data, including messages inside them', () => {
    const tr = t('en');
    expect(tr.text(msg('schemaChip.title', { name: 'Prices' }))).toBe('Schema: Prices');
    expect(tr.text(msg('schemaChip.label', { name: 'Prices', status: msg('schemaChip.errors', { count: 2 }) }))).toBe('Prices: 2 errors');
    expect(tr.text('As written in the workbook')).toBe('As written in the workbook');
    expect(t('de').text(msg('schemaChip.title', { name: { en: 'Prices', de: 'Preise' } }))).toBe('Schema: Preise');
  });

  it('picks host-written text in the right language', () => {
    const label = { en: 'Price', de: 'Preis', 'de-CH': 'Preis (CHF)' };
    expect(t('de-CH').localized(label)).toBe('Preis (CHF)');
    expect(t('de-AT').localized(label)).toBe('Preis');
    expect(t('fr').localized(label)).toBe('Price');
    expect(t('fr').localized({ it: 'Prezzo' })).toBe('Prezzo');
  });

  it('follows a change of language', () => {
    const i18n = createI18n('en');
    const tr = translator(i18n);
    void i18n.changeLanguage('de');
    expect(tr('toolbar.open')).toBe('Öffnen');
    expect(tr.number(1234.5)).toBe('1.234,5');
  });

  it('knows which languages are written right to left', () => {
    expect(t('ar').dir).toBe('rtl');
    expect(t('he-IL').dir).toBe('rtl');
    expect(t('de').dir).toBe('ltr');
  });
});

describe('the built-in catalogues', () => {
  it('translate every English message, with the same placeholders', () => {
    for (const [language, catalogue] of Object.entries(BUILT_IN)) {
      expect(Object.keys(catalogue).sort(), language).toEqual(Object.keys(en).sort());
      for (const [key, message] of Object.entries(catalogue)) {
        expect(placeholders(message), `${language} ${key}`).toEqual(placeholders(en[key as keyof typeof en]));
      }
    }
    expect(Object.keys(de).length).toBe(Object.keys(en).length);
  });
});
