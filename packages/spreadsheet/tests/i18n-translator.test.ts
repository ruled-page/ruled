import i18next from 'i18next';
import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { I18nextProvider } from 'react-i18next';
import { describe, expect, it } from 'vitest';
import { createI18n, english, errorText, isMessage, msg, pickLanguage, TextError, translator } from '../src/i18n';
import { useT } from '../src/i18n/react';

const t = (language: string, translations?: Parameters<typeof createI18n>[1]) => translator(createI18n(language, translations));

describe('a host’s own languages', () => {
  it('can add a language that isn’t built in, falling back to English for what it leaves out', () => {
    const pl = t('pl', { pl: { 'toolbar.open': 'Otwórz' } });
    expect(pl('toolbar.open')).toBe('Otwórz');
    expect(pl('toolbar.redo')).toBe('Redo');
    expect(pl.language).toBe('pl');
  });

  it('use the plural forms the language has, beyond one and other', () => {
    const pl = t('pl', {
      pl: {
        'toolbar.changes_one': '{{count, number}} zmiana',
        'toolbar.changes_few': '{{count, number}} zmiany',
        'toolbar.changes_many': '{{count, number}} zmian',
        'toolbar.changes_other': '{{count, number}} zmiany',
      },
    });
    expect([1, 3, 5, 12, 22, 25, 1.5].map((count) => pl('toolbar.changes', { count }))).toEqual([
      '1 zmiana',
      '3 zmiany',
      '5 zmian',
      '12 zmian',
      '22 zmiany',
      '25 zmian',
      '1,5 zmiany',
    ]);
  });

  it('can reword part of a built-in language and keep the rest of it', () => {
    const de = t('de', { de: { 'toolbar.open': 'Datei öffnen' } });
    expect(de('toolbar.open')).toBe('Datei öffnen');
    expect(de('toolbar.undo')).toBe('Rückgängig');
    // Other instances keep the built-in wording.
    expect(t('de')('toolbar.open')).toBe('Öffnen');
  });

  it('write right to left when the language does', () => {
    expect(t('fa', { fa: { 'toolbar.open': 'باز کردن' } }).dir).toBe('rtl');
    expect(t('en-US').dir).toBe('ltr');
  });
});

describe('the translator', () => {
  it('formats dates and lists the language’s way', () => {
    const day = new Date(2026, 9, 7, 12);
    expect(t('en').date(day)).toBe('Oct 7, 2026');
    expect(t('de').date(day.getTime())).toBe('07.10.2026');
    expect(t('en').date(day, { dateStyle: 'long' })).toBe('October 7, 2026');
    expect(t('en').list(['Prices', 'Budget', 'Notes'])).toBe('Prices, Budget, and Notes');
    expect(t('en').list(['Prices', 'Budget', 'Notes'], 'disjunction')).toBe('Prices, Budget, or Notes');
    expect(t('de').list(['Preise', 'Budget', 'Notizen'])).toBe('Preise, Budget und Notizen');
    expect(t('de').list(['Preise', 'Budget'], 'disjunction')).toBe('Preise oder Budget');
  });

  it('shows nothing for missing text, and host-written text in its language', () => {
    const de = t('de');
    expect(de.text(null)).toBe('');
    expect(de.text(undefined)).toBe('');
    expect(de.localized(undefined)).toBe('');
    expect(de.text({ en: 'Price', de: 'Preis' })).toBe('Preis');
  });

  it('leaves a list inside a message alone, for the message to format', () => {
    // A list of strings looks like text in several languages, but isn't.
    const en = t('en', { en: { 'schemaChip.title': 'Schema: {{name, list}}' } });
    expect(en.text(msg('schemaChip.title', { name: ['Prices', 'Budget'] }))).toBe('Schema: Prices and Budget');
  });

  it('speaks English through an i18next instance that has no language yet', () => {
    const tr = translator(i18next.createInstance());
    expect(tr.language).toBe('en');
    expect(tr.number(1234.5)).toBe('1,234.5');
    expect(tr.localized({ de: 'Preis', en: 'Price' })).toBe('Price');
  });

  it('is what React components get from useT, in the editor’s language', () => {
    const Label = () => {
      const tr = useT();
      return createElement('span', { lang: tr.language, dir: tr.dir }, tr('toolbar.open'));
    };
    const render = (language: string) => renderToStaticMarkup(createElement(I18nextProvider, { i18n: createI18n(language) }, createElement(Label)));
    expect(render('de')).toBe('<span lang="de" dir="ltr">Öffnen</span>');
    expect(render('en-GB')).toBe('<span lang="en-GB" dir="ltr">Open</span>');
  });
});

describe('messages as data', () => {
  it('carry params only when given some', () => {
    expect(msg('toolbar.open')).toStrictEqual({ key: 'toolbar.open' });
    expect(msg('toolbar.changes', { count: 2 })).toStrictEqual({ key: 'toolbar.changes', params: { count: 2 } });
  });

  it('are told apart from other values by their string key', () => {
    expect([null, undefined, 'toolbar.open', { key: 1 }, { en: 'Price' }].map(isMessage)).toEqual([false, false, false, false, false]);
    expect(isMessage({ key: 'toolbar.open' })).toBe(true);
  });
});

describe('errors people see', () => {
  it('carry their message across, as text to translate', () => {
    const e = new TextError(msg('core.open.notWorkbook'));
    expect(e).toBeInstanceOf(Error);
    expect(e.name).toBe('TextError');
    expect(e.message).toBe('core.open.notWorkbook');
    expect(errorText(e)).toEqual({ key: 'core.open.notWorkbook' });
    expect(english.text(errorText(e))).toBe('This file is not an Excel workbook (.xlsx or .xlsm).');
    expect(t('de').text(errorText(e))).toBe('Diese Datei ist keine Excel-Arbeitsmappe (.xlsx oder .xlsm).');
  });

  it('can carry words already written', () => {
    const e = new TextError('Already in words.');
    expect(e.message).toBe('Already in words.');
    expect(errorText(e)).toBe('Already in words.');
  });

  it('come out as words whatever was thrown', () => {
    // As the worker hands an error back: plain data with the text.
    expect(errorText({ message: 'x', text: 'From the worker.' })).toBe('From the worker.');
    expect(errorText({ text: msg('core.open.oldFormat') })).toEqual({ key: 'core.open.oldFormat' });
    expect(errorText(new RangeError('Out of range.'))).toBe('Out of range.');
    expect(errorText('a string')).toBe('a string');
    expect(errorText(42)).toBe('42');
    expect(errorText(null)).toBe('null');
    expect(errorText(undefined)).toBe('undefined');
    // An Error whose text isn't text to show falls back to its message.
    expect(errorText(Object.assign(new Error('Plain.'), { text: 7 }))).toBe('Plain.');
  });

  it('keep host-written text in several languages', () => {
    const e = new TextError({ en: 'Too high', de: 'Zu hoch' });
    expect(english.text(errorText(e))).toBe('Too high');
    expect(e.message).toBe('Too high');
    expect(translator(createI18n('de')).text(errorText(e))).toBe('Zu hoch');
  });
});

describe('host-written text in several languages', () => {
  const price = { EN: 'Price', de: 'Preis', 'DE-ch': 'Preis (CHF)', zh: '价格' };

  it('picks the exact language, then its base, then English, then the first, ignoring case', () => {
    expect(pickLanguage(price, 'de-CH')).toBe('Preis (CHF)');
    expect(pickLanguage(price, 'DE-at')).toBe('Preis');
    expect(pickLanguage(price, 'zh-Hant-TW')).toBe('价格');
    expect(pickLanguage(price, 'fr')).toBe('Price');
    expect(pickLanguage(price, '')).toBe('Price');
    expect(pickLanguage({ it: 'Prezzo', es: 'Precio' }, 'fr')).toBe('Prezzo');
  });

  it('is empty when there is none', () => {
    expect(pickLanguage({}, 'en')).toBe('');
    expect(pickLanguage(null, 'en')).toBe('');
    expect(pickLanguage(undefined, 'en')).toBe('');
    expect(pickLanguage('As written', 'de')).toBe('As written');
  });
});
