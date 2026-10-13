---
title: Languages
description: The editor in English, German or a language you add, with dates and numbers in the user's locale.
sidebar:
  order: 4
---

The editor's own text comes in English and German. Dates and numbers are read and shown in the user's locale, separately from the language of the interface.

```ts
createWorkbookEditor({
  language: 'de',     // the editor's text
  locale: 'de-CH',    // dates and numbers: 1'234.50, 31.01.2026
});
```

Both default to the browser's. Change the language while a workbook is open with `editor.setLanguage('de')`.

## Adding a language

`translations` takes message catalogues by language, keyed as in the English one. Start from `en`, which has every message:

```ts
import { en } from '@sitepen/spreadsheet';

const fr = { ...en, 'toolbar.save': 'Enregistrer', /* … */ };
createWorkbookEditor({ language: 'fr', translations: { fr } });
```

Messages are i18next JSON, with `{{placeholders}}` and plurals (`_one`, `_other`).

## Rewording a message

The same option rewords a few messages in a built-in language:

```ts
translations: { en: { 'toolbar.save': 'Submit' } },
```

For the Save button alone, `saveLabel` is simpler.

## Your own text

Text your app gives the editor (a schema's names, descriptions and messages, `saveLabel`) can be one string or one per language. The editor shows the one that fits its language best, falling back from `de-CH` to `de` and then to English:

```ts
saveLabel: { en: 'Submit price list', de: 'Preisliste einreichen' },
```
