import type { PluginText } from '@ruled/sheet/plugin';

/** The plugin's id; its messages are `plugin.page.ruled.google-drive.<key>`. */
export const GOOGLE_DRIVE_PLUGIN_ID = 'page.ruled.google-drive';

/** The file source's id, which a workbook's `origin.source` names when it came from Google Drive. */
export const GOOGLE_DRIVE_SOURCE_ID = 'google-drive';

/** The connector's messages, which the editor shows as `plugin.page.ruled.google-drive.<key>`; a host's `translations` can reword them. */
export const translations = {
  en: {
    googleDrive: 'Google Drive',
    signOut: 'Sign out of {{source}}',
    signedOut: 'Signed out of {{source}} in this editor.',
    signInBody: 'Sign in with your Google account to open and save workbooks in {{source}}.',
    signIn: 'Sign in',
    signInAgain: 'Your {{source}} sign-in has expired. Sign in again to carry on.',
    notAllowed: 'This editor wasn’t given access to {{source}}. Sign in again and allow it to see the files you choose.',
    pickFile: 'Choose a workbook',
    pickFolder: 'Choose the folder to save the copy in',
    tooBig: '{{name}} is {{size}} MB, larger than the {{max}} MB this editor opens.',
    exportTooBig: '{{name}} is too large for Google to export as an Excel workbook: Google exports Sheets of up to 10 MB. Split it into smaller Sheets in Google Sheets to open them here.',
    googleSheet:
      '{{name}} is a Google Sheet, open here as an Excel copy. Saving makes a new .xlsx file in your {{source}}; the Google Sheet stays as it is. Apps Script and features only Google Sheets has aren’t in the copy.',
    googleSheetDropped:
      '{{name}} is a Google Sheet, open here as an Excel copy. Saving makes a new .xlsx file in your {{source}}; the Google Sheet stays as it is. These functions only Google Sheets has became the values Google last calculated:',
    droppedCells_one: '{{function}}, in {{count, number}} cell',
    droppedCells_other: '{{function}}, in {{count, number}} cells',
    refused: '{{source}} refused the request ({{status}}): {{reason}}',
  },
  de: {
    googleDrive: 'Google Drive',
    signOut: 'Von {{source}} abmelden',
    signedOut: 'In diesem Editor von {{source}} abgemeldet.',
    signInBody: 'Melden Sie sich mit Ihrem Google-Konto an, um Arbeitsmappen in {{source}} zu öffnen und zu speichern.',
    signIn: 'Anmelden',
    signInAgain: 'Ihre Anmeldung bei {{source}} ist abgelaufen. Melden Sie sich erneut an, um fortzufahren.',
    notAllowed: 'Dieser Editor hat keinen Zugriff auf {{source}} erhalten. Melden Sie sich erneut an und erlauben Sie den Zugriff auf die Dateien, die Sie auswählen.',
    pickFile: 'Arbeitsmappe auswählen',
    pickFolder: 'Ordner für die Kopie auswählen',
    tooBig: '{{name}} ist {{size}} MB groß, größer als die {{max}} MB, die dieser Editor öffnet.',
    exportTooBig:
      '{{name}} ist zu groß, um von Google als Excel-Arbeitsmappe exportiert zu werden: Google exportiert Tabellen bis 10 MB. Teilen Sie sie in Google Tabellen in kleinere Tabellen auf, um sie hier zu öffnen.',
    googleSheet:
      '{{name}} ist eine Google-Tabelle und hier als Excel-Kopie geöffnet. Beim Speichern entsteht eine neue .xlsx-Datei in Ihrem {{source}}; die Google-Tabelle bleibt unverändert. Apps Script und Funktionen, die nur Google Tabellen hat, sind nicht in der Kopie.',
    googleSheetDropped:
      '{{name}} ist eine Google-Tabelle und hier als Excel-Kopie geöffnet. Beim Speichern entsteht eine neue .xlsx-Datei in Ihrem {{source}}; die Google-Tabelle bleibt unverändert. Diese Funktionen, die nur Google Tabellen hat, wurden zu den Werten, die Google zuletzt berechnet hat:',
    droppedCells_one: '{{function}}, in {{count, number}} Zelle',
    droppedCells_other: '{{function}}, in {{count, number}} Zellen',
    refused: '{{source}} hat die Anfrage abgelehnt ({{status}}): {{reason}}',
  },
} as const;

/** A message's key as code uses it: plural forms share one ("droppedCells" for droppedCells_one and droppedCells_other). */
export type MessageKey = keyof (typeof translations)['en'] extends infer K ? (K extends `${infer Base}_${'one' | 'other'}` ? Base : K) : never;

/** One of the connector's messages, for the editor to show translated. */
export const msg = (key: MessageKey, params?: Record<string, unknown>): PluginText => ({ key: `plugin.${GOOGLE_DRIVE_PLUGIN_ID}.${key}`, params });
