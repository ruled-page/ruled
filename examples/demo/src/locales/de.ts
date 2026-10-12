import type { en } from './en';

/** The demo page's own text, in German; every English key needs one (the type checks it). */
export const de: Record<keyof typeof en, string> = {
  'page.title': 'Arbeitsmappen-Editor (Prototyp)',

  'toolbar.samples': 'Beispiele',
  'picker.theme': 'Design',
  'picker.language': 'Sprache',
  'theme.default': 'Standard',
  'theme.indigo': 'Indigo',
  'theme.paper': 'Papier',
  'theme.dark': 'Dunkel',
  'theme.contrast': 'Hoher Kontrast',

  'welcome.heading': 'Eine Excel-Arbeitsmappe im Browser bearbeiten',
  'welcome.examples': 'Weitere Beispiele:',
  'welcome.exampleEmbedded': 'ein Lieferantenportal mit eingebettetem Editor',
  'welcome.exampleAngular': 'das Element in einer Angular-App',
  'welcome.lede':
    'Öffnen Sie eine .xlsx-Datei, ändern Sie Zellen, fügen Sie aus Excel ein, füllen Sie durch Ziehen aus und speichern Sie sie wieder als .xlsx. Formeln werden beim Bearbeiten neu berechnet. Diagramme, PivotTables, Makros und Formatierungen bleiben in der gespeicherten Datei erhalten, auch die, die dieser Editor nicht anzeigen kann.',
  'welcome.choose': '<strong>Arbeitsmappe auswählen</strong> oder an einer beliebigen Stelle auf dieser Seite ablegen',
  'welcome.fileTypes': '.xlsx, .xlsm oder .csv, vollständig in Ihrem Browser verarbeitet',
  'welcome.schemaSamples': 'Von einem Schema geprüfte Beispiele',
  'welcome.schemaCount_one': '{{count, number}} Schema ist registriert.',
  'welcome.schemaCount_other': '{{count, number}} Schemas sind registriert.',
  'welcome.manageSchemas': 'Schemas verwalten, importieren oder erstellen',
  'welcome.plainSamples': 'Arbeitsmappen ohne Schema',
  'welcome.sampleMissing': 'Das Beispiel {{file}} wurde nicht gefunden. Führen Sie „npm run samples“ aus, um es zu erstellen.',

  'sample.priceList.title': 'Preisliste',
  'sample.priceList.detail':
    'Datensätze als Zeilen. Typen, Muster, eine Auswahlliste, eine asynchrone Katalogprüfung und ein Nachschlagen, das neue Zeilen ausfüllt.',
  'sample.forecast.title': 'Quartalsprognose',
  'sample.forecast.detail': 'Datensätze als Spalten, mit berechneten Zeilen und einer Regel, nach der die Quartale aufeinander folgen.',
  'sample.budget.title': 'Halbjahresbudget',
  'sample.budget.detail': 'Ein festes Raster aus Kategorien nach Monat, eine Zelle für die Freigabe und eine asynchrone Prüfung über einen Budgetdienst.',
  'sample.priceListOld.title': 'Preisliste, altes Layout',
  'sample.priceListOld.detail': 'Eine umbenannte Spalte: Das Schema passt daher nicht, und die Datei wird als gewöhnliche Arbeitsmappe geöffnet.',
  'sample.salesReport.title': 'Umsatzbericht',
  'sample.salesReport.detail':
    'Diagramme, ein Diagrammblatt, eine Tabelle, Datenüberprüfung, Kontrollkästchen, Notizen, Bilder, ein ausgeblendetes Blatt sowie LET-, dynamische Matrix- und Strg+Umschalt+Eingabe-Formeln.',
  'sample.orders.title': 'Bestellprotokoll',
  'sample.orders.detail':
    'Die Datenüberprüfung von Excel: Dropdownlisten (direkt eingegeben, benannt, mit OFFSET und abhängig mit INDIRECT) sowie Regeln für Zahlen, Datumsangaben, Textlänge und benutzerdefinierte Formeln.',
  'sample.formatting.title': 'Formatierung',
  'sample.formatting.detail': 'Rahmen, Text, der in leere Zellen hineinragt, Notizen, bedingte Formatierung, Tabellenformatvorlagen und fixierte Zeilen.',
  'sample.pivot.title': 'PivotTable',
  'sample.pivot.detail': 'Eine PivotTable und nach unten ausgefüllte Formeln, die Excel in einer kompakten, gemeinsamen Form speichert.',
  'sample.big.title': '100.000 Zeilen',
  'sample.big.detail': '1,1 Millionen Zellen und 300.000 Formeln. Erstellen Sie die Datei mit npm run samples.',
  'sample.lookups.title': '80.000 Nachschlagevorgänge',
  'sample.lookups.detail': '20.000 Bestellungen schlagen 20.000 Produkte nach, mit Summen je Kategorie: Ändern Sie einen Preis, und jede Bestellung wird neu berechnet. Erstellen Sie die Datei mit npm run samples.',
};
