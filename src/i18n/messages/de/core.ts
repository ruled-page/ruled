import type { core as en } from '../en/core';

/** German messages; every English key needs one (the type checks it). */
export const core: Record<keyof typeof en, string> = {
  'core.open.tooBig': 'Diese Datei ist {{size, number}} MB groß; dieser Editor akzeptiert höchstens {{max, number}} MB.',
  'core.open.oldFormat': 'Diese Datei ist kennwortgeschützt oder im alten .xls-Format. Nur .xlsx- und .xlsm-Dateien werden unterstützt.',
  'core.open.notWorkbook': 'Diese Datei ist keine Excel-Arbeitsmappe (.xlsx oder .xlsm).',
  'core.open.tooManyParts': 'Diese Datei hat mehr als {{max, number}} Teile, mehr als eine Arbeitsmappe haben sollte.',
  'core.open.unpacksTooBig': 'Diese Datei ist entpackt größer als {{max, number}} MB, mehr als dieser Editor aufnehmen kann.',
  'core.open.zipBomb': 'Der Teil {{part}} ist verdächtig stark komprimiert ({{ratio, number}}:1), daher wurde die Datei nicht geöffnet.',
  'core.open.notPackage': 'Diese Datei ist kein gültiges Office Open XML-Paket.',
  'core.open.damaged': 'Diese Datei ist beschädigt und wurde nicht geöffnet. Excel kann sie möglicherweise reparieren.',
  'core.open.noWorkbookPart': 'In dieser Datei fehlt der Teil mit der Arbeitsmappe.',
  'core.open.tooManyCells': 'Diese Arbeitsmappe hat mehr als {{max, number}} Zellen; mehr kann dieser Editor nicht aufnehmen.',
  'core.open.macrosBlocked':
    'Diese Arbeitsmappe enthält Makros (VBA oder Excel 4.0-Makroblätter), die hier nicht zulässig sind. Speichern Sie sie in Excel als .xlsx, und öffnen Sie dann diese Datei.',
  'core.open.activeContentBlocked': 'Diese Arbeitsmappe enthält ActiveX-Steuerelemente oder eingebettete Objekte, die hier nicht zulässig sind.',

  'core.finding.hiddenSheet.title': 'Ausgeblendetes Tabellenblatt „{{sheet}}“',
  'core.finding.hiddenSheet.detail': 'Bleibt in der Datei erhalten, wird aber wie in Excel nicht angezeigt.',
  'core.finding.chartsheet.title': 'Diagrammblatt „{{sheet}}“',
  'core.finding.chartsheet.detail':
    'Ein Diagramm, das ein ganzes Blatt füllt. Bleibt in der Datei erhalten und wird beim Öffnen der Datei von Excel aus den bearbeiteten Daten neu gezeichnet.',
  'core.finding.chartsheet.detailTitled':
    'Ein Diagramm mit dem Titel „{{title}}“, das ein ganzes Blatt füllt. Bleibt in der Datei erhalten und wird beim Öffnen der Datei von Excel aus den bearbeiteten Daten neu gezeichnet.',
  'core.finding.chartsheet.detailTyped':
    'Ein Diagramm vom Typ „{{type}}“, das ein ganzes Blatt füllt. Bleibt in der Datei erhalten und wird beim Öffnen der Datei von Excel aus den bearbeiteten Daten neu gezeichnet.',
  'core.finding.chartsheet.detailTypedTitled':
    'Ein Diagramm vom Typ „{{type}}“ mit dem Titel „{{title}}“, das ein ganzes Blatt füllt. Bleibt in der Datei erhalten und wird beim Öffnen der Datei von Excel aus den bearbeiteten Daten neu gezeichnet.',
  'core.finding.unsupportedSheet.title': 'Nicht unterstütztes Blatt „{{sheet}}“',
  'core.finding.unsupportedSheet.detail':
    'Ein Dialog- oder Makroblatt aus älteren Excel-Versionen. Bleibt in der Datei erhalten; wird in diesem Editor nicht angezeigt.',
  'core.finding.macroSheet.title': 'Excel 4.0-Makroblatt „{{sheet}}“',
  'core.finding.macroSheet.detail':
    'Enthält Excel 4.0-Makros, die beim Öffnen der Datei in Excel ausgeführt werden können. Bleibt in der Datei erhalten; wird hier weder angezeigt noch ausgeführt.',
  'core.finding.macroSheet.removedDetail': 'Enthält Excel 4.0-Makros. Das Blatt wird beim Speichern geleert, und die Datei wird als .xlsx gespeichert.',
  'core.finding.mergedCells.title': 'Verbundene Zellen über mehrere Zeilen',
  'core.finding.mergedCells.detail_one':
    '{{count, number}} verbundener Bereich erstreckt sich über mehr als eine Zeile. Sein Wert wird nur in der oberen linken Zelle angezeigt; die Verbindung bleibt in der Datei erhalten.',
  'core.finding.mergedCells.detail_other':
    '{{count, number}} verbundene Bereiche erstrecken sich über mehr als eine Zeile. Ihr Wert wird nur in der oberen linken Zelle angezeigt; die Verbindung bleibt in der Datei erhalten.',
  'core.finding.protectedSheet.title': 'Geschütztes Tabellenblatt',
  'core.finding.protectedSheet.detail':
    'In Excel geschützt: Hier lassen sich nur die nicht gesperrten Zellen bearbeiten, und Zeilen und Spalten lassen sich weder einfügen, löschen noch sortieren.',
  'core.finding.arrayFormulas.title': 'Matrixformeln',
  'core.finding.arrayFormulas.detail_one':
    '{{count, number}} Matrix- oder Überlaufformel erstreckt sich über mehr als eine Zelle. Sie behält den zuletzt von Excel berechneten Wert, Excel berechnet sie beim Öffnen neu, und ihre Zellen können hier nicht bearbeitet werden.',
  'core.finding.arrayFormulas.detail_other':
    '{{count, number}} Matrix- oder Überlaufformeln erstrecken sich über mehr als eine Zelle. Sie behalten die zuletzt von Excel berechneten Werte, Excel berechnet sie beim Öffnen neu, und ihre Zellen können hier nicht bearbeitet werden.',
  'core.finding.conditionalFormat.title': 'Bedingte Formatierung',
  'core.finding.conditionalFormat.detail_one':
    '{{count, number}} Regel. Wird hier angezeigt und beim Bearbeiten aktualisiert; bleibt für Excel in der Datei erhalten.',
  'core.finding.conditionalFormat.detail_other':
    '{{count, number}} Regeln. Werden hier angezeigt und beim Bearbeiten aktualisiert; bleiben für Excel in der Datei erhalten.',
  'core.finding.dataValidation.title': 'Datenüberprüfung',
  'core.finding.dataValidation.detail': '{{rules}}. Eingefügte Werte werden wie in Excel nicht geprüft.',
  'core.finding.dataValidation.lists_one': '{{count, number}} Listenregel zeigt eine Dropdownliste an und prüft eingegebene Werte wie Excel',
  'core.finding.dataValidation.lists_other': '{{count, number}} Listenregeln zeigen eine Dropdownliste an und prüfen eingegebene Werte wie Excel',
  'core.finding.dataValidation.rules_one':
    '{{count, number}} Zahlen-, Datums-, Textlängen- oder benutzerdefinierte Regel prüft eingegebene Werte wie Excel',
  'core.finding.dataValidation.rules_other':
    '{{count, number}} Zahlen-, Datums-, Textlängen- oder benutzerdefinierte Regeln prüfen eingegebene Werte wie Excel',
  'core.finding.dataValidation.other_one': '{{count, number}} weitere Regel bleibt in der Datei erhalten, wird hier aber nicht angewendet',
  'core.finding.dataValidation.other_other': '{{count, number}} weitere Regeln bleiben in der Datei erhalten, werden hier aber nicht angewendet',
  'core.finding.hyperlinks.title': 'Hyperlinks',
  'core.finding.hyperlinks.detail_one': '{{count, number}} Link. Bleibt in der Datei erhalten; hier nicht anklickbar.',
  'core.finding.hyperlinks.detail_other': '{{count, number}} Links. Bleiben in der Datei erhalten; hier nicht anklickbar.',
  'core.finding.sparklines.title': 'Sparklines',
  'core.finding.sparklines.detail_one':
    '{{count, number}} Sparklinegruppe. Bleibt in der Datei erhalten und wird von Excel neu gezeichnet; wird hier nicht angezeigt.',
  'core.finding.sparklines.detail_other':
    '{{count, number}} Sparklinegruppen. Bleiben in der Datei erhalten und werden von Excel neu gezeichnet; werden hier nicht angezeigt.',
  'core.finding.filter.title': 'Filter',
  'core.finding.filter.detail':
    'Bleibt in der Datei erhalten. Zeilen, die er ausblendet, bleiben hier ausgeblendet. Filtern in diesem Editor ändert nur die Ansicht; der Filter der Datei wird unverändert gespeichert.',
  'core.finding.richText.title': 'Gemischte Formatierung in Zellen',
  'core.finding.richText.detail_one':
    '{{count, number}} Zelle verwendet mehr als einen Schriftschnitt in ihrem Text. Der Text wird ohne diese Formatierung angezeigt; wird eine solche Zelle bearbeitet, bleibt ihr Text erhalten, die gemischte Formatierung aber nicht.',
  'core.finding.richText.detail_other':
    '{{count, number}} Zellen verwenden mehr als einen Schriftschnitt in ihrem Text. Der Text wird ohne diese Formatierung angezeigt; wird eine solche Zelle bearbeitet, bleibt ihr Text erhalten, die gemischte Formatierung aber nicht.',
  'core.finding.pivot.title': 'PivotTable „{{name}}“',
  'core.finding.pivot.detail':
    'Erstellt aus {{source}}. Bleibt in der Datei erhalten und wird beim Öffnen in Excel aktualisiert, sodass Ihre Änderungen übernommen werden. Ihre Zellen können hier nicht bearbeitet werden.',
  'core.finding.table.title': 'Tabelle „{{name}}“',
  'core.finding.table.detail':
    'Bleibt in der Datei erhalten. Ihre Spaltennamen (Überschriftenzeile) können hier nicht bearbeitet werden, und direkt darunter eingegebene Zeilen werden der Tabelle nicht hinzugefügt.',
  'core.finding.slicers.title': 'Datenschnitte',
  'core.finding.timelines.title': 'Zeitachsen',
  'core.finding.slicers.detail': 'Bleiben in der Datei erhalten; werden in diesem Editor nicht angezeigt.',
  'core.finding.comments.title': 'Notizen und Kommentare',
  'core.finding.comments.detail':
    '{{what}}. In der Ecke ihrer Zellen markiert und angezeigt, wenn Sie auf die Zelle zeigen oder sie auswählen. Bleiben in der Datei erhalten; sie können hier nicht hinzugefügt oder geändert werden.',
  'core.finding.comments.notes_one': '{{count, number}} Notiz',
  'core.finding.comments.notes_other': '{{count, number}} Notizen',
  'core.finding.comments.threads_one': '{{count, number}} Kommentarthread',
  'core.finding.comments.threads_other': '{{count, number}} Kommentarthreads',
  'core.finding.comments.notesAndThreads': '{{notes}} und {{threads}}',
  'core.finding.chart.title': 'Diagramm',
  'core.finding.chart.titleNamed': 'Diagramm „{{title}}“',
  'core.finding.chart.detail':
    'Diagramm. Bleibt in der Datei erhalten und wird beim Öffnen der Datei von Excel aus den bearbeiteten Daten neu gezeichnet.',
  'core.finding.chart.detailTyped':
    'Diagramm vom Typ „{{type}}“. Bleibt in der Datei erhalten und wird beim Öffnen der Datei von Excel aus den bearbeiteten Daten neu gezeichnet.',
  'core.finding.pictures.title': 'Bilder',
  'core.finding.pictures.detail_one': '{{count, number}} Bild. Bleibt in der Datei erhalten; wird hier als Platzhalter angezeigt.',
  'core.finding.pictures.detail_other': '{{count, number}} Bilder. Bleiben in der Datei erhalten; werden hier als Platzhalter angezeigt.',
  'core.finding.shapes.title': 'Formen und Textfelder',
  'core.finding.shapes.detail_one': '{{count, number}} Form. Bleibt in der Datei erhalten; wird hier als Platzhalter angezeigt.',
  'core.finding.shapes.detail_other': '{{count, number}} Formen. Bleiben in der Datei erhalten; werden hier als Platzhalter angezeigt.',
  'core.finding.controls.title': 'Formularsteuerelemente',
  'core.finding.controls.detail_one':
    '{{count, number}} Schaltfläche oder Formularsteuerelement. Bleibt in der Datei erhalten; wird hier weder angezeigt noch kann es verwendet werden.',
  'core.finding.controls.detail_other':
    '{{count, number}} Schaltflächen oder Formularsteuerelemente. Bleiben in der Datei erhalten; werden hier weder angezeigt noch können sie verwendet werden.',
  'core.finding.macros.title': 'Makros (VBA)',
  'core.finding.activeContent.title': 'ActiveX-Steuerelemente und eingebettete Objekte',
  'core.finding.activeContent.detail':
    'ActiveX-Steuerelemente: {{controls, number}}. Eingebettete Objekte: {{objects, number}}. Sie bleiben in der Datei erhalten und können hier nicht ausgeführt werden, beim Öffnen in Excel aber schon.',
  'core.finding.macros.detail': 'Diese Arbeitsmappe enthält Makros. Sie bleiben in der Datei erhalten, können hier aber nicht ausgeführt werden.',
  'core.finding.macros.removedDetail':
    'Diese Arbeitsmappe enthält Makros. Sie werden beim Speichern entfernt, und die Datei wird als .xlsx gespeichert.',
  'core.finding.externalLinks.title': 'Verknüpfungen mit anderen Arbeitsmappen',
  'core.finding.externalLinks.detail_one':
    '{{count, number}} verknüpfte Arbeitsmappe. Bleibt in der Datei erhalten. Formeln, die sie verwenden, zeigen den zuletzt von Excel berechneten Wert und werden hier nicht neu berechnet.',
  'core.finding.externalLinks.detail_other':
    '{{count, number}} verknüpfte Arbeitsmappen. Bleiben in der Datei erhalten. Formeln, die sie verwenden, zeigen den zuletzt von Excel berechneten Wert und werden hier nicht neu berechnet.',
  'core.finding.connections.title': 'Datenverbindungen',
  'core.finding.connections.detail': 'Externe Datenverbindungen oder Power Query. Bleiben in der Datei erhalten; werden hier nicht aktualisiert.',
  'core.finding.cellImages.title': 'Bilder oder Datentypen in Zellen',
  'core.finding.cellImages.detail':
    'Einige Zellen enthalten Bilder oder verknüpfte Datentypen. Bleiben in der Datei erhalten; hier wird ihr Ersatzwert angezeigt.',
  'core.finding.workbookProtection.title': 'Die Arbeitsmappenstruktur ist geschützt',
  'core.finding.workbookProtection.detail':
    'Bleibt in der Datei erhalten. Dieser Editor fügt ohnehin keine Blätter hinzu, entfernt sie nicht und benennt sie nicht um.',
  'core.finding.notRecalculated.title': 'Formeln, die hier nicht neu berechnet werden',
  'core.finding.notRecalculated.detail':
    '{{unsupported, number}} von {{total, number}} Formeln enthalten Elemente, die dieser Editor nicht berechnen kann: {{reasons}}. Sie zeigen den zuletzt von Excel berechneten Wert und sind mit einer grauen Ecke markiert; ändern sich ihre Eingaben, werden sie orange. Excel berechnet sie beim Öffnen der Datei neu.',
  'core.finding.calculatedDifferently.title': 'Formeln, die hier anders berechnet werden',
  'core.finding.calculatedDifferently.detail':
    '{{mismatched, number}} Formeln ergaben hier ein anderes Ergebnis als den von Excel gespeicherten Wert. Daher zeigt dieser Editor den Wert von Excel an und berechnet sie nicht neu. Excel berechnet sie beim Öffnen der Datei neu.',
  'core.finding.recalcOff.title': 'Die Live-Neuberechnung ist aus',
  'core.finding.recalcOff.detail':
    'Diese Arbeitsmappe hat {{formulas, number}} Formeln in {{cells, number}} Zellen, mehr als dieser Editor ohne Verzögerung live neu berechnet (das Erstellen der Berechnungsengine würde lange dauern und viel Speicher benötigen). Formeln zeigen die zuletzt von Excel gespeicherten Werte, und von Ihnen eingegebene Formeln werden als nicht berechnet markiert; Excel berechnet beim Öffnen der Datei alles neu. Sie können die Live-Neuberechnung in der Statusleiste einschalten.',

  // Excel's German names for kinds of chart.
  'core.chartType.bar': 'Balken',
  'core.chartType.bar3D': '3D-Balken',
  'core.chartType.column': 'Säule',
  'core.chartType.line': 'Linie',
  'core.chartType.line3D': '3D-Linie',
  'core.chartType.pie': 'Kreis',
  'core.chartType.pie3D': '3D-Kreis',
  'core.chartType.doughnut': 'Ring',
  'core.chartType.area': 'Fläche',
  'core.chartType.area3D': '3D-Fläche',
  'core.chartType.scatter': 'Punkt (X Y)',
  'core.chartType.bubble': 'Blase',
  'core.chartType.radar': 'Netz',
  'core.chartType.stock': 'Kurs',
  'core.chartType.surface': 'Oberfläche',
  'core.chartType.surface3D': '3D-Oberfläche',
  'core.chartType.pieOfPie': 'Kreis aus Kreis',

  'core.pivot.defaultName': 'PivotTable',
  'core.pivot.externalSource': 'einer externen Quelle',
  'core.pivot.otherRange': 'einem anderen Bereich',

  'core.reason.protected': 'Dieses Tabellenblatt ist in Excel geschützt und daher hier schreibgeschützt.',
  'core.reason.locked': 'Diese Zelle ist gesperrt: Ihr Tabellenblatt ist in Excel geschützt.',
  'core.reason.notWorksheet': 'Dieses Blatt kann hier nicht bearbeitet werden.',
  'core.reason.arrayFormula': 'Diese Zelle gehört zu einer Matrixformel. Bearbeiten Sie sie in Excel.',
  'core.reason.pivot': 'Diese Zelle gehört zur PivotTable „{{name}}“. Ändern Sie die PivotTable in Excel.',
  'core.reason.tableHeader':
    'Dies ist ein Spaltenname der Tabelle „{{name}}“. Excel speichert die Spaltennamen von Tabellen separat; benennen Sie die Spalte daher in Excel um.',

  'core.unsupported.function': 'Funktion {{name}}',
  'core.unsupported.tableRefs': 'Tabellenbezüge (Tabelle[Spalte])',
  'core.unsupported.externalLinks': 'Verknüpfungen mit anderen Arbeitsmappen',
  'core.unsupported.threeD': '3D-Bezüge (Tabelle1:Tabelle3!A1)',
  'core.unsupported.syntax': 'Formelsyntax',
  'core.unsupported.array': 'Mehrzellige Matrixformeln',
  'core.unsupported.sharedFormula': 'Gemeinsame Formel ohne ihre Ausgangszelle',
  'core.unsupported.tooLarge': 'Bereiche oder Matrizen, die zu groß zum Berechnen sind',
  'core.calc.reasonCount': '{{reason}} ({{n, number}})',
  'core.calc.mismatch': 'Anders berechnet als in Excel',

  'core.cellReason.function': 'Verwendet die Funktion {{name}}, die dieser Editor nicht berechnen kann.',
  'core.cellReason.tableRefs': 'Verwendet Tabellenbezüge (Tabelle[Spalte]), die dieser Editor nicht berechnen kann.',
  'core.cellReason.externalLinks': 'Verwendet Verknüpfungen mit anderen Arbeitsmappen, die dieser Editor nicht berechnen kann.',
  'core.cellReason.threeD': 'Verwendet 3D-Bezüge (Tabelle1:Tabelle3!A1), die dieser Editor nicht berechnen kann.',
  'core.cellReason.syntax': 'Verwendet eine Formelsyntax, die dieser Editor nicht lesen kann.',
  'core.cellReason.array': 'Eine mehrzellige Matrix- oder Überlaufformel, die dieser Editor nicht berechnen kann.',
  'core.cellReason.sharedFormula': 'Teil einer beschädigten gemeinsamen Formel; Excel repariert sie beim Öffnen der Datei.',
  'core.cellReason.tooLarge': 'Verwendet einen Bereich oder eine Matrix, die für diesen Editor zu groß zum Berechnen ist.',
  'core.cellReason.mismatch':
    'Dieser Editor hat {{engine}} berechnet, Excel hat aber {{excel}} gespeichert; daher wird der Wert von Excel beibehalten.',
  'core.value.empty': '(leer)',

  'core.list.items': '{{first}}, {{rest}}',
  'core.list.clauses': '{{first}}; {{rest}}',

  'core.edit.badFormula': 'Diese Formel enthält einen Fehler: {{problem}}',
  'core.formula.empty': 'Die Formel ist leer.',
  'core.edit.pastedAsText_one':
    '{{count, number}} eingefügter Wert sah wie eine Formel aus, war aber ungültig und wurde daher als Text eingefügt.',
  'core.edit.pastedAsText_other':
    '{{count, number}} eingefügte Werte sahen wie Formeln aus, waren aber ungültig und wurden daher als Text eingefügt.',
  'core.edit.copiedAsValues_one':
    '{{count, number}} Formel konnte nicht an die neue Position angepasst werden und wurde als Wert eingefügt.',
  'core.edit.copiedAsValues_other':
    '{{count, number}} Formeln konnten nicht an die neue Position angepasst werden und wurden als Werte eingefügt.',
  'core.edit.sortedAsValues_one': '{{count, number}} Formel konnte nicht an ihre neue Zeile angepasst werden und wurde als Wert beibehalten.',
  'core.edit.sortedAsValues_other':
    '{{count, number}} Formeln konnten nicht an ihre neuen Zeilen angepasst werden und wurden als Werte beibehalten.',
  'core.edit.filledAsValues_one': '{{count, number}} Formel konnte nicht angepasst werden und wurde als Wert ausgefüllt.',
  'core.edit.filledAsValues_other': '{{count, number}} Formeln konnten nicht angepasst werden und wurden als Werte ausgefüllt.',
  'core.edit.sortMerged':
    'Beim Sortieren würde ein Teil der verbundenen Zelle {{range}} verschoben. Heben Sie die Zellverbindung auf, oder sortieren Sie nur außerhalb davon.',

  'core.structure.notWorksheet': 'Nur Tabellenblätter können so geändert werden.',
  'core.structure.moveNotWorksheet': 'Zellen können nur zwischen Tabellenblättern verschoben werden.',
  'core.structure.sheetProtected': '{{sheet}} ist in Excel geschützt und daher hier schreibgeschützt.',
  'core.structure.noRoomBelow': 'Kein Platz: Durch das Einfügen würden Zellen mit Daten über das untere Ende des Blatts hinaus verschoben.',
  'core.structure.noRoomRight': 'Kein Platz: Durch das Einfügen würden Zellen mit Daten über den rechten Rand des Blatts hinaus verschoben.',
  'core.structure.doesNotFit': 'Die Zellen passen dort nicht hin.',
  'core.structure.splitsArray': 'Damit würde ein Teil der Matrixformel in {{range}} geändert. Ändern Sie die ganze Matrix oder nichts davon.',
  'core.structure.splitsPivot': 'Damit würde ein Teil der PivotTable in {{range}} geändert. Ändern Sie die PivotTable in Excel.',
  'core.structure.deletesTableHeader':
    'Damit würde die Überschriftenzeile der Tabelle „{{table}}“ gelöscht, aber nicht die Tabelle. Löschen Sie die ganze Tabelle, oder behalten Sie ihre Überschrift.',
  'core.structure.deletesTable': 'Damit würde die ganze Tabelle „{{table}}“ gelöscht. Löschen Sie die Tabelle in Excel.',
  'core.structure.splitsTable':
    'Damit würde ein Teil der Tabelle „{{table}}“ verschoben. Verschieben Sie die ganze Tabelle oder nur Zellen außerhalb davon.',
  'core.structure.splitsMerge': 'Damit würde ein Teil der verbundenen Zelle {{range}} verschoben. Verschieben Sie sie ganz.',
  'core.structure.mergeToOtherSheet': 'Verbundene Zellen können nur innerhalb ihres Blatts verschoben werden. Kopieren Sie sie stattdessen.',
  'core.structure.tableToOtherSheet': 'Tabellen können nur innerhalb ihres Blatts verschoben werden.',
  'core.structure.brokenRefs_one':
    '{{count, number}} Formel zeigt jetzt {{error}}, da sie Zellen verwendet hat, die gelöscht oder überschrieben wurden.',
  'core.structure.brokenRefs_other':
    '{{count, number}} Formeln zeigen jetzt {{error}}, da sie Zellen verwendet haben, die gelöscht oder überschrieben wurden.',
  'core.structure.cantUndo': 'Diese Änderung kann nicht mehr rückgängig gemacht werden.',
  'core.structure.nothingToRedo': 'Es gibt nichts zu wiederholen.',

  'core.change.movedTo': 'Verschoben nach {{where}}',
  'core.change.insertedRows_one': '{{count, number}} Zeile eingefügt',
  'core.change.insertedRows_other': '{{count, number}} Zeilen eingefügt',
  'core.change.insertedColumns_one': '{{count, number}} Spalte eingefügt',
  'core.change.insertedColumns_other': '{{count, number}} Spalten eingefügt',
  'core.change.deletedRows_one': '{{count, number}} Zeile gelöscht',
  'core.change.deletedRows_other': '{{count, number}} Zeilen gelöscht',
  'core.change.deletedColumns_one': '{{count, number}} Spalte gelöscht',
  'core.change.deletedColumns_other': '{{count, number}} Spalten gelöscht',

  'core.validation.kind.whole': 'Ganze Zahl',
  'core.validation.kind.decimal': 'Zahl',
  'core.validation.kind.date': 'Datum',
  'core.validation.kind.time': 'Uhrzeit',
  'core.validation.kind.textLength': 'Textlänge',
  'core.validation.between': '{{what}} zwischen {{a}} und {{b}}',
  'core.validation.notBetween': '{{what}} nicht zwischen {{a}} und {{b}}',
  'core.validation.equal': '{{what}} gleich {{a}}',
  'core.validation.notEqual': '{{what}} ungleich {{a}}',
  'core.validation.greaterThan': '{{what}} größer als {{a}}',
  'core.validation.lessThan': '{{what}} kleiner als {{a}}',
  'core.validation.greaterThanOrEqual': '{{what}} mindestens {{a}}',
  'core.validation.lessThanOrEqual': '{{what}} höchstens {{a}}',
  'core.validation.custom': 'Excel prüft diese Zelle mit ={{formula}}',
};
