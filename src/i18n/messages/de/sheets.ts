import type { sheets as en } from '../en/sheets';

/** German messages for adding, renaming, deleting, moving, hiding and unhiding sheets; every English key needs one (the type checks it). */
export const sheets: Record<keyof typeof en, string> = {
  'sheets.defaultName': 'Tabelle',

  'sheets.name.blank': 'Geben Sie einen Namen für das Blatt ein.',
  'sheets.name.tooLong': 'Ein Blattname darf höchstens {{max, number}} Zeichen lang sein.',
  'sheets.name.characters': 'Ein Blattname darf weder : \\ / ? * [ noch ] und keinen Zeilenumbruch enthalten.',
  'sheets.name.apostrophe': 'Ein Blattname darf nicht mit einem Apostroph (\') beginnen oder enden.',
  'sheets.name.reserved': 'Den Namen „History“ behält Excel sich selbst vor.',
  'sheets.name.taken': 'Es gibt schon ein Blatt namens „{{name}}“ (Groß- und Kleinschreibung zählt dabei nicht).',

  'sheets.refuse.protected': 'Die Struktur der Arbeitsmappe ist geschützt, daher können Blätter nicht eingefügt, gelöscht, umbenannt, verschoben, aus- oder eingeblendet werden.',
  'sheets.refuse.noSheet': 'Dieses Blatt ist nicht mehr in der Arbeitsmappe.',
  'sheets.refuse.deleteLast': 'Eine Arbeitsmappe muss mindestens ein sichtbares Blatt behalten, daher kann dieses nicht gelöscht werden.',
  'sheets.refuse.hideLast': 'Eine Arbeitsmappe muss mindestens ein sichtbares Blatt behalten, daher kann dieses nicht ausgeblendet werden.',
  'sheets.refuse.veryHidden': '„{{sheet}}“ ist so ausgeblendet, dass nur der Visual-Basic-Editor von Excel es wieder einblenden kann.',
  'sheets.refuse.slicers': '„{{sheet}}“ enthält PivotTables oder Tabellen, die Datenschnitte oder Zeitachsen verwenden. Löschen Sie diese zuerst in Excel.',
  'sheets.refuse.schemaRename': 'Das Schema „{{schema}}“ findet dieses Blatt über seinen Namen, der {{expected}} bleiben muss.',
  'sheets.refuse.schemaDelete': 'Das Schema „{{schema}}“ prüft dieses Blatt, daher kann es nicht gelöscht werden.',

  'sheets.change.added': 'Blatt eingefügt',
  'sheets.change.renamed': 'Blatt umbenannt, vorher „{{from}}“',
  'sheets.change.deleted': 'Blatt gelöscht',
  'sheets.change.moved': 'Blatt an Position {{position, number}} verschoben',
  'sheets.change.hidden': 'Blatt ausgeblendet',
  'sheets.change.unhidden': 'Blatt eingeblendet',

  'sheets.history.added': 'Blatt „{{sheet}}“ eingefügt',
  'sheets.history.renamed': 'Blatt „{{from}}“ in „{{to}}“ umbenannt',
  'sheets.history.deleted': 'Blatt „{{sheet}}“ gelöscht',
  'sheets.history.moved': 'Blatt „{{sheet}}“ verschoben',
  'sheets.history.hidden': 'Blatt „{{sheet}}“ ausgeblendet',
  'sheets.history.unhidden': 'Blatt „{{sheet}}“ eingeblendet',

  'sheets.delete.title': 'Blatt „{{sheet}}“ löschen?',
  'sheets.delete.body': 'Seine Zellen werden gelöscht, mit allen Diagrammen, Tabellen und PivotTables darauf, und Formeln, die es verwenden, zeigen #REF!. Sie können das rückgängig machen.',
  'sheets.delete.confirm': 'Blatt löschen',

  'sheets.add': 'Neues Blatt',
  'sheets.addTitle': 'Nach diesem Blatt ein neues einfügen',
  'sheets.unhideButtonTitle': 'Ein ausgeblendetes Blatt wieder einblenden',
  'sheets.menu.label': 'Blatt „{{sheet}}“',
  'sheets.menu.insert': 'Blatt einfügen',
  'sheets.menu.rename': 'Umbenennen…',
  'sheets.menu.delete': 'Löschen',
  'sheets.menu.moveLeft': 'Nach links verschieben',
  'sheets.menu.moveRight': 'Nach rechts verschieben',
  'sheets.menu.hide': 'Ausblenden',
  'sheets.menu.unhide': 'Einblenden…',

  'sheets.rename.title': '„{{sheet}}“ umbenennen',
  'sheets.rename.label': 'Blattname',
  'sheets.rename.hint': 'Bis zu {{max, number}} Zeichen, ohne : \\ / ? * [ oder ].',
  'sheets.rename.save': 'Umbenennen',

  'sheets.unhide.title': 'Blatt einblenden',
  'sheets.unhide.label': 'Ausgeblendete Blätter',
  'sheets.unhide.confirm': 'Einblenden',

  'sheets.shortcuts.group': 'Blattregister',
  'sheets.keys.move': 'Alt+Umschalt+Links, Alt+Umschalt+Rechts',
  'sheets.shortcuts.move': 'Das Blatt, dessen Register den Fokus hat, nach links oder rechts verschieben',
  'sheets.keys.rename': 'F2',
  'sheets.shortcuts.rename': 'Das Blatt, dessen Register den Fokus hat, umbenennen',
  'sheets.shortcuts.menu': 'Blattmenü: Blätter einfügen, umbenennen, löschen, verschieben, aus- und einblenden',
};
