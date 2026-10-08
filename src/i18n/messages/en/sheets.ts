/** English messages for adding, renaming, deleting, moving, hiding and unhiding sheets. */
export const sheets = {
  /** What a new sheet is called before its number, as Excel in this language calls it ("Sheet2"). */
  'sheets.defaultName': 'Sheet',

  'sheets.name.blank': 'Type a name for the sheet.',
  'sheets.name.tooLong': 'A sheet name can be at most {{max, number}} characters long.',
  'sheets.name.characters': "A sheet name can't contain : \\ / ? * [ or ], or a line break.",
  'sheets.name.apostrophe': "A sheet name can't start or end with an apostrophe (').",
  'sheets.name.reserved': 'Excel keeps the name “History” for itself.',
  'sheets.name.taken': "There's already a sheet called “{{name}}” (capital letters don't make a name different).",

  'sheets.refuse.protected': "The workbook's structure is protected, so sheets can't be added, deleted, renamed, moved, hidden or unhidden.",
  'sheets.refuse.noSheet': "That sheet isn't in the workbook any more.",
  'sheets.refuse.deleteLast': "A workbook must keep at least one visible sheet, so this one can't be deleted.",
  'sheets.refuse.hideLast': "A workbook must keep at least one visible sheet, so this one can't be hidden.",
  'sheets.refuse.veryHidden': "“{{sheet}}” is hidden so that only Excel's Visual Basic editor can show it again.",
  'sheets.refuse.slicers': '“{{sheet}}” has pivot tables or tables that slicers or timelines use. Delete those in Excel first.',
  /** {{expected}} is schema.match.text or schema.match.pattern. */
  'sheets.refuse.schemaRename': 'The “{{schema}}” schema finds this sheet by its name, which has to stay {{expected}}.',
  'sheets.refuse.schemaDelete': "The “{{schema}}” schema checks this sheet, so it can't be deleted.",

  /** Each change to a sheet, in the list of changes since the file was opened. */
  'sheets.change.added': 'Sheet added',
  'sheets.change.renamed': 'Sheet renamed from “{{from}}”',
  'sheets.change.deleted': 'Sheet deleted',
  'sheets.change.moved': 'Sheet moved to position {{position, number}}',
  'sheets.change.hidden': 'Sheet hidden',
  'sheets.change.unhidden': 'Sheet unhidden',

  /** What was done, in the history and on the undo and redo buttons. */
  'sheets.history.added': 'Added sheet “{{sheet}}”',
  'sheets.history.renamed': 'Renamed sheet “{{from}}” to “{{to}}”',
  'sheets.history.deleted': 'Deleted sheet “{{sheet}}”',
  'sheets.history.moved': 'Moved sheet “{{sheet}}”',
  'sheets.history.hidden': 'Hid sheet “{{sheet}}”',
  'sheets.history.unhidden': 'Unhid sheet “{{sheet}}”',

  'sheets.delete.title': 'Delete the sheet “{{sheet}}”?',
  'sheets.delete.body': 'Its cells go, with any charts, tables and pivot tables on it, and formulas that use it will show #REF!. You can undo this.',
  'sheets.delete.confirm': 'Delete sheet',

  'sheets.add': 'New sheet',
  'sheets.addTitle': 'Add a sheet after this one',
  'sheets.unhideButtonTitle': 'Show a hidden sheet again',
  'sheets.menu.label': 'Sheet “{{sheet}}”',
  'sheets.menu.insert': 'Insert sheet',
  'sheets.menu.rename': 'Rename…',
  'sheets.menu.delete': 'Delete',
  'sheets.menu.moveLeft': 'Move left',
  'sheets.menu.moveRight': 'Move right',
  'sheets.menu.hide': 'Hide',
  'sheets.menu.unhide': 'Unhide…',

  'sheets.rename.title': 'Rename “{{sheet}}”',
  'sheets.rename.label': 'Sheet name',
  'sheets.rename.hint': 'Up to {{max, number}} characters, without : \\ / ? * [ or ].',
  'sheets.rename.save': 'Rename',

  'sheets.unhide.title': 'Unhide a sheet',
  'sheets.unhide.label': 'Hidden sheets',
  'sheets.unhide.confirm': 'Unhide',

  'sheets.shortcuts.group': 'Sheet tabs',
  'sheets.keys.move': 'Alt+Shift+Left, Alt+Shift+Right',
  'sheets.shortcuts.move': 'Move the sheet whose tab has the focus left or right',
  'sheets.keys.rename': 'F2',
  'sheets.shortcuts.rename': 'Rename the sheet whose tab has the focus',
  'sheets.shortcuts.menu': 'Sheet menu: insert, rename, delete, move, hide and unhide sheets',
} satisfies Record<string, string>;
