/** English messages from the worker (findings, reasons, errors people see). Keys are dotted, grouped by where they are shown. */
export const core = {
  // Files the editor refuses to open.
  'core.open.tooBig': 'This file is {{size, number}} MB; the most this editor accepts is {{max, number}} MB.',
  'core.open.oldFormat': 'This file is password-protected or in the old .xls format. Only .xlsx and .xlsm files are supported.',
  'core.open.notWorkbook': 'This file is not an Excel workbook (.xlsx or .xlsm).',
  'core.open.tooManyParts': 'This file has more than {{max, number}} parts, which is more than a workbook should.',
  'core.open.unpacksTooBig': 'This file unpacks to more than {{max, number}} MB, which is more than this editor can hold.',
  'core.open.zipBomb': 'Part {{part}} is compressed suspiciously well ({{ratio, number}}:1), so the file was not opened.',
  'core.open.notPackage': 'This file is not a valid Office Open XML package.',
  'core.open.damaged': "This file is damaged, so it wasn't opened. Excel may be able to repair it.",
  'core.open.noWorkbookPart': 'The workbook part is missing from this file.',
  'core.open.tooManyCells': 'This workbook has more than {{max, number}} cells, the most this editor can hold.',
  'core.open.macrosBlocked': "This workbook contains macros (VBA or Excel 4.0 macro sheets), which aren't allowed here. Save it as .xlsx in Excel and open that instead.",
  'core.open.activeContentBlocked': "This workbook contains ActiveX controls or embedded objects, which aren't allowed here.",

  // "What's in this file": what the workbook holds that the editor keeps but doesn't show.
  'core.finding.hiddenSheet.title': 'Hidden sheet "{{sheet}}"',
  'core.finding.hiddenSheet.detail': 'Kept in the file but not shown, as in Excel.',
  'core.finding.chartsheet.title': 'Chart sheet "{{sheet}}"',
  'core.finding.chartsheet.detail':
    'A chart that fills a whole sheet. Kept in the file and redrawn by Excel from the edited data when the file is opened.',
  'core.finding.chartsheet.detailTitled':
    'A chart titled "{{title}}" that fills a whole sheet. Kept in the file and redrawn by Excel from the edited data when the file is opened.',
  /** {{type}} is a core.chartType message, or Excel's own name for a kind of chart it has no message for. */
  'core.finding.chartsheet.detailTyped':
    '{{type}} chart that fills a whole sheet. Kept in the file and redrawn by Excel from the edited data when the file is opened.',
  'core.finding.chartsheet.detailTypedTitled':
    '{{type}} chart titled "{{title}}" that fills a whole sheet. Kept in the file and redrawn by Excel from the edited data when the file is opened.',
  'core.finding.unsupportedSheet.title': 'Unsupported sheet "{{sheet}}"',
  'core.finding.unsupportedSheet.detail': 'A dialog or macro sheet from older versions of Excel. Kept in the file; not shown in this editor.',
  'core.finding.macroSheet.title': 'Excel 4.0 macro sheet "{{sheet}}"',
  'core.finding.macroSheet.detail': "Holds Excel 4.0 macros, which can run when the file is opened in Excel. Kept in the file; not shown or run here.",
  'core.finding.macroSheet.removedDetail': 'Holds Excel 4.0 macros. The sheet is emptied when you save, and the file is saved as .xlsx.',
  'core.finding.mergedCells.title': 'Merged cells spanning rows',
  'core.finding.mergedCells.detail_one':
    '{{count, number}} merged range span more than one row. Their value is shown in the top-left cell only; the merge is kept in the file.',
  'core.finding.mergedCells.detail_other':
    '{{count, number}} merged ranges span more than one row. Their value is shown in the top-left cell only; the merge is kept in the file.',
  'core.finding.protectedSheet.title': 'Protected sheet',
  'core.finding.protectedSheet.detail':
    "Protected in Excel: only the cells it leaves unlocked can be edited here, and its rows and columns can't be inserted, deleted or sorted.",
  'core.finding.arrayFormulas.title': 'Array formulas',
  'core.finding.arrayFormulas.detail_one':
    "{{count, number}} array or spill formula cover more than one cell. They keep Excel's last values, Excel recalculates them on open, and their cells can't be edited here.",
  'core.finding.arrayFormulas.detail_other':
    "{{count, number}} array or spill formulas cover more than one cell. They keep Excel's last values, Excel recalculates them on open, and their cells can't be edited here.",
  'core.finding.conditionalFormat.title': 'Conditional formatting',
  'core.finding.conditionalFormat.detail_one': '{{count, number}} rule. Shown here and updated as you edit; kept in the file for Excel.',
  'core.finding.conditionalFormat.detail_other': '{{count, number}} rules. Shown here and updated as you edit; kept in the file for Excel.',
  'core.finding.dataValidation.title': 'Data validation',
  /** {{rules}} is one or more of the dataValidation.lists, .rules and .other clauses. */
  'core.finding.dataValidation.detail': "{{rules}}. Pasted values aren't checked, as in Excel.",
  'core.finding.dataValidation.lists_one': '{{count, number}} list rule show a dropdown and check typed values as Excel does',
  'core.finding.dataValidation.lists_other': '{{count, number}} list rules show a dropdown and check typed values as Excel does',
  'core.finding.dataValidation.rules_one': '{{count, number}} number, date, text-length or custom rule check typed values as Excel does',
  'core.finding.dataValidation.rules_other': '{{count, number}} number, date, text-length or custom rules check typed values as Excel does',
  'core.finding.dataValidation.other_one': '{{count, number}} other rule are kept in the file but not enforced here',
  'core.finding.dataValidation.other_other': '{{count, number}} other rules are kept in the file but not enforced here',
  'core.finding.hyperlinks.title': 'Hyperlinks',
  'core.finding.hyperlinks.detail_one': '{{count, number}} link. Kept in the file; not clickable here.',
  'core.finding.hyperlinks.detail_other': '{{count, number}} links. Kept in the file; not clickable here.',
  'core.finding.sparklines.title': 'Sparklines',
  'core.finding.sparklines.detail_one': '{{count, number}} sparkline group. Kept in the file and redrawn by Excel; not shown here.',
  'core.finding.sparklines.detail_other': '{{count, number}} sparkline groups. Kept in the file and redrawn by Excel; not shown here.',
  'core.finding.filter.title': 'Filter',
  'core.finding.filter.detail':
    "Kept in the file. Rows it hides stay hidden here. Filtering in this editor only changes what you see; the file's filter is saved as it was.",
  'core.finding.richText.title': 'Mixed formatting inside cells',
  'core.finding.richText.detail_one':
    '{{count, number}} cell use more than one font style within their text. The text is shown plainly; editing one of those cells keeps its text but not the mixed formatting.',
  'core.finding.richText.detail_other':
    '{{count, number}} cells use more than one font style within their text. The text is shown plainly; editing one of those cells keeps its text but not the mixed formatting.',
  'core.finding.pivot.title': 'Pivot table "{{name}}"',
  /** {{source}} is a range such as Data!A1:D20, or core.pivot.externalSource or core.pivot.otherRange. */
  'core.finding.pivot.detail':
    "Built from {{source}}. Kept in the file and set to refresh when the file is opened in Excel, so it picks up your edits. Its cells can't be edited here.",
  'core.finding.table.title': 'Table "{{name}}"',
  'core.finding.table.detail':
    'Kept in the file. Its column names (header row) can’t be edited here, and rows typed directly below it are not added to the table.',
  'core.finding.slicers.title': 'Slicers',
  'core.finding.timelines.title': 'Timelines',
  'core.finding.slicers.detail': 'Kept in the file; not shown in this editor.',
  'core.finding.comments.title': 'Notes and comments',
  /** {{what}} is comments.notes, comments.threads or comments.notesAndThreads. */
  'core.finding.comments.detail':
    "{{what}}. Marked in the corner of their cells and shown when you point at the cell or select it. Kept in the file; they can't be added or changed here.",
  'core.finding.comments.notes_one': '{{count, number}} note',
  'core.finding.comments.notes_other': '{{count, number}} notes',
  'core.finding.comments.threads_one': '{{count, number}} comment thread',
  'core.finding.comments.threads_other': '{{count, number}} comment threads',
  'core.finding.comments.notesAndThreads': '{{notes}} and {{threads}}',
  'core.finding.chart.title': 'Chart',
  'core.finding.chart.titleNamed': 'Chart "{{title}}"',
  'core.finding.chart.detail':
    'Chart. Kept in the file and redrawn by Excel from the edited data when the file is opened.',
  /** {{type}} is a core.chartType message, or Excel's own name for a kind of chart it has no message for. */
  'core.finding.chart.detailTyped':
    '{{type}} chart. Kept in the file and redrawn by Excel from the edited data when the file is opened.',
  'core.finding.pictures.title': 'Pictures',
  'core.finding.pictures.detail_one': '{{count, number}} picture. Kept in the file; shown here as placeholders.',
  'core.finding.pictures.detail_other': '{{count, number}} pictures. Kept in the file; shown here as placeholders.',
  'core.finding.shapes.title': 'Shapes and text boxes',
  'core.finding.shapes.detail_one': '{{count, number}} shape. Kept in the file; shown here as placeholders.',
  'core.finding.shapes.detail_other': '{{count, number}} shapes. Kept in the file; shown here as placeholders.',
  'core.finding.controls.title': 'Form controls',
  'core.finding.controls.detail_one': '{{count, number}} button or form control. Kept in the file; not shown or usable here.',
  'core.finding.controls.detail_other': '{{count, number}} buttons or form controls. Kept in the file; not shown or usable here.',
  'core.finding.macros.title': 'Macros (VBA)',
  'core.finding.activeContent.title': 'ActiveX controls and embedded objects',
  'core.finding.activeContent.detail':
    "ActiveX controls: {{controls, number}}. Embedded objects: {{objects, number}}. They are kept in the file and can't run here, but could run when it's opened in Excel.",
  'core.finding.macros.detail': 'This workbook contains macros. They are kept in the file but cannot run here.',
  'core.finding.macros.removedDetail': 'This workbook contains macros. They will be removed when you save, and the file is saved as .xlsx.',
  'core.finding.externalLinks.title': 'Links to other workbooks',
  'core.finding.externalLinks.detail_one':
    "{{count, number}} linked workbook. Kept in the file. Formulas that use them show Excel's last value and are not recalculated here.",
  'core.finding.externalLinks.detail_other':
    "{{count, number}} linked workbooks. Kept in the file. Formulas that use them show Excel's last value and are not recalculated here.",
  'core.finding.connections.title': 'Data connections',
  'core.finding.connections.detail': 'External data connections or Power Query. Kept in the file; not refreshed here.',
  'core.finding.cellImages.title': 'Pictures or data types in cells',
  'core.finding.cellImages.detail': 'Some cells hold pictures or linked data types. Kept in the file; shown as their fallback value here.',
  'core.finding.workbookProtection.title': 'Workbook structure is protected',
  'core.finding.workbookProtection.detail': 'Kept in the file. This editor does not add, remove or rename sheets anyway.',
  'core.finding.notRecalculated.title': 'Formulas not recalculated here',
  /** {{reasons}} lists core.calc.reasonCount messages, such as "LET function (3), Formula syntax (1)". */
  'core.finding.notRecalculated.detail':
    "{{unsupported, number}} of {{total, number}} formulas use features this editor can't calculate: {{reasons}}. They show Excel's last value and are marked with a grey corner; if their inputs change they turn amber. Excel recalculates them when the file is opened.",
  'core.finding.calculatedDifferently.title': 'Formulas calculated differently here',
  'core.finding.calculatedDifferently.detail':
    "{{mismatched, number}} formulas gave a different result here than Excel's saved value, so this editor shows Excel's value and does not recalculate them. Excel recalculates them when the file is opened.",
  'core.finding.recalcOff.title': 'Live recalculation is off',
  'core.finding.recalcOff.detail':
    'This workbook has {{formulas, number}} formulas in {{cells, number}} cells, more than this editor recalculates live without slowing down (building the engine would take a while and a lot of memory). Formulas show the values Excel last saved, and formulas you type are marked as not calculated; Excel recalculates everything when the file is opened. You can turn live recalculation on from the status bar.',

  // Kinds of chart (SheetObject.chartType), as Excel names them: shown as "{{type}} chart".
  'core.chartType.bar': 'Bar',
  'core.chartType.bar3D': '3-D bar',
  'core.chartType.column': 'Column',
  'core.chartType.line': 'Line',
  'core.chartType.line3D': '3-D line',
  'core.chartType.pie': 'Pie',
  'core.chartType.pie3D': '3-D pie',
  'core.chartType.doughnut': 'Doughnut',
  'core.chartType.area': 'Area',
  'core.chartType.area3D': '3-D area',
  'core.chartType.scatter': 'Scatter',
  'core.chartType.bubble': 'Bubble',
  'core.chartType.radar': 'Radar',
  'core.chartType.stock': 'Stock',
  'core.chartType.surface': 'Surface',
  'core.chartType.surface3D': '3-D surface',
  'core.chartType.pieOfPie': 'Pie of pie',

  'core.pivot.defaultName': 'Pivot table',
  'core.pivot.externalSource': 'an external source',
  'core.pivot.otherRange': 'another range',

  // Why a cell can't be edited.
  'core.reason.protected': 'This sheet is protected in Excel, so it is read-only here.',
  'core.reason.locked': 'This cell is locked: its sheet is protected in Excel.',
  'core.reason.notWorksheet': 'This sheet cannot be edited here.',
  'core.reason.arrayFormula': 'This cell is part of an array formula. Edit it in Excel.',
  'core.reason.pivot': 'This cell is part of pivot table "{{name}}". Change the pivot table in Excel.',
  'core.reason.tableHeader':
    'This is a column name in table "{{name}}". Excel stores table column names separately, so rename the column in Excel.',

  // What the calculation engine can't calculate, as the recalculation report lists it.
  'core.unsupported.function': '{{name}} function',
  'core.unsupported.tableRefs': 'Table references (Table[Column])',
  'core.unsupported.externalLinks': 'Links to other workbooks',
  'core.unsupported.threeD': '3-D references (Sheet1:Sheet3!A1)',
  'core.unsupported.syntax': 'Formula syntax',
  'core.unsupported.array': 'Multi-cell array formulas',
  'core.unsupported.sharedFormula': 'Shared formula without its master cell',
  'core.unsupported.tooLarge': 'Ranges or arrays too large to calculate here',
  /** One entry in the list of what isn't calculated: {{reason}} is a core.unsupported message, {{n}} how many formulas. */
  'core.calc.reasonCount': '{{reason}} ({{n, number}})',
  'core.calc.mismatch': 'Calculated differently from Excel',

  // Why a formula cell isn't calculated, shown when it is selected.
  'core.cellReason.function': "Uses the {{name}} function, which this editor can't calculate.",
  'core.cellReason.tableRefs': "Uses table references (Table[Column]), which this editor can't calculate.",
  'core.cellReason.externalLinks': "Uses links to other workbooks, which this editor can't calculate.",
  'core.cellReason.threeD': "Uses 3-D references (Sheet1:Sheet3!A1), which this editor can't calculate.",
  'core.cellReason.syntax': "Uses formula syntax this editor can't read.",
  'core.cellReason.array': "A multi-cell array or spill formula, which this editor can't calculate.",
  'core.cellReason.sharedFormula': 'Part of a damaged shared formula; Excel will repair it when the file is opened.',
  'core.cellReason.tooLarge': "Uses a range or an array too large for this editor to calculate.",
  /** {{engine}} and {{excel}} are values: a number, "text" in quotes, an error such as #N/A, or core.value.empty. */
  'core.cellReason.mismatch': "This editor calculated {{engine}} but Excel saved {{excel}}, so Excel's value is kept.",
  'core.value.empty': '(empty)',

  // Lists of messages: {{first}} is the first item, {{rest}} the rest of the list.
  'core.list.items': '{{first}}, {{rest}}',
  'core.list.clauses': '{{first}}; {{rest}}',

  // Edits.
  /** {{problem}} is core.formula.empty or the formula parser's own (English) description. */
  'core.edit.badFormula': "There's a problem with this formula: {{problem}}",
  'core.formula.empty': 'The formula is empty.',
  'core.edit.pastedAsText_one': "{{count, number}} pasted value looked like a formula but wasn't valid, so it was pasted as text.",
  'core.edit.pastedAsText_other': "{{count, number}} pasted values looked like a formula but wasn't valid, so they were pasted as text.",
  'core.edit.copiedAsValues_one': "{{count, number}} formula couldn't be adjusted for the new position and was pasted as values.",
  'core.edit.copiedAsValues_other': "{{count, number}} formulas couldn't be adjusted for the new position and were pasted as values.",
  'core.edit.sortedAsValues_one': "{{count, number}} formula couldn't be adjusted for its new row and was kept as values.",
  'core.edit.sortedAsValues_other': "{{count, number}} formulas couldn't be adjusted for their new row and were kept as values.",
  'core.edit.filledAsValues_one': "{{count, number}} formula couldn't be adjusted and was filled as values.",
  'core.edit.filledAsValues_other': "{{count, number}} formulas couldn't be adjusted and were filled as values.",
  'core.edit.sortMerged': 'Sorting would move part of the merged cell {{range}}. Unmerge it, or sort around it.',

  // Inserting and deleting rows and columns, and moving cells.
  'core.structure.notWorksheet': 'Only worksheets can be changed like this.',
  'core.structure.moveNotWorksheet': 'Cells can only be moved between worksheets.',
  'core.structure.sheetProtected': '{{sheet}} is protected in Excel, so it is read-only here.',
  'core.structure.noRoomBelow': "There isn't room: inserting would push cells with data off the bottom of the sheet.",
  'core.structure.noRoomRight': "There isn't room: inserting would push cells with data off the right of the sheet.",
  'core.structure.doesNotFit': "The cells don't fit there.",
  'core.structure.splitsArray': 'That would change part of the array formula in {{range}}. Change the whole array, or none of it.',
  'core.structure.splitsPivot': 'That would change part of the pivot table at {{range}}. Change the pivot table in Excel.',
  'core.structure.deletesTableHeader':
    'That would delete the header row of table "{{table}}" but not the table. Delete the whole table, or leave its header.',
  'core.structure.deletesTable': 'That would delete all of table "{{table}}". Delete the table in Excel.',
  'core.structure.splitsTable': 'That would move part of table "{{table}}". Move the whole table, or cells outside it.',
  'core.structure.splitsMerge': 'That would move part of the merged cell {{range}}. Move all of it.',
  'core.structure.mergeToOtherSheet': 'Merged cells can only be moved within their sheet. Copy them instead.',
  'core.structure.tableToOtherSheet': 'Tables can only be moved within their sheet.',
  /** {{error}} is Excel's error code, #REF!. */
  'core.structure.brokenRefs_one': '{{count, number}} formula now shows {{error}} because it used cells that were deleted or pasted over.',
  'core.structure.brokenRefs_other': '{{count, number}} formulas now show {{error}} because they used cells that were deleted or pasted over.',
  'core.structure.cantUndo': "That change can't be undone any more.",
  'core.structure.nothingToRedo': 'There is nothing to redo.',

  // The change list's description of a structural change.
  'core.change.movedTo': 'Moved to {{where}}',
  'core.layout.mergeTable': "Cells in an Excel table can't be merged ({{table}}).",
  'core.change.columnWidths_one': 'Changed the width of {{count, number}} column',
  'core.change.columnWidths_other': 'Changed the width of {{count, number}} columns',
  'core.change.rowHeights_one': 'Changed the height of {{count, number}} row',
  'core.change.rowHeights_other': 'Changed the height of {{count, number}} rows',
  'core.change.hidRows_one': 'Hid {{count, number}} row',
  'core.change.hidRows_other': 'Hid {{count, number}} rows',
  'core.change.hidColumns_one': 'Hid {{count, number}} column',
  'core.change.hidColumns_other': 'Hid {{count, number}} columns',
  'core.change.shownRows_one': 'Unhid {{count, number}} row',
  'core.change.shownRows_other': 'Unhid {{count, number}} rows',
  'core.change.shownColumns_one': 'Unhid {{count, number}} column',
  'core.change.shownColumns_other': 'Unhid {{count, number}} columns',
  'core.change.froze': 'Froze the rows above and the columns left of {{cell}}',
  'core.change.unfroze': 'Unfroze the panes',
  'core.change.merged': 'Merged the cells',
  'core.change.unmerged': 'Unmerged the cells',
  'core.change.tableGrew_one': 'Table {{table}} took in {{count, number}} more row',
  'core.change.tableGrew_other': 'Table {{table}} took in {{count, number}} more rows',
  'core.change.insertedRows_one': 'Inserted {{count, number}} row',
  'core.change.insertedRows_other': 'Inserted {{count, number}} rows',
  'core.change.insertedColumns_one': 'Inserted {{count, number}} column',
  'core.change.insertedColumns_other': 'Inserted {{count, number}} columns',
  'core.change.deletedRows_one': 'Deleted {{count, number}} row',
  'core.change.deletedRows_other': 'Deleted {{count, number}} rows',
  'core.change.deletedColumns_one': 'Deleted {{count, number}} column',
  'core.change.deletedColumns_other': 'Deleted {{count, number}} columns',

  // An Excel data validation rule, described in the formula bar. {{a}} and {{b}} are the rule's bounds as Excel shows them.
  'core.validation.kind.whole': 'Whole number',
  'core.validation.kind.decimal': 'Number',
  'core.validation.kind.date': 'Date',
  'core.validation.kind.time': 'Time',
  'core.validation.kind.textLength': 'Text length',
  /** {{what}} is a core.validation.kind message. */
  'core.validation.between': '{{what}} between {{a}} and {{b}}',
  'core.validation.notBetween': '{{what}} not between {{a}} and {{b}}',
  'core.validation.equal': '{{what}} equal to {{a}}',
  'core.validation.notEqual': '{{what}} not equal to {{a}}',
  'core.validation.greaterThan': '{{what}} greater than {{a}}',
  'core.validation.lessThan': '{{what}} less than {{a}}',
  'core.validation.greaterThanOrEqual': '{{what}} at least {{a}}',
  'core.validation.lessThanOrEqual': '{{what}} at most {{a}}',
  'core.validation.custom': 'Excel checks this cell with ={{formula}}',
} satisfies Record<string, string>;
