/**
 * English messages for defined names: the name box, the dialogs that define
 * and manage names, and what the worker says about them.
 */
export const names = {
  'names.error.empty': 'Type a name.',
  'names.error.tooLong': 'A name can be at most {{max, number}} characters long.',
  'names.error.start': 'A name must start with a letter, an underscore (_) or a backslash (\\).',
  'names.error.space': "A name can't contain spaces; use an underscore (_) or a period (.) instead.",
  'names.error.character': 'A name can only contain letters, digits, periods (.) and underscores (_), not "{{character}}".',
  'names.error.reference': '"{{name}}" is a cell reference, so it can\'t be a name.',
  'names.error.reserved': '"{{name}}" is reserved by Excel, so it can\'t be a name.',
  'names.error.takenWorkbook': 'The workbook already has a name "{{name}}".',
  'names.error.takenSheet': 'Sheet {{sheet}} already has a name "{{name}}".',
  'names.error.refersToEmpty': 'Type what the name refers to, such as =Sheet1!$A$1:$A$10.',
  'names.error.refersToSyntax': "This isn't a formula Excel can read.",
  'names.error.noSheet': 'This workbook has no sheet called "{{sheet}}".',
  'names.error.gone': 'The name "{{name}}" no longer exists.',
  'names.error.noScope': "The sheet that name belongs to isn't in this workbook.",

  'names.warn.notCalculated': "Formulas that use {{name}} aren't calculated here; Excel calculates them when it opens the file.",

  'names.change.defined': 'Defined the name {{name}} as {{formula}}',
  'names.change.changed': 'Changed the name {{name}} to {{formula}}',
  'names.change.renamed': 'Renamed {{from}} to {{name}}, which refers to {{formula}}',
  'names.change.commented': 'Changed the comment on the name {{name}}',
  'names.change.deleted': 'Deleted the name {{name}}',

  'names.history.defined': 'Define name {{name}}',
  'names.history.edited': 'Edit name {{name}}',
  'names.history.deleted': 'Delete name {{name}}',

  'names.goTo.notRange': "{{name}} refers to {{formula}}, which isn't a range of cells, so it can't be selected.",
  'names.goTo.unknown': '"{{text}}" isn\'t a cell, a range or a name in this workbook.',
  'names.goTo.notWorksheet': 'Only cells on a worksheet can be named.',

  'names.delete.title': 'Delete the name {{name}}?',
  'names.delete.body': 'Formulas that use it will show #NAME? until it is defined again.',
  'names.delete.confirm': 'Delete name',

  'names.box.title': 'The selected cell, or its name. Type a cell, a range or a name and press Enter to go there, or a new name to name the selection.',
  'names.box.list': 'Names',
  'names.box.show': 'Show names',
  'names.box.none': 'This workbook has no names yet.',
  'names.box.noMatch': 'No names start with "{{text}}".',

  'names.menu.define': 'Define name…',
  'names.menu.manager': 'Name manager…',

  'names.define.title': 'New name',
  'names.edit.title': 'Edit name',
  'names.field.name': 'Name',
  'names.field.scope': 'Scope',
  'names.field.comment': 'Comment',
  'names.field.refersTo': 'Refers to',
  'names.scope.workbook': 'Workbook',
  'names.scope.fixed': "A name's scope can't be changed once it is defined.",
  'names.define.submit': 'OK',

  'names.manager.title': 'Name manager',
  'names.manager.new': 'New…',
  'names.manager.edit': 'Edit…',
  'names.manager.delete': 'Delete',
  'names.manager.close': 'Close',
  'names.manager.filter': 'Show',
  'names.manager.empty': 'This workbook has no names yet.',
  'names.manager.noMatch': 'No names match.',
  'names.manager.list': 'Names',
  'names.filter.all': 'All names',
  'names.filter.workbook': 'Names scoped to the workbook',
  'names.filter.sheet': 'Names scoped to a sheet',
  'names.filter.errors': 'Names with errors',
  'names.filter.noErrors': 'Names without errors',
  'names.col.name': 'Name',
  'names.col.value': 'Value',
  'names.col.refersTo': 'Refers to',
  'names.col.scope': 'Scope',
  'names.col.comment': 'Comment',
};
