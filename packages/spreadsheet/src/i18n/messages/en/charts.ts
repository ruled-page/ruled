/** English messages for drawn charts: what screen readers hear in place of the drawing. */
export const charts = {
  /** A series the workbook doesn't name, called what Excel calls it. */
  'charts.seriesName': 'Series{{n}}',
  'charts.noData': 'There is no data to plot.',
  'charts.menu.insert': 'Insert chart',
  'charts.history.insert': 'Insert chart',
  /** The title Excel gives a new chart of several series, written into the file. */
  'charts.placeholderTitle': 'Chart Title',
  'charts.seriesCount_one': '{{count, number}} series.',
  'charts.seriesCount_other': '{{count, number}} series.',
  /** {{points}} lists charts.point (charts.xy on a scatter chart), perhaps ending with charts.more. */
  'charts.series': '{{name}}: {{points}}.',
  /** A category and its value, as their cells show them. */
  'charts.point': '{{label}} {{value}}',
  /** A point on a scatter chart. */
  'charts.xy': 'x {{x}}, y {{y}}',
  'charts.blank': 'blank',
  'charts.more_one': '{{count, number}} more',
  'charts.more_other': '{{count, number}} more',
  'charts.horizontalAxis': 'Horizontal axis: {{title}}.',
  'charts.verticalAxis': 'Vertical axis: {{title}}.',
  'charts.secondaryAxis': 'Secondary axis: {{title}}.',
} satisfies Record<string, string>;
