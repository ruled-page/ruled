import type { charts as en } from '../en/charts';

/** German messages for drawn charts; every English key needs one (the type checks it). */
export const charts: Record<keyof typeof en, string> = {
  'charts.seriesName': 'Datenreihe{{n}}',
  'charts.noData': 'Es gibt keine Daten zum Darstellen.',
  'charts.menu.insert': 'Diagramm einfügen',
  'charts.history.insert': 'Diagramm einfügen',
  /** The title Excel gives a new chart of several series, written into the file. */
  'charts.placeholderTitle': 'Diagrammtitel',
  'charts.seriesCount_one': '{{count, number}} Datenreihe.',
  'charts.seriesCount_other': '{{count, number}} Datenreihen.',
  'charts.series': '{{name}}: {{points}}.',
  'charts.point': '{{label}} {{value}}',
  'charts.xy': 'x {{x}}, y {{y}}',
  'charts.blank': 'leer',
  'charts.more_one': '{{count, number}} weiterer',
  'charts.more_other': '{{count, number}} weitere',
  'charts.horizontalAxis': 'Horizontale Achse: {{title}}.',
  'charts.verticalAxis': 'Vertikale Achse: {{title}}.',
  'charts.secondaryAxis': 'Sekundärachse: {{title}}.',
};
