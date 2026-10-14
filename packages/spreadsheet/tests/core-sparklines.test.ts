import { describe, expect, it } from 'vitest';
import { evaluateSparklines, parseSparklines, sparklineSource } from '../src/core/sparklines';
import { parseXml } from '../src/core/xml';
import { open } from './helpers';

const groups = (xml: string) => parseSparklines(parseXml(`<worksheet><extLst><ext><x14:sparklineGroups>${xml}</x14:sparklineGroups></ext></extLst></worksheet>`), (c) => c && `#${c.attrs.rgb?.slice(2)}`);

describe('sparklines', () => {
  it('read a group with its options, taking Excel defaults for the rest', () => {
    const [g] = groups(
      '<x14:sparklineGroup type="column" negative="1" high="1" displayXAxis="1" minAxisType="group" maxAxisType="custom" manualMax="10" lineWeight="1.5">' +
        '<x14:colorSeries rgb="FF112233"/><x14:colorNegative rgb="FFAA0000"/>' +
        '<x14:sparklines><x14:sparkline><xm:f>\'Q 1\'!$B$2:$E$2</xm:f><xm:sqref>F2</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>',
    );
    expect(g).toMatchObject({ type: 'column', lineWeight: 1.5, empty: 'zero', minAxis: 'group', maxAxis: 'custom', manualMax: 10, displayHidden: false, rightToLeft: false });
    expect(g.colors).toMatchObject({ series: '#112233', negative: '#AA0000' });
    expect(g.show).toEqual({ markers: false, high: true, low: false, first: false, last: false, negative: true, axis: true });
    expect(g.lines).toEqual([{ row: 1, col: 5, source: "'Q 1'!$B$2:$E$2" }]);
    expect(sparklineSource(g.lines[0].source)).toEqual({ sheet: 'Q 1', rect: { r1: 1, c1: 1, r2: 1, c2: 4 } });
    expect(sparklineSource('A1:A3')).toEqual({ rect: { r1: 0, c1: 0, r2: 2, c2: 0 } });
  });

  it('work out the range each line is drawn against, its own, its group\'s or one set, empty cells counting as zero where asked', () => {
    const lines = '<x14:sparklines><x14:sparkline><xm:f>A1:C1</xm:f><xm:sqref>D1</xm:sqref></x14:sparkline><x14:sparkline><xm:f>A2:C2</xm:f><xm:sqref>D2</xm:sqref></x14:sparkline></x14:sparklines>';
    const data: Record<string, (number | null)[]> = { 'A1:C1': [1, null, 3], 'A2:C2': [5, 6, 9] };
    const at = (xml: string) => evaluateSparklines(groups(xml), (source) => data[source]).map((c) => [c.min, c.max]);
    expect(at(`<x14:sparklineGroup displayEmptyCellsAs="gap">${lines}</x14:sparklineGroup>`)).toEqual([[1, 3], [5, 9]]);
    expect(at(`<x14:sparklineGroup>${lines}</x14:sparklineGroup>`)).toEqual([[0, 3], [5, 9]]);
    expect(at(`<x14:sparklineGroup minAxisType="group" maxAxisType="group">${lines}</x14:sparklineGroup>`)).toEqual([[0, 9], [0, 9]]);
    expect(at(`<x14:sparklineGroup minAxisType="custom" manualMin="-2">${lines}</x14:sparklineGroup>`)).toEqual([[-2, 3], [-2, 9]]);
  });

  it('are drawn from the values their cells hold now', () => {
    const s = open('kitchen-sink.xlsx');
    const notes = s.model.sheets.findIndex((m) => m.info.name === 'Notes');
    const [spark] = s.conditionalFormats(notes).sparklines!;
    expect(spark).toMatchObject({ row: 9, col: 6, group: 0 });
    expect(spark.values).toHaveLength(5);
    expect(s.model.sheets[notes].info.sparklines![0].show.markers).toBe(true);
    s.apply({ type: 'input', sheet: notes, row: 9, col: 1, text: '1000' });
    expect(s.conditionalFormats(notes).sparklines![0].values[0]).toBe(1000);
  });
});
