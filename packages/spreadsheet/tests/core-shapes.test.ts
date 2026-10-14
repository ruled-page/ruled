import { describe, expect, it } from 'vitest';
import { parseShape } from '../src/core/shapes';
import { parseXml } from '../src/core/xml';
import { shapePath } from '../src/ui/ShapeView';
import { open } from './helpers';

// Office's theme: lt1, dk1, lt2, dk2, then the accents.
const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31'];
const shape = (xml: string) => parseShape(parseXml(`<xdr:sp>${xml}</xdr:sp>`), THEME);

describe('shapes', () => {
  it("read a text box as Excel draws it: its fill, line and text, sizes in pixels", () => {
    const s = open('kitchen-sink.xlsx');
    const box = s.model.sheets.find((m) => m.info.name === 'Notes')!.info.objects.find((o) => o.name === 'TextBox 2')!;
    expect(box.shape).toEqual({
      geometry: 'rect',
      fill: '#FFFFFF',
      line: { color: '#BCBCBC', width: 1 },
      text: {
        paragraphs: [{ align: 'left', runs: [{ text: 'Text boxes are kept too.', size: 1100 / 75, color: '#000000' }] }],
        anchor: 'top',
        insets: [9.6, 4.8, 9.6, 4.8],
      },
    });
  });

  it("take a shape's fill, line and text colour from its style where it sets none, as Excel's default shapes do", () => {
    const s = shape(
      '<xdr:spPr><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 30000"/></a:avLst></a:prstGeom></xdr:spPr>' +
        '<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
        '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>' +
        '<xdr:txBody><a:bodyPr anchor="ctr"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr b="1"/><a:t>Go</a:t></a:r><a:br/><a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>now</a:t></a:r></a:p></xdr:txBody>',
    )!;
    expect(s).toMatchObject({ geometry: 'roundRect', adjust: { adj: 30000 }, fill: '#4472C4', line: { color: '#2F528F', width: 12700 / 9525 } });
    expect(s.text).toMatchObject({ anchor: 'middle', paragraphs: [{ align: 'center', runs: [{ text: 'Go', bold: true, color: '#FFFFFF' }, { text: '\n' }, { text: 'now', color: '#FF0000', font: 'Arial' }] }] });
  });

  it('read a line with arrowheads and a dash, turned and flipped, and leave kinds the editor does not draw', () => {
    const line = parseShape(
      parseXml(
        '<xdr:cxnSp><xdr:spPr><a:xfrm rot="5400000" flipH="1"/><a:prstGeom prst="straightConnector1"/><a:ln w="19050"><a:solidFill><a:srgbClr val="00B050"/></a:solidFill><a:prstDash val="dash"/><a:tailEnd type="triangle"/><a:headEnd type="oval"/></a:ln></xdr:spPr></xdr:cxnSp>',
      ),
      THEME,
    );
    expect(line).toEqual({ geometry: 'straightConnector1', line: { color: '#00B050', width: 2, dash: 'dash', head: 'oval', tail: 'arrow' }, rotate: 90, flipH: true });
    expect(shape('<xdr:spPr><a:prstGeom prst="cloud"/></xdr:spPr>')).toBeUndefined();
    expect(shape('<xdr:spPr><a:custGeom/></xdr:spPr>')).toBeUndefined();
  });

  it('draw the presets with their adjustments', () => {
    expect(shapePath('rect', 100, 50)).toBe('M0 0 L100 0 L100 50 L0 50 Z');
    expect(shapePath('rightArrow', 100, 40)).toBe('M0 10 L80 10 L80 0 L100 20 L80 40 L80 30 L0 30 Z');
    expect(shapePath('diamond', 40, 20)).toBe('M20 0 L40 10 L20 20 L0 10 Z');
    expect(shapePath('triangle', 40, 20, { adj: 0 })).toBe('M0 0 L40 20 L0 20 Z');
  });
});
