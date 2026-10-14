import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { sample } from './helpers';

const notes = (s: WorkbookSession) => s.model.sheets.find((m) => m.info.name === 'Notes')!.info;

describe('pictures', () => {
  it('know their image, which the session hands over for the pictures that use it only', () => {
    const s = WorkbookSession.open(sample('kitchen-sink.xlsx'), 'k.xlsx');
    const pic = notes(s).objects.find((o) => o.kind === 'image')!;
    expect(pic.picture).toEqual({ part: 'xl/media/image1.png' });
    expect(s.picture('xl/media/image1.png')).toMatchObject({ type: 'image/png' });
    expect(s.picture('xl/media/image1.png')!.bytes.slice(1, 4)).toEqual(strToU8('PNG'));
    expect(s.picture('xl/workbook.xml')).toBeUndefined();
    expect(s.picture('xl/media/other.png')).toBeUndefined();
  });

  it('read how a picture is cropped, turned and flipped', () => {
    const files = unzipSync(sample('kitchen-sink.xlsx'));
    files['xl/drawings/drawing2.xml'] = strToU8(
      strFromU8(files['xl/drawings/drawing2.xml'])
        .replace('<a:stretch>', '<a:srcRect l="10000" t="25000" b="-5000"/><a:stretch>')
        .replace('<xdr:spPr><a:xfrm>', '<xdr:spPr><a:xfrm rot="5400000" flipH="1">'),
    );
    const pic = notes(WorkbookSession.open(zipSync(files), 'k.xlsx')).objects.find((o) => o.kind === 'image')!;
    expect(pic.picture).toEqual({ part: 'xl/media/image1.png', crop: { l: 0.1, t: 0.25, r: 0, b: -0.05 }, rotate: 90, flipH: true });
  });
});
