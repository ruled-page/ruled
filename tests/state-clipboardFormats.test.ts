import { describe, expect, it } from 'vitest';
import { clipboardFormats } from '../src/state/clipboardFormats';

/** A cut-down copy of what Excel puts on the clipboard for two rows of formatted cells. */
const EXCEL = `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><style>
<!--table
	{mso-displayed-decimal-separator:"\\.";
	mso-displayed-thousand-separator:"\\,";}
td
	{padding-top:1px; color:black; font-size:11.0pt; font-weight:400; font-family:Calibri, sans-serif; border:none;}
.xl65
	{font-weight:700;
	text-decoration:underline;
	text-underline-style:single;}
.xl66
	{color:red;
	mso-number-format:"\\0022$\\0022\\#\\,\\#\\#0\\.00";
	text-align:center;
	background:yellow;
	mso-pattern:black none;}
.xl67
	{border-top:.5pt solid windowtext;
	border-bottom:1.5pt double #0070C0;
	white-space:normal;
	font-style:italic;
	mso-number-format:Percent;}
-->
</style></head><body><table>
<tr><td class=xl65>Item</td><td class=xl66 x:num="1234.5">$1,234.50</td></tr>
<tr><td colspan=2 class=xl67 x:num="0.5">50.00%</td></tr>
<tr><td>plain</td><td class="xl99">unknown class</td></tr>
</table></body></html>`;

describe('formats pasted from another app', () => {
  it("are read from Excel's style classes, number formats included", () => {
    expect(clipboardFormats(EXCEL)).toEqual([
      [{ bold: true, underline: true }, { color: '#FF0000', numFmt: '"$"#,##0.00', hAlign: 'center', fill: '#FFFF00' }],
      [
        {
          italic: true,
          wrap: true,
          numFmt: '0.00%',
          border: { top: { width: 1, pattern: 'solid' }, bottom: { width: 3, pattern: 'double', color: '#0070C0' } },
        },
        null,
      ],
      [null, null],
    ]);
  });

  it("are read from Google Sheets' style attributes", () => {
    const html =
      '<google-sheets-html-origin><style type="text/css"><!--td {border: 1px solid #cccccc;}--></style><table><tbody><tr>' +
      '<td style="overflow:hidden;vertical-align:middle;font-family:&quot;Roboto&quot;;font-size:10pt;font-weight:bold;background-color:rgb(255,242,204);">A</td>' +
      '<td style="color:#1155cc;text-decoration:line-through;border-right:2px dashed #000">B</td></tr></tbody></table>';
    expect(clipboardFormats(html)).toEqual([
      [
        { vAlign: 'center', fontFamily: 'Roboto', fontSizePt: 10, bold: true, fill: '#FFF2CC' },
        { color: '#1155CC', strike: true, border: { right: { width: 2, pattern: 'dashed', color: '#000000' } } },
      ],
    ]);
  });

  it('are nothing without a table, or with no formats in it, or past the size limit', () => {
    expect(clipboardFormats('<p>hello</p>')).toBeUndefined();
    expect(clipboardFormats('<table><tr><td>a</td></tr></table>')).toBeUndefined();
    expect(clipboardFormats(`<table><tr><td style="font-weight:bold">a</td></tr></table>${' '.repeat(4_000_001)}`)).toBeUndefined();
  });
});
