import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { it } from 'vitest';
import { tolerance } from '@sitepen/formula-engine';
import { functionNames, isDynamicArrayFormula, toFileFormula } from '@sitepen/formula-engine/formula';
import { dynamicArrayMetadata } from '../src/core/writer';
import { OoxmlPackage } from '../src/core/zip';
import { nativeEngine } from '@sitepen/formula-engine/native';
import { checkEngine, DATA, FAMILIES, ITERATIONS, NAMES } from '../tests/excel-checks/cases';

/**
 * Writes tests/excel-checks/<family>.xlsx for each family of check formulas:
 * each formula, the native engine's answer beside it, and a cell Excel fills
 * on opening with whether the two agree (and Excel's own value where they
 * don't). The top says how many match. Also writes <family>.pending.json,
 * the native engine's answers: once Excel agrees with them, renamed to
 * <family>.json they are the expected values tests/excel-checks.test.ts holds the engine to.
 */

const dir = join(__dirname, '..', 'tests', 'excel-checks');

type Answer = number | string | boolean | { error: string } | null;

/** ERROR.TYPE's number for each error; the file can hold only the first seven as error values, the rest as text. */
const ERROR_TYPES: Record<string, number> = {
  '#NULL!': 1,
  '#DIV/0!': 2,
  '#VALUE!': 3,
  '#REF!': 4,
  '#NAME?': 5,
  '#NUM!': 6,
  '#N/A': 7,
  '#GETTING_DATA': 8,
  '#SPILL!': 9,
  '#CONNECT!': 10,
  '#BLOCKED!': 11,
  '#UNKNOWN!': 12,
  '#FIELD!': 13,
  '#CALC!': 14,
};

it('writes the Excel check workbooks', async () => {
  for (const [family, formulas] of Object.entries(FAMILIES)) {
    // FAMILY=name writes that family's workbook alone.
    if (process.env.FAMILY && process.env.FAMILY !== family) continue;
    const engine = checkEngine(family);
    const iteration = ITERATIONS[family];
    const answers: Answer[] = formulas.map((_f, i) => engine.value(0, i, 0) as Answer);
    const lambdas = new Set(Object.keys(NAMES[family] ?? {}));
    const missing = formulas.filter((f) => nativeEngine.unsupportedReason(toFileFormula(f).slice(1), lambdas));
    engine.destroy();

    const wb = new ExcelJS.Workbook();
    wb.calcProperties.fullCalcOnLoad = true;
    const ws = wb.addWorksheet('Check');
    ws.columns = [{ width: 60 }, { width: 18 }, { width: 22 }, { width: 8 }, { width: 24 }];
    const first = 4;
    const last = first + formulas.length - 1;
    ws.getCell('A1').value = 'Excel agrees with the native engine on';
    ws.getCell('B1').value = { formula: `COUNTIF(D${first}:D${last},"✓")&" of "&ROWS(D${first}:D${last})` };
    ws.getCell('A2').value = { formula: `IF(COUNTIF(D${first}:D${last},"✓")=ROWS(D${first}:D${last}),"Every formula matches","Rows in red differ: column E has Excel's value")` };
    ws.getRow(3).values = ['Formula', 'Excel', 'Native engine', 'Same?', "Excel's value, where different"];
    ws.getRow(1).font = { bold: true };
    ws.getRow(3).font = { bold: true };
    formulas.forEach((f, i) => {
      const r = first + i;
      const ours = answers[i];
      ws.getCell(`A${r}`).value = f;
      ws.getCell(`B${r}`).value = { formula: toFileFormula(f).slice(1) } as ExcelJS.CellFormulaValue;
      ws.getCell(`C${r}`).value = ours === null ? '(nothing)' : (ours as ExcelJS.CellValue);
      const same =
        typeof ours === 'number'
          ? `AND(ISNUMBER(B${r}),ABS(B${r}-C${r})<=${Math.max(tolerance(functionNames(toFileFormula(f).slice(1))), iteration?.delta ?? 0)}*MAX(1,ABS(C${r})))`
          : typeof ours === 'string'
            ? `AND(ISTEXT(B${r}),EXACT(B${r},C${r}))`
            : typeof ours === 'boolean'
              ? `AND(ISLOGICAL(B${r}),B${r}=C${r})`
              : ours && typeof ours === 'object'
                ? `AND(ISERROR(B${r}),IFERROR(ERROR.TYPE(B${r})=${ERROR_TYPES[ours.error]},FALSE))`
                : 'FALSE';
      ws.getCell(`D${r}`).value = { formula: `IF(IFERROR(${same},FALSE),"✓","✗")` };
      ws.getCell(`E${r}`).value = { formula: `IF(D${r}="✓","",IF(ISERROR(B${r}),B${r},B${r}&""))` };
    });
    ws.addConditionalFormatting({
      ref: `A${first}:E${last}`,
      rules: [{ type: 'expression', priority: 1, formulae: [`$D${first}="✗"`], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }],
    });
    // Iterating, Excel starts from the values saved in the file, and a formula saved without one reads as text: #VALUE! all the way round. So
    // those start from 0, as the engine's do.
    const start = iteration ? { result: 0 } : {};
    for (const [name, rows] of Object.entries(DATA[family] ?? {})) {
      const data = wb.addWorksheet(name);
      rows.forEach((row, r) =>
        row.forEach((v, c) => {
          if (v !== null) data.getCell(r + 1, c + 1).value = typeof v === 'string' && v.startsWith('=') ? { formula: toFileFormula(v).slice(1), ...start } as ExcelJS.CellFormulaValue : v;
        }),
      );
    }
    // Formulas using array functions are marked as dynamic array formulas, as
    // the editor saves them; without the mark Excel shows #NAME? for RANDARRAY.
    const pkg = OoxmlPackage.open(new Uint8Array(await wb.xlsx.writeBuffer()));
    const dynamic = formulas.map((f) => isDynamicArrayFormula(toFileFormula(f).slice(1)));
    if (dynamic.some(Boolean)) {
      const cm = dynamicArrayMetadata(pkg, 'xl/workbook.xml');
      const sheet = 'xl/worksheets/sheet1.xml';
      let xml = pkg.text(sheet)!;
      dynamic.forEach((d, i) => {
        const r = first + i;
        if (d) xml = xml.replace(new RegExp(`<c r="B${r}"([^>]*)><f>`), `<c r="B${r}"$1 cm="${cm}"><f t="array" ref="B${r}">`);
      });
      pkg.setText(sheet, xml);
    }
    if (iteration) {
      const xml = pkg.text('xl/workbook.xml')!;
      pkg.setText('xl/workbook.xml', xml.replace(/<calcPr\b/, `<calcPr iterate="1" iterateCount="${iteration.count}" iterateDelta="${iteration.delta}"`));
    }
    const names = Object.entries(NAMES[family] ?? {});
    if (names.length) {
      const escape = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      const defined = names.map(([name, f]) => `<definedName name="${name}">${escape(toFileFormula(f).slice(1))}</definedName>`).join('');
      pkg.setText('xl/workbook.xml', pkg.text('xl/workbook.xml')!.replace('</sheets>', `</sheets><definedNames>${defined}</definedNames>`));
    }
    writeFileSync(join(dir, `${family}.xlsx`), pkg.toBytes());
    writeFileSync(join(dir, `${family}.pending.json`), `${JSON.stringify(Object.fromEntries(formulas.map((f, i) => [f, answers[i]])), null, 2)}\n`);
    console.log(`${family}: ${formulas.length} formulas, written to tests/excel-checks/${family}.xlsx${missing.length ? `; not calculated here: ${missing.join(' ')}` : ''}`);
  }
});
