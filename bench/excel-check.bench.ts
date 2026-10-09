import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { it } from 'vitest';
import { tolerance } from '../src/core/engine';
import { engineFormula, functionNames, toFileFormula } from '../src/core/formula';
import { nativeEngine } from '../src/engine/native';
import { FAMILIES } from '../tests/excel-checks/cases';

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

it('writes the Excel check workbooks', async () => {
  for (const [family, formulas] of Object.entries(FAMILIES)) {
    const engine = nativeEngine.create([{ name: 'Check', rows: formulas.map((f) => [engineFormula(toFileFormula(f).slice(1))]) }], [], false);
    const answers: Answer[] = formulas.map((_f, i) => engine.value(0, i, 0) as Answer);
    const missing = formulas.filter((f) => nativeEngine.unsupportedReason(toFileFormula(f).slice(1)));
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
          ? `AND(ISNUMBER(B${r}),ABS(B${r}-C${r})<=${tolerance(functionNames(toFileFormula(f).slice(1)))}*MAX(1,ABS(C${r})))`
          : typeof ours === 'string'
            ? `AND(ISTEXT(B${r}),EXACT(B${r},C${r}))`
            : typeof ours === 'boolean'
              ? `AND(ISLOGICAL(B${r}),B${r}=C${r})`
              : ours && typeof ours === 'object'
                ? `AND(ISERROR(B${r}),IFERROR(ERROR.TYPE(B${r})=ERROR.TYPE(C${r}),FALSE))`
                : 'FALSE';
      ws.getCell(`D${r}`).value = { formula: `IF(IFERROR(${same},FALSE),"✓","✗")` };
      ws.getCell(`E${r}`).value = { formula: `IF(D${r}="✓","",IF(ISERROR(B${r}),B${r},B${r}&""))` };
    });
    ws.addConditionalFormatting({
      ref: `A${first}:E${last}`,
      rules: [{ type: 'expression', priority: 1, formulae: [`$D${first}="✗"`], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }],
    });
    await wb.xlsx.writeFile(join(dir, `${family}.xlsx`));
    writeFileSync(join(dir, `${family}.pending.json`), `${JSON.stringify(Object.fromEntries(formulas.map((f, i) => [f, answers[i]])), null, 2)}\n`);
    console.log(`${family}: ${formulas.length} formulas, written to tests/excel-checks/${family}.xlsx${missing.length ? `; not calculated here: ${missing.join(' ')}` : ''}`);
  }
});
