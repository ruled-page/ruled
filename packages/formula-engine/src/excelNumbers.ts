/** Excel's own arithmetic on numbers, shared by the calculation engines. */

/** Rounds the 15-digit decimal value, as Excel does, rather than the binary one. */
export function excelRound(x: number, digits: number, mode: 'half' | 'up' | 'down'): number {
  if (!Number.isFinite(x) || x === 0) return x;
  const d = Math.trunc(digits);
  const sign = x < 0 ? -1 : 1;
  const [mantissa, exp] = Math.abs(x).toExponential(14).split('e');
  // All 15 digits are already left of the rounding point.
  if (Number(exp) + d >= 14) return x;
  const shifted = Number(`${mantissa}e${Number(exp) + d}`);
  const r = mode === 'half' ? Math.round(shifted) : mode === 'up' ? Math.ceil(shifted) : Math.floor(shifted);
  return sign * Number(`${r}e${-d}`);
}

/**
 * A number as Excel writes it when it becomes text ("a"&0.1+0.2 is "a0.3"):
 * 15 significant digits, in scientific notation when very large or small.
 */
export function numberText(n: number): string {
  if (n === 0) return '0';
  const [mantissa, e] = n.toExponential(14).split('e');
  const exp = Number(e);
  if (exp >= -4 && exp < 15) return String(Number(n.toPrecision(15)));
  const m = mantissa.replace(/\.?0+$/, '');
  return `${m}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
}

/**
 * A number as Excel reads one written out (in a formula, or typed into a
 * cell): digits past the fifteenth significant one become zeros rather than
 * being rounded, so 1234567890123456789 is 1234567890123450000.
 */
export function fifteenDigits(written: string): number {
  const m = /^([+-]?)(\d*)(?:\.(\d*))?(?:e([+-]?\d+))?$/i.exec(written.trim());
  if (!m) return Number(written);
  const [, sign, int, frac = '', exp = '0'] = m;
  const all = int + frac;
  const lead = all.length - all.replace(/^0+/, '').length;
  if (all.length - lead <= 15) return Number(written);
  const kept = all.slice(0, lead + 15).padEnd(all.length, '0');
  return Number(`${sign}${kept.slice(0, int.length) || '0'}.${kept.slice(int.length)}e${exp}`);
}
