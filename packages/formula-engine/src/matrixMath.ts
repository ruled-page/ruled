/** Matrix arithmetic shared by the calculation engines. */

/**
 * Gauss-Jordan elimination with partial pivoting, on the matrix and an
 * identity beside it. Returns the determinant and the inverse (undefined
 * when the matrix is singular).
 */
export function eliminate(a: number[][]): { det: number; inverse?: number[][] } {
  const n = a.length;
  const inv = a.map((_, i) => a.map((__, j) => (i === j ? 1 : 0)));
  let det = 1;
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(a[r][c]) > Math.abs(a[p][c])) p = r;
    if (a[p][c] === 0) return { det: 0 };
    if (p !== c) {
      [a[p], a[c]] = [a[c], a[p]];
      [inv[p], inv[c]] = [inv[c], inv[p]];
      det = -det;
    }
    const pivot = a[c][c];
    det *= pivot;
    for (let j = 0; j < n; j++) {
      a[c][j] /= pivot;
      inv[c][j] /= pivot;
    }
    for (let r = 0; r < n; r++) {
      if (r === c || a[r][c] === 0) continue;
      const f = a[r][c];
      for (let j = 0; j < n; j++) {
        a[r][j] -= f * a[c][j];
        inv[r][j] -= f * inv[c][j];
      }
    }
  }
  return { det, inverse: inv };
}
