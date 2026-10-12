/**
 * Whether a regular expression can take exponential time on some text: a
 * group that repeats without limit (*, + or {n,}) and holds something that
 * repeats or alternates, such as (a+)+ or (a|aa)*. JavaScript's engine
 * backtracks through every way of splitting the text between them, which
 * for 30 characters can take minutes. Schemas loaded from JSON or made in
 * the designer are checked with this; a host's own code is trusted.
 */
export function isUnsafePattern(source: string): boolean {
  /** For each open group: whether it holds a quantifier or an alternation so far. */
  const groups: boolean[] = [];
  let inClass = false;
  for (let i = 0; i < source.length; i++) {
    const c = source[i];
    if (c === '\\') {
      i++;
      continue;
    }
    if (inClass) {
      if (c === ']') inClass = false;
      continue;
    }
    if (c === '[') inClass = true;
    else if (c === '(') groups.push(false);
    else if (c === ')') {
      const risky = groups.pop() ?? false;
      const after = source.slice(i + 1);
      const unbounded = /^(?:[*+]|\{\d+,\})/.test(after);
      if (risky && unbounded) return true;
      // A group that repeats makes the group around it risky too.
      if (groups.length && (risky || /^(?:[*+?]|\{\d*,?\d*\})/.test(after))) groups[groups.length - 1] = true;
    } else if ((c === '*' || c === '+' || c === '?' || c === '{' || c === '|') && groups.length) {
      groups[groups.length - 1] = true;
    }
  }
  return false;
}
