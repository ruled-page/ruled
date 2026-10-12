import { expect, it } from 'vitest';
import { argumentAt, completionAt, functionHelpLoaded, loadFunctionHelp } from '../src/core/formulaAssist';

it('offers functions once their list has loaded, which happens once', async () => {
  expect(functionHelpLoaded()).toBe(false);
  expect(completionAt('=SU', 3)).toBeUndefined();
  expect(argumentAt('=SUM(1,', 7)).toBeUndefined();
  const first = loadFunctionHelp();
  expect(loadFunctionHelp()).toBe(first);
  await first;
  expect(functionHelpLoaded()).toBe(true);
  expect(completionAt('=SU', 3)?.matches.map((f) => f.name)).toContain('SUM');
  expect(argumentAt('=SUM(1,', 7)?.index).toBe(1);
});
