import { defineConfig } from 'rolldown';
import { dts } from 'rolldown-plugin-dts';

/**
 * The library's type declarations, one file for each entry point
 * (dist/types/index.d.ts and standalone.d.ts): tsc's declarations for the
 * editor (dist/declarations) bundled with only the types their exports use,
 * so they import nothing but React, and a host installs no other package for
 * them.
 */
const entry = (name: string, input: string, cleanDir = false) =>
  defineConfig({
    input: { [name]: input },
    plugins: [dts({ dtsInput: true })],
    output: { dir: 'dist/types', entryFileNames: '[name].d.ts', cleanDir },
  });

// The first clears what earlier builds left in dist/types.
export default [entry('index', 'dist/declarations/editor/index.d.ts', true), entry('standalone', 'dist/declarations/editor/standalone.d.ts')];
