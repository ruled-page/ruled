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
    // Packages are never loaded: the types the exports use are the editor's own, so imports of
    // other packages' types fall away, and check:bundle fails if any but React's remain.
    external: (id) => !id.startsWith('.') && !id.startsWith('/'),
    output: { dir: 'dist/types', entryFileNames: '[name].d.ts', cleanDir },
  });

// The first clears what earlier builds left in dist/types.
export default [entry('index', 'dist/declarations/editor/index.d.ts', true), entry('standalone', 'dist/declarations/editor/standalone.d.ts')];
