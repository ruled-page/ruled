import { fileURLToPath } from 'node:url';

/** The app is a project of its own inside this repository: its root is here, not the repository's. */
const root = fileURLToPath(new URL('.', import.meta.url));

export default { turbopack: { root }, outputFileTracingRoot: root };
