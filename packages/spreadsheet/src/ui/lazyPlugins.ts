import { lazyModule } from './lazyModule';

/** What plugins add to the editor's UI, loaded only once a plugin is running. */
export const pluginParts = lazyModule(() => import('./PluginParts'));
