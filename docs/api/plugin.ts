/**
 * What plugins are written against: the contract's types, and
 * `definePlugin` for type inference. A plugin imports only types from it,
 * so it never brings a second copy of the editor:
 * `import type { RuledPlugin } from '@ruled/sheet/plugin'`.
 *
 * @module @ruled/sheet/plugin
 */
export * from '../../packages/spreadsheet/src/plugin/index';
