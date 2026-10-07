import { core } from './core';
import { designer } from './designer';
import { schema } from './schema';
import { store } from './store';
import { ui } from './ui';

/** Every message, in English. */
export const en = { ...ui, ...designer, ...store, ...schema, ...core };
