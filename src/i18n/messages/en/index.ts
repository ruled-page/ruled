import { charts } from './charts';
import { core } from './core';
import { designer } from './designer';
import { names } from './names';
import { schema } from './schema';
import { sheets } from './sheets';
import { store } from './store';
import { ui } from './ui';

/** Every message, in English. */
export const en = { ...ui, ...designer, ...store, ...schema, ...core, ...charts, ...sheets, ...names };
