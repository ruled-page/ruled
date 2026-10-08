import { charts } from './charts';
import { core } from './core';
import { designer } from './designer';
import { names } from './names';
import { links } from './links';
import { schema } from './schema';
import { sheets } from './sheets';
import { store } from './store';
import { ui } from './ui';

/** Every message, in German. */
export const de = { ...ui, ...designer, ...store, ...schema, ...core, ...charts, ...sheets, ...names, ...links };
