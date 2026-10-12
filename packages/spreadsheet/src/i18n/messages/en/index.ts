import { charts } from './charts';
import { core } from './core';
import { designer } from './designer';
import { fill } from './fill';
import { names } from './names';
import { links } from './links';
import { outline } from './outline';
import { schema } from './schema';
import { sheets } from './sheets';
import { store } from './store';
import { ui } from './ui';

/** Every message, in English. */
export const en = { ...ui, ...designer, ...store, ...schema, ...core, ...charts, ...sheets, ...names, ...links, ...fill, ...outline };
