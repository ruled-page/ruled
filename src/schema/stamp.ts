import { msg, pickLanguage, type Text } from '../i18n/message';
import type { WorkbookSchema } from './types';

/** The custom document property that records the schema a file was saved with, as "id@version". */
export const SCHEMA_PROPERTY = 'WorkbookSchema';

/** The schema id and version a file was stamped with, if any. */
export function parseStamp(value: string | undefined): { id: string; version: string } | undefined {
  if (!value) return undefined;
  const at = value.lastIndexOf('@');
  return at < 0 ? { id: value, version: '' } : { id: value.slice(0, at), version: value.slice(at + 1) };
}

/** The properties to stamp a file with when saving under `schema`; none for an unversioned schema. */
export function stampFor(schema: WorkbookSchema | undefined): Record<string, string> | undefined {
  return schema?.version ? { [SCHEMA_PROPERTY]: `${schema.id}@${schema.version}` } : undefined;
}

/** Schemas in the order to try them: the one a file was stamped with first. */
export function stampedFirst(schemas: WorkbookSchema[], properties: Record<string, string>): WorkbookSchema[] {
  const stamp = parseStamp(properties[SCHEMA_PROPERTY]);
  return stamp ? [...schemas].sort((a, b) => Number(b.id === stamp.id) - Number(a.id === stamp.id)) : schemas;
}

/**
 * A notice when a file was saved under a different version of the schema now
 * checking it. The schema's name is given in `language` when it has several.
 */
export function versionNote(schema: WorkbookSchema, properties: Record<string, string>, language = 'en'): Text | undefined {
  const stamp = parseStamp(properties[SCHEMA_PROPERTY]);
  if (stamp?.id !== schema.id || stamp.version === (schema.version ?? '')) return undefined;
  const name = pickLanguage(schema.name, language);
  if (!stamp.version) return msg('schema.version.wasUnversioned', { schema: name, current: schema.version });
  return schema.version
    ? msg('schema.version.changed', { schema: name, saved: stamp.version, current: schema.version })
    : msg('schema.version.nowUnversioned', { schema: name, saved: stamp.version });
}
