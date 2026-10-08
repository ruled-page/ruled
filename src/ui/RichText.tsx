import type { ComponentProps } from 'react';
import { Trans } from 'react-i18next';

/**
 * A message with markup such as <b>, whose values are text: a schema's name
 * or a sheet's is escaped before the markup is read, and shown as written,
 * so a name like "<i style=…>" can't become markup.
 */
export function RichText(props: ComponentProps<typeof Trans>) {
  return <Trans {...props} tOptions={{ ...props.tOptions, interpolation: { escapeValue: true } }} shouldUnescape />;
}
