import { useMemo } from 'react';
import { useTranslation } from 'react-i18next';
import { NAMESPACE, translator, type Translator } from './index';

/**
 * The editor's translation functions, from the i18next instance the editor
 * provides. Components re-render, and memos keyed on it recompute, when the
 * language changes.
 */
export function useT(): Translator {
  const { i18n } = useTranslation(NAMESPACE);
  const language = i18n.language;
  // eslint-disable-next-line react-hooks/exhaustive-deps
  return useMemo(() => translator(i18n), [i18n, language]);
}
