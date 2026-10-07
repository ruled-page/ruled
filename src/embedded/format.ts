import { useTranslation } from 'react-i18next';
import { NS } from './i18n';

/** Numbers, money, dates and percentages in the portal's language. */
export function useFormat() {
  const { i18n } = useTranslation(NS);
  const lng = i18n.language;
  return {
    number: (n: number) => new Intl.NumberFormat(lng).format(n),
    money: (n: number, currency: string) =>
      currency ? new Intl.NumberFormat(lng, { style: 'currency', currency }).format(n) : new Intl.NumberFormat(lng, { minimumFractionDigits: 2 }).format(n),
    change: (fraction: number) => new Intl.NumberFormat(lng, { style: 'percent', maximumFractionDigits: 1, signDisplay: 'exceptZero' }).format(fraction),
    /** An ISO date ("2026-11-01") as the language writes it, without shifting it by time zone. */
    day: (iso: string) => (iso ? new Intl.DateTimeFormat(lng, { dateStyle: 'medium', timeZone: 'UTC' }).format(new Date(`${iso.slice(0, 10)}T00:00:00Z`)) : ''),
    moment: (iso: string) => new Intl.DateTimeFormat(lng, { dateStyle: 'medium', timeStyle: 'short' }).format(new Date(iso)),
  };
}
