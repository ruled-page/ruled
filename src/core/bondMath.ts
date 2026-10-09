/** Dates as bonds and YEARFRAC count them, shared by the calculation engines. */

export const DAY_MS = 86_400_000;

export const isLeap = (y: number) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
export const daysInMonth = (m: number, y: number) => [31, isLeap(y) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1];

/** Serial numbers to calendar dates and back, from the workbook's date system. */
export class Calendar {
  private readonly base: number;

  constructor(nullDate: { year: number; month: number; day: number }) {
    this.base = Date.UTC(nullDate.year, nullDate.month - 1, nullDate.day);
  }

  ymd(serial: number): [number, number, number] {
    const d = new Date(this.base + Math.floor(serial) * DAY_MS);
    return [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()];
  }

  serial(y: number, m: number, d: number): number {
    return Math.round((Date.UTC(y, m - 1, d) - this.base) / DAY_MS);
  }
}

/**
 * A date as the coupon calculations move it: by months from its original
 * day, so 31 Aug less six months is 28 or 29 Feb and back again 31 Aug, and
 * in the 30-day bases (0 and 4) with every month 30 days long.
 */
export class CouponDate {
  year: number;
  month: number;
  day = 0;
  readonly origDay: number;
  readonly lastDay: boolean;
  readonly thirty: boolean;
  readonly us: boolean;

  constructor(calendar: Calendar, serial: number, basis: number) {
    const [y, m, d] = calendar.ymd(serial);
    this.year = y;
    this.month = m;
    this.origDay = d;
    this.lastDay = d >= daysInMonth(m, y);
    this.thirty = basis === 0 || basis === 4;
    this.us = basis === 0;
    this.setDay();
  }

  copy(): CouponDate {
    return Object.assign(Object.create(CouponDate.prototype) as CouponDate, this);
  }

  private setDay() {
    const last = daysInMonth(this.month, this.year);
    if (this.thirty) this.day = this.lastDay || Math.min(this.origDay, 30) >= last ? 30 : Math.min(this.origDay, 30);
    else this.day = this.lastDay ? last : Math.min(this.origDay, last);
  }

  daysInMonth(): number {
    return this.thirty ? 30 : daysInMonth(this.month, this.year);
  }

  addMonths(n: number): this {
    const m = this.month - 1 + n;
    this.year += Math.floor(m / 12);
    this.month = (((m % 12) + 12) % 12) + 1;
    this.setDay();
    return this;
  }

  setYear(y: number): this {
    this.year = y;
    this.setDay();
    return this;
  }

  serial(calendar: Calendar): number {
    const last = daysInMonth(this.month, this.year);
    return calendar.serial(this.year, this.month, this.lastDay ? last : Math.min(last, this.origDay));
  }

  compare(o: CouponDate): number {
    if (this.year !== o.year) return this.year - o.year;
    if (this.month !== o.month) return this.month - o.month;
    if (this.day !== o.day) return this.day - o.day;
    if (this.lastDay || o.lastDay) return Number(this.lastDay) - Number(o.lastDay);
    return this.origDay - o.origDay;
  }

  /** Days from a to b in the basis the dates were made with, never negative. */
  static diff(a: CouponDate, b: CouponDate): number {
    if (a.compare(b) > 0) return CouponDate.diff(b, a);
    const from = a.copy();
    const to = b.copy();
    if (b.thirty) {
      if (b.us) {
        // NASD: a 31st stays unless the start is the 30th or later; February's last day stays as it is.
        if ((a.month === 2 || a.day < 30) && to.origDay === 31) to.day = 31;
        else if (to.month === 2 && to.lastDay) to.day = daysInMonth(2, to.year);
      } else {
        if (from.month === 2 && from.day === 30) from.day = daysInMonth(2, from.year);
        if (to.month === 2 && to.day === 30) to.day = daysInMonth(2, to.year);
      }
    }
    let days = 0;
    if (from.year < to.year || (from.year === to.year && from.month < to.month)) {
      days = from.daysInMonth() - from.day + 1;
      const next = { year: from.year + (from.month === 12 ? 1 : 0), month: (from.month % 12) + 1 };
      while (next.year < to.year || (next.year === to.year && next.month < to.month)) {
        days += from.thirty ? 30 : daysInMonth(next.month, next.year);
        next.year += next.month === 12 ? 1 : 0;
        next.month = (next.month % 12) + 1;
      }
      days += to.day - 1;
    } else {
      days = to.day - from.day;
    }
    return Math.max(days, 0);
  }
}

/** The last coupon date on or before settlement. */
export function previousCoupon(settle: CouponDate, maturity: CouponDate, frequency: number): CouponDate {
  const d = maturity.copy().setYear(settle.year);
  if (d.compare(settle) < 0) d.setYear(d.year + 1);
  while (d.compare(settle) > 0) d.addMonths(-12 / frequency);
  return d;
}

/** The first coupon date after settlement. */
export function nextCoupon(settle: CouponDate, maturity: CouponDate, frequency: number): CouponDate {
  const d = maturity.copy().setYear(settle.year);
  if (d.compare(settle) > 0) d.setYear(d.year - 1);
  while (d.compare(settle) <= 0) d.addMonths(12 / frequency);
  return d;
}

/** The fraction of a year between two serials in a basis, as YEARFRAC works it out. */
export function yearFrac(calendar: Calendar, start: number, end: number, basis: number): number {
  if (start === end) return 0;
  if (start > end) [start, end] = [end, start];
  const [y1, m1, d1] = calendar.ymd(start);
  const [y2, m2, d2] = calendar.ymd(end);
  const days = end - start;
  switch (basis) {
    case 0: {
      let [a, b] = [d1, d2];
      const febEnd = (m: number, d: number, y: number) => m === 2 && d === daysInMonth(2, y);
      if (a === 31 && b === 31) [a, b] = [30, 30];
      else if (a === 31) a = 30;
      else if (a === 30 && b === 31) b = 30;
      else if (febEnd(m1, d1, y1) && febEnd(m2, d2, y2)) [a, b] = [30, 30];
      else if (febEnd(m1, d1, y1)) a = 30;
      return ((y2 - y1) * 360 + (m2 - m1) * 30 + (b - a)) / 360;
    }
    case 1: {
      // Up to a year apart: 366 days if a 29 Feb falls in the span; longer: the average year length.
      const withinYear = y1 === y2 || (y2 === y1 + 1 && (m1 > m2 || (m1 === m2 && d1 >= d2)));
      if (!withinYear) return days / ((calendar.serial(y2 + 1, 1, 1) - calendar.serial(y1, 1, 1)) / (y2 - y1 + 1));
      const feb29 =
        (isLeap(y1) && start < calendar.serial(y1, 3, 1) && end >= calendar.serial(y1, 3, 1)) ||
        (isLeap(y2) && end >= calendar.serial(y2, 3, 1) && start < calendar.serial(y2, 3, 1)) ||
        (m2 === 2 && d2 === 29);
      return days / ((y1 === y2 && isLeap(y1)) || feb29 ? 366 : 365);
    }
    case 2:
      return days / 360;
    case 3:
      return days / 365;
    default:
      return ((y2 - y1) * 360 + (m2 - m1) * 30 + (Math.min(d2, 30) - Math.min(d1, 30))) / 360;
  }
}

export const validFrequency = (f: number) => f === 1 || f === 2 || f === 4;
export const validBasis = (b: number) => b >= 0 && b <= 4;

/** The most periods VDB works through, so that a huge life can't hang the editor. */
export const MAX_PERIODS = 1_000_000;

/** The coupon schedule between settlement and maturity: COUPDAYBS, COUPDAYS, COUPDAYSNC, COUPNUM and the dates. */
export interface Coupons {
  previous: CouponDate;
  next: CouponDate;
  /** Days from the previous coupon to settlement. */
  before: number;
  /** Days in the coupon period. */
  period: number;
  /** Days from settlement to the next coupon. */
  after: number;
  count: number;
}
