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

/**
 * 30/360 days from one date to another in that order (negative going back):
 * by US (NASD) rules, the last day of February and a 31st at the start
 * count as the 30th, as do a 31st at the end after a start on the 30th or
 * later and February's last day after another; or, `european`, every 31st
 * as the 30th.
 */
function thirty(calendar: Calendar, from: number, to: number, european: boolean): number {
  const [y1, m1, d1] = calendar.ymd(from);
  const [y2, m2, d2] = calendar.ymd(to);
  let [a, z] = [d1, d2];
  if (european) [a, z] = [Math.min(a, 30), Math.min(z, 30)];
  else {
    const febEnd1 = m1 === 2 && d1 === daysInMonth(2, y1);
    if (febEnd1 && m2 === 2 && d2 === daysInMonth(2, y2)) z = 30;
    if (z === 31 && a >= 30) z = 30;
    if (a === 31 || febEnd1) a = 30;
  }
  return (y2 - y1) * 360 + (m2 - m1) * 30 + (z - a);
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

/** An argument Excel refuses, with #NUM!. */
export class BondError extends Error {}

function refuse(): never {
  throw new BondError('#NUM!');
}

/**
 * Excel's bond, accrual and discounted-security functions, on a workbook's
 * date system. Each throws a BondError for arguments Excel refuses with #NUM!.
 */
export class Bonds {
  constructor(private readonly calendar: Calendar) {}

  /** Whether serials are dates Excel has: from its first day to the end of 9999. */
  private validDates(...serials: number[]): boolean {
    return serials.every((d) => d >= 0 && this.calendar.ymd(d)[0] <= 9999);
  }

  /** The coupon periods around settlement. */
  coupons(settlement: number, maturity: number, frequency: number, basis: number): Coupons {
    const [s, m, f, b] = [settlement, maturity, frequency, basis].map(Math.trunc);
    if (s >= m || !validFrequency(f) || !validBasis(b) || !this.validDates(s, m)) refuse();
    const calendar = this.calendar;
    const settle = new CouponDate(calendar, s, b);
    const mat = new CouponDate(calendar, m, b);
    const previous = previousCoupon(settle, mat, f);
    const next = nextCoupon(settle, mat, f);
    const before = CouponDate.diff(previous, settle);
    // Actual days in the period for basis 1; otherwise a fixed share of the year.
    const period = b === 1 ? CouponDate.diff(previous, previous.copy().addMonths(12 / f)) : (b === 3 ? 365 : 360) / f;
    const after = b === 0 || b === 4 ? period - before : CouponDate.diff(settle, next);
    const count = (((mat.year - previous.year) * 12 + mat.month - previous.month) * f) / 12;
    return { previous, next, before, period, after, count };
  }

  coupdaybs = (s: number, m: number, f: number, b: number) => this.coupons(s, m, f, b).before;
  coupdays = (s: number, m: number, f: number, b: number) => this.coupons(s, m, f, b).period;
  coupdaysnc = (s: number, m: number, f: number, b: number) => this.coupons(s, m, f, b).after;
  coupncd = (s: number, m: number, f: number, b: number) => this.coupons(s, m, f, b).next.serial(this.calendar);
  coupnum = (s: number, m: number, f: number, b: number) => this.coupons(s, m, f, b).count;
  couppcd = (s: number, m: number, f: number, b: number) => this.coupons(s, m, f, b).previous.serial(this.calendar);

  /** PRICE per 100 face value; in the last coupon period the yield is simple interest. */
  private priceOf(c: Coupons, rate: number, yld: number, redemption: number, frequency: number): number {
    const coupon = (100 * rate) / frequency;
    const accrued = (coupon * c.before) / c.period;
    // The rest of the period, not the actual days to the next coupon, even where those differ (actual/360 and actual/365).
    const dsc = (c.period - c.before) / c.period;
    if (c.count === 1) return (redemption + coupon) / (1 + (dsc * yld) / frequency) - accrued;
    const v = 1 + yld / frequency;
    let price = redemption / v ** (c.count - 1 + dsc);
    for (let k = 1; k <= c.count; k++) price += coupon / v ** (k - 1 + dsc);
    return price - accrued;
  }

  price(settlement: number, maturity: number, rate: number, yld: number, redemption: number, frequency: number, basis: number): number {
    const c = this.coupons(settlement, maturity, frequency, basis);
    if (rate < 0 || yld < 0 || redemption <= 0) refuse();
    return this.priceOf(c, rate, yld, redemption, Math.trunc(frequency));
  }

  yield(settlement: number, maturity: number, rate: number, pr: number, redemption: number, frequency: number, basis: number): number {
    const c = this.coupons(settlement, maturity, frequency, basis);
    if (rate < 0 || pr <= 0 || redemption <= 0) refuse();
    const f = Math.trunc(frequency);
    if (c.count <= 1) {
      const paid = pr / 100 + ((c.before / c.period) * rate) / f;
      return ((redemption / 100 + rate / f - paid) / paid) * ((f * c.period) / c.after);
    }
    // Newton's method on PRICE, as Excel does, for up to 100 steps.
    let y = rate || 0.05;
    let step = 1;
    for (let i = 0; i < 100 && Math.abs(step) > 1e-12; i++) {
      const p = this.priceOf(c, rate, y, redemption, f) - pr;
      const slope = (this.priceOf(c, rate, y + 1e-6, redemption, f) - this.priceOf(c, rate, y - 1e-6, redemption, f)) / 2e-6;
      step = p / slope;
      y -= step;
      if (!Number.isFinite(y) || y <= -f) refuse();
    }
    if (Math.abs(step) > 1e-9) refuse();
    return y;
  }

  /** Macaulay duration in years (or with `modified`, divided by one period's yield): the cash flows' times weighted by their present values. */
  duration(settlement: number, maturity: number, coupon: number, yld: number, frequency: number, basis: number, modified = false): number {
    const c = this.coupons(settlement, maturity, frequency, basis);
    if (coupon < 0 || yld < 0) refuse();
    const f = Math.trunc(frequency);
    const dsc = c.after / c.period;
    const v = 1 + yld / f;
    let weighted = 0;
    let total = 0;
    for (let k = 1; k <= c.count; k++) {
      const t = k - 1 + dsc;
      const pv = ((100 * coupon) / f + (k === c.count ? 100 : 0)) / v ** t;
      weighted += t * pv;
      total += pv;
    }
    const d = weighted / total / f;
    return modified ? d / (1 + yld / f) : d;
  }

  /**
   * ACCRINT as Excel counts it, in coupon periods (stepping from the first
   * interest date) from a starting coupon back to the issue. The starting
   * coupon is the one before the first interest date or, with calc_method
   * TRUE once settlement has passed that date, the last before settlement.
   * The days from it (or from the issue, if later) to settlement are a share
   * of a period (for actual/actual, the one ending on the first interest
   * date), less than nothing when settlement comes first. Each whole period
   * before it counts 1 with calc_method TRUE and nothing with FALSE, and the
   * period the issue falls in counts its share from the issue.
   */
  accrint(issue: number, firstInterest: number, settlement: number, rate: number, par: number, frequency: number, basis: number, fromIssue: boolean): number {
    const [i, first, s, f, b] = [issue, firstInterest, settlement, frequency, basis].map(Math.trunc);
    if (i >= s || rate <= 0 || par <= 0 || !validFrequency(f) || !validBasis(b) || !this.validDates(i, first, s)) refuse();
    const calendar = this.calendar;
    const months = 12 / f;
    const coupon = (k: number) => new CouponDate(calendar, first, b).addMonths(k * months).serial(calendar);
    // Days from one date to another as the basis counts them, negative going back.
    const days = (from: number, to: number) => (b === 0 || b === 4 ? thirty(calendar, from, to, b === 4) : to - from);
    const period = (lo: number, hi: number) => (b === 1 ? hi - lo : (b === 3 ? 365 : 360) / f);
    let k = -1;
    if (fromIssue && s > first) for (k = 0; coupon(k + 1) < s; k++);
    let share = days(Math.max(i, coupon(k)), s) / period(coupon(-1), first);
    for (; coupon(k) > i; k--) {
      const [lo, hi] = [coupon(k - 1), coupon(k)];
      share += lo >= i ? (fromIssue ? 1 : 0) : days(i, hi) / period(lo, hi);
    }
    return ((par * rate) / f) * share;
  }

  accrintm(issue: number, settlement: number, rate: number, par: number, basis: number): number {
    const [i, s, b] = [issue, settlement, basis].map(Math.trunc);
    if (i >= s || rate <= 0 || par <= 0 || !validBasis(b) || !this.validDates(i, s)) refuse();
    return par * rate * yearFrac(this.calendar, i, s, b);
  }

  /** The functions for discounted securities, each a formula in the fraction of a year to maturity. */
  private discounted(settlement: number, maturity: number, a: number, b: number, basis: number, formula: (years: number, a: number, b: number) => number | undefined): number {
    const [s, m, base] = [settlement, maturity, basis].map(Math.trunc);
    if (s >= m || a <= 0 || b <= 0 || !validBasis(base) || !this.validDates(s, m)) refuse();
    return formula(yearFrac(this.calendar, s, m, base), a, b) ?? refuse();
  }

  disc = (s: number, m: number, pr: number, redemption: number, basis: number) => this.discounted(s, m, pr, redemption, basis, (years, p, r) => (r - p) / r / years);
  intrate = (s: number, m: number, investment: number, redemption: number, basis: number) => this.discounted(s, m, investment, redemption, basis, (years, i, r) => (r - i) / i / years);
  received = (s: number, m: number, investment: number, discount: number, basis: number) =>
    this.discounted(s, m, investment, discount, basis, (years, i, d) => {
      const left = 1 - d * years;
      return left > 0 ? i / left : undefined;
    });
  pricedisc = (s: number, m: number, discount: number, redemption: number, basis: number) => this.discounted(s, m, discount, redemption, basis, (years, d, r) => r - d * r * years);
  yielddisc = (s: number, m: number, pr: number, redemption: number, basis: number) => this.discounted(s, m, pr, redemption, basis, (years, p, r) => (r - p) / p / years);

  pricemat(settlement: number, maturity: number, issue: number, rate: number, yld: number, basis: number): number {
    const [s, m, i, b] = [settlement, maturity, issue, basis].map(Math.trunc);
    if (s >= m || rate < 0 || yld < 0 || !validBasis(b) || !this.validDates(s, m, i)) refuse();
    const [dim, dsm, a] = [yearFrac(this.calendar, i, m, b), yearFrac(this.calendar, s, m, b), yearFrac(this.calendar, i, s, b)];
    return (100 + dim * rate * 100) / (1 + dsm * yld) - a * rate * 100;
  }

  yieldmat(settlement: number, maturity: number, issue: number, rate: number, pr: number, basis: number): number {
    const [s, m, i, b] = [settlement, maturity, issue, basis].map(Math.trunc);
    if (s >= m || rate < 0 || pr <= 0 || !validBasis(b) || !this.validDates(s, m, i)) refuse();
    const [dim, dsm, a] = [yearFrac(this.calendar, i, m, b), yearFrac(this.calendar, s, m, b), yearFrac(this.calendar, i, s, b)];
    const paid = pr / 100 + a * rate;
    return (1 + dim * rate - paid) / paid / dsm;
  }
}

/** VDB: depreciation between two periods, declining by `factor`, switching to straight-line when that gives more unless `noSwitch`. */
export function vdb(cost: number, salvage: number, life: number, start: number, end: number, factor: number, noSwitch: boolean): number {
  if (start < 0 || end < start || end > life || life <= 0 || cost < 0 || salvage < 0 || salvage > cost || factor <= 0 || end > MAX_PERIODS) refuse();
  return noSwitch ? decliningOnly(cost, salvage, life, start, end, factor) : switching(cost, salvage, life, start, end, factor);
}

/** Double-declining (or factor-declining) depreciation for one whole period. */
export function ddb(cost: number, salvage: number, life: number, period: number, factor: number): number {
  let rate = factor / life;
  let before: number;
  if (rate >= 1) {
    rate = 1;
    before = period === 1 ? cost : 0;
  } else {
    before = cost * (1 - rate) ** (period - 1);
  }
  const after = cost * (1 - rate) ** period;
  return Math.max(after < salvage ? before - salvage : before - after, 0);
}

/** Depreciation over the first `periods` periods (the last one possibly partial), switching to straight-line once that is larger. */
function depreciation(cost: number, salvage: number, life: number, remaining: number, periods: number, factor: number): number {
  const end = Math.ceil(periods);
  let total = 0;
  let left = cost - salvage;
  let straight = 0;
  let switched = false;
  for (let i = 1; i <= end; i++) {
    let term: number;
    if (switched) {
      term = straight;
    } else {
      const declining = ddb(cost, salvage, life, i, factor);
      straight = left / (remaining - (i - 1));
      if (straight > declining) {
        term = straight;
        switched = true;
      } else {
        term = declining;
        left -= declining;
      }
    }
    total += i === end ? term * (periods + 1 - end) : term;
  }
  return total;
}

/** VDB with the switch to straight-line, for any fractional start and end. */
function switching(cost: number, salvage: number, life: number, start: number, end: number, factor: number): number {
  const s = Math.floor(start);
  const e = Math.ceil(end);
  let part = 0;
  if (start !== s) {
    const value = cost - depreciation(cost, salvage, life, life, s, factor);
    part += (start - s) * depreciation(value, salvage, life, life - s, 1, factor);
  }
  if (end !== e) {
    const value = cost - depreciation(cost, salvage, life, life, e - 1, factor);
    part += (e - end) * depreciation(value, salvage, life, life - e + 1, 1, factor);
  }
  const value = cost - depreciation(cost, salvage, life, life, s, factor);
  return depreciation(value, salvage, life, life - s, e - s, factor) - part;
}

/** VDB staying with the declining balance, partial first and last periods prorated. */
function decliningOnly(cost: number, salvage: number, life: number, start: number, end: number, factor: number): number {
  const s = Math.floor(start);
  const e = Math.ceil(end);
  let total = 0;
  for (let i = s + 1; i <= e; i++) {
    let term = ddb(cost, salvage, life, i, factor);
    if (i === s + 1) term *= Math.min(end, s + 1) - start;
    else if (i === e) term *= end + 1 - e;
    total += term;
  }
  return total;
}
