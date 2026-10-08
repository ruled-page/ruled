import { CellError, ErrorType, FunctionArgumentType, type ImplementedFunctions } from 'hyperformula';
import { ExcelPlugin, error, type Call, type State } from './common';

const DAY_MS = 86_400_000;

const isLeap = (y: number) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
const daysInMonth = (m: number, y: number) => [31, isLeap(y) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1];

/** Serial numbers to calendar dates and back, from the workbook's date system. */
class Calendar {
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
class CouponDate {
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
function previousCoupon(settle: CouponDate, maturity: CouponDate, frequency: number): CouponDate {
  const d = maturity.copy().setYear(settle.year);
  if (d.compare(settle) < 0) d.setYear(d.year + 1);
  while (d.compare(settle) > 0) d.addMonths(-12 / frequency);
  return d;
}

/** The first coupon date after settlement. */
function nextCoupon(settle: CouponDate, maturity: CouponDate, frequency: number): CouponDate {
  const d = maturity.copy().setYear(settle.year);
  if (d.compare(settle) > 0) d.setYear(d.year - 1);
  while (d.compare(settle) <= 0) d.addMonths(12 / frequency);
  return d;
}

/** The fraction of a year between two serials in a basis, as YEARFRAC works it out. */
function yearFrac(calendar: Calendar, start: number, end: number, basis: number): number {
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

const validFrequency = (f: number) => f === 1 || f === 2 || f === 4;
const validBasis = (b: number) => b >= 0 && b <= 4;

/** The coupon schedule between settlement and maturity: COUPDAYBS, COUPDAYS, COUPDAYSNC, COUPNUM and the dates. */
interface Coupons {
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

const date = { argumentType: FunctionArgumentType.NUMBER };
const number = { argumentType: FunctionArgumentType.NUMBER };
const basisParameter = { argumentType: FunctionArgumentType.NUMBER, defaultValue: 0, emptyAsDefault: true };
const couponParameters = [date, date, number, basisParameter];

export class BondFunctions extends ExcelPlugin {
  static implementedFunctions: ImplementedFunctions = {
    COUPDAYBS: { method: 'coupdaybs', parameters: couponParameters },
    COUPDAYS: { method: 'coupdays', parameters: couponParameters },
    COUPDAYSNC: { method: 'coupdaysnc', parameters: couponParameters },
    COUPNCD: { method: 'coupncd', parameters: couponParameters },
    COUPNUM: { method: 'coupnum', parameters: couponParameters },
    COUPPCD: { method: 'couppcd', parameters: couponParameters },
    PRICE: { method: 'price', parameters: [date, date, number, number, number, number, basisParameter] },
    YIELD: { method: 'yield', parameters: [date, date, number, number, number, number, basisParameter] },
    DURATION: { method: 'duration', parameters: [date, date, number, number, number, basisParameter] },
    MDURATION: { method: 'mduration', parameters: [date, date, number, number, number, basisParameter] },
    ACCRINT: {
      method: 'accrint',
      parameters: [
        date,
        date,
        date,
        number,
        { argumentType: FunctionArgumentType.NUMBER, defaultValue: 1000, emptyAsDefault: true },
        number,
        basisParameter,
        { argumentType: FunctionArgumentType.BOOLEAN, defaultValue: true, emptyAsDefault: true },
      ],
    },
    ACCRINTM: { method: 'accrintm', parameters: [date, date, number, { argumentType: FunctionArgumentType.NUMBER, defaultValue: 1000, emptyAsDefault: true }, basisParameter] },
    DISC: { method: 'disc', parameters: [date, date, number, number, basisParameter] },
    INTRATE: { method: 'intrate', parameters: [date, date, number, number, basisParameter] },
    RECEIVED: { method: 'received', parameters: [date, date, number, number, basisParameter] },
    PRICEDISC: { method: 'pricedisc', parameters: [date, date, number, number, basisParameter] },
    YIELDDISC: { method: 'yielddisc', parameters: [date, date, number, number, basisParameter] },
    PRICEMAT: { method: 'pricemat', parameters: [date, date, date, number, number, basisParameter] },
    YIELDMAT: { method: 'yieldmat', parameters: [date, date, date, number, number, basisParameter] },
    VDB: {
      method: 'vdb',
      parameters: [
        number,
        number,
        number,
        number,
        number,
        { argumentType: FunctionArgumentType.NUMBER, defaultValue: 2, emptyAsDefault: true },
        { argumentType: FunctionArgumentType.BOOLEAN, defaultValue: false, emptyAsDefault: true },
      ],
    },
  };

  private get calendar(): Calendar {
    return new Calendar((this.config as unknown as { nullDate: { year: number; month: number; day: number } }).nullDate);
  }

  /** The coupon periods around settlement, or #NUM! for arguments Excel refuses. */
  private coupons(settlement: number, maturity: number, frequency: number, basis: number): Coupons | CellError {
    const [s, m, f, b] = [settlement, maturity, frequency, basis].map(Math.trunc);
    if (s >= m || !validFrequency(f) || !validBasis(b)) return error(ErrorType.NUM);
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

  private couponFunction(ast: Call, state: State, name: string, pick: (c: Coupons) => number) {
    return this.run(ast, state, name, (settlement: number, maturity: number, frequency: number, basis: number) => {
      const c = this.coupons(settlement, maturity, frequency, basis);
      return c instanceof CellError ? c : pick(c);
    });
  }

  coupdaybs(ast: Call, state: State) {
    return this.couponFunction(ast, state, 'COUPDAYBS', (c) => c.before);
  }

  coupdays(ast: Call, state: State) {
    return this.couponFunction(ast, state, 'COUPDAYS', (c) => c.period);
  }

  coupdaysnc(ast: Call, state: State) {
    return this.couponFunction(ast, state, 'COUPDAYSNC', (c) => c.after);
  }

  coupncd(ast: Call, state: State) {
    return this.couponFunction(ast, state, 'COUPNCD', (c) => c.next.serial(this.calendar));
  }

  coupnum(ast: Call, state: State) {
    return this.couponFunction(ast, state, 'COUPNUM', (c) => c.count);
  }

  couppcd(ast: Call, state: State) {
    return this.couponFunction(ast, state, 'COUPPCD', (c) => c.previous.serial(this.calendar));
  }

  /** PRICE per 100 face value; in the last coupon period the yield is simple interest. */
  private priceOf(c: Coupons, rate: number, yld: number, redemption: number, frequency: number): number {
    const coupon = (100 * rate) / frequency;
    const accrued = (coupon * c.before) / c.period;
    const dsc = c.after / c.period;
    if (c.count === 1) return (redemption + coupon) / (1 + (dsc * yld) / frequency) - accrued;
    const v = 1 + yld / frequency;
    let price = redemption / v ** (c.count - 1 + dsc);
    for (let k = 1; k <= c.count; k++) price += coupon / v ** (k - 1 + dsc);
    return price - accrued;
  }

  price(ast: Call, state: State) {
    return this.run(ast, state, 'PRICE', (settlement: number, maturity: number, rate: number, yld: number, redemption: number, frequency: number, basis: number) => {
      const c = this.coupons(settlement, maturity, frequency, basis);
      if (c instanceof CellError) return c;
      if (rate < 0 || yld < 0 || redemption <= 0) return error(ErrorType.NUM);
      return this.priceOf(c, rate, yld, redemption, Math.trunc(frequency));
    });
  }

  yield(ast: Call, state: State) {
    return this.run(ast, state, 'YIELD', (settlement: number, maturity: number, rate: number, pr: number, redemption: number, frequency: number, basis: number) => {
      const c = this.coupons(settlement, maturity, frequency, basis);
      if (c instanceof CellError) return c;
      if (rate < 0 || pr <= 0 || redemption <= 0) return error(ErrorType.NUM);
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
        if (!Number.isFinite(y) || y <= -f) return error(ErrorType.NUM);
      }
      return Math.abs(step) > 1e-9 ? error(ErrorType.NUM) : y;
    });
  }

  /** Macaulay duration in years: the cash flows' times weighted by their present values. */
  private durationOf(c: Coupons, coupon: number, yld: number, frequency: number): number {
    const dsc = c.after / c.period;
    const v = 1 + yld / frequency;
    let weighted = 0;
    let total = 0;
    for (let k = 1; k <= c.count; k++) {
      const t = k - 1 + dsc;
      const flow = (100 * coupon) / frequency + (k === c.count ? 100 : 0);
      const pv = flow / v ** t;
      weighted += t * pv;
      total += pv;
    }
    return weighted / total / frequency;
  }

  duration(ast: Call, state: State) {
    return this.run(ast, state, 'DURATION', this.durationFunction(false));
  }

  mduration(ast: Call, state: State) {
    return this.run(ast, state, 'MDURATION', this.durationFunction(true));
  }

  private durationFunction(modified: boolean) {
    return (settlement: number, maturity: number, coupon: number, yld: number, frequency: number, basis: number) => {
      const c = this.coupons(settlement, maturity, frequency, basis);
      if (c instanceof CellError) return c;
      if (coupon < 0 || yld < 0) return error(ErrorType.NUM);
      const f = Math.trunc(frequency);
      const d = this.durationOf(c, coupon, yld, f);
      return modified ? d / (1 + yld / f) : d;
    };
  }

  accrint(ast: Call, state: State) {
    return this.run(
      ast,
      state,
      'ACCRINT',
      (issue: number, firstInterest: number, settlement: number, rate: number, par: number, frequency: number, basis: number, fromIssue: boolean) => {
        const [i, first, s, f, b] = [issue, firstInterest, settlement, frequency, basis].map(Math.trunc);
        if (i >= s || rate <= 0 || par <= 0 || !validFrequency(f) || !validBasis(b)) return error(ErrorType.NUM);
        const calendar = this.calendar;
        // Coupon dates run from the first interest date in steps of 12/f months, back and forward.
        const anchor = new CouponDate(calendar, first, b);
        const at = (k: number) => anchor.copy().addMonths((k * 12) / f).serial(calendar);
        // From the issue date, or with calc_method FALSE from the first interest date once that has passed.
        const start = !fromIssue && s > first ? first : i;
        let k = 0;
        while (at(k) > start) k--;
        while (at(k + 1) <= start) k++;
        // Each coupon period counts its share of a coupon: days accrued over the period's length.
        let periods = 0;
        for (; at(k) < s; k++) {
          const [lo, hi] = [at(k), at(k + 1)];
          const [from, to] = [Math.max(start, lo), Math.min(s, hi)];
          const days = b === 0 || b === 4 ? CouponDate.diff(new CouponDate(calendar, from, b), new CouponDate(calendar, to, b)) : to - from;
          periods += days / (b === 1 ? hi - lo : (b === 3 ? 365 : 360) / f);
        }
        return ((par * rate) / f) * periods;
      },
    );
  }

  accrintm(ast: Call, state: State) {
    return this.run(ast, state, 'ACCRINTM', (issue: number, settlement: number, rate: number, par: number, basis: number) => {
      const [i, s, b] = [issue, settlement, basis].map(Math.trunc);
      if (i >= s || rate <= 0 || par <= 0 || !validBasis(b)) return error(ErrorType.NUM);
      return par * rate * yearFrac(this.calendar, i, s, b);
    });
  }

  /** The functions for discounted securities, each a formula in the fraction of a year to maturity. */
  private discounted(ast: Call, state: State, name: string, formula: (years: number, a: number, b: number) => number | undefined) {
    return this.run(ast, state, name, (settlement: number, maturity: number, a: number, b: number, basis: number) => {
      const [s, m, base] = [settlement, maturity, basis].map(Math.trunc);
      if (s >= m || a <= 0 || b <= 0 || !validBasis(base)) return error(ErrorType.NUM);
      return formula(yearFrac(this.calendar, s, m, base), a, b) ?? error(ErrorType.NUM);
    });
  }

  disc(ast: Call, state: State) {
    return this.discounted(ast, state, 'DISC', (years, pr, redemption) => (redemption - pr) / redemption / years);
  }

  intrate(ast: Call, state: State) {
    return this.discounted(ast, state, 'INTRATE', (years, investment, redemption) => (redemption - investment) / investment / years);
  }

  received(ast: Call, state: State) {
    return this.discounted(ast, state, 'RECEIVED', (years, investment, discount) => {
      const left = 1 - discount * years;
      return left > 0 ? investment / left : undefined;
    });
  }

  pricedisc(ast: Call, state: State) {
    return this.discounted(ast, state, 'PRICEDISC', (years, discount, redemption) => redemption - discount * redemption * years);
  }

  yielddisc(ast: Call, state: State) {
    return this.discounted(ast, state, 'YIELDDISC', (years, pr, redemption) => (redemption - pr) / pr / years);
  }

  pricemat(ast: Call, state: State) {
    return this.run(ast, state, 'PRICEMAT', (settlement: number, maturity: number, issue: number, rate: number, yld: number, basis: number) => {
      const [s, m, i, b] = [settlement, maturity, issue, basis].map(Math.trunc);
      if (s >= m || rate < 0 || yld < 0 || !validBasis(b)) return error(ErrorType.NUM);
      const calendar = this.calendar;
      const [dim, dsm, a] = [yearFrac(calendar, i, m, b), yearFrac(calendar, s, m, b), yearFrac(calendar, i, s, b)];
      return (100 + dim * rate * 100) / (1 + dsm * yld) - a * rate * 100;
    });
  }

  yieldmat(ast: Call, state: State) {
    return this.run(ast, state, 'YIELDMAT', (settlement: number, maturity: number, issue: number, rate: number, pr: number, basis: number) => {
      const [s, m, i, b] = [settlement, maturity, issue, basis].map(Math.trunc);
      if (s >= m || rate < 0 || pr <= 0 || !validBasis(b)) return error(ErrorType.NUM);
      const calendar = this.calendar;
      const [dim, dsm, a] = [yearFrac(calendar, i, m, b), yearFrac(calendar, s, m, b), yearFrac(calendar, i, s, b)];
      const paid = pr / 100 + a * rate;
      return (1 + dim * rate - paid) / paid / dsm;
    });
  }

  vdb(ast: Call, state: State) {
    return this.run(ast, state, 'VDB', (cost: number, salvage: number, life: number, start: number, end: number, factor: number, noSwitch: boolean) => {
      if (start < 0 || end < start || end > life || life <= 0 || cost < 0 || salvage < 0 || salvage > cost || factor <= 0) return error(ErrorType.NUM);
      return noSwitch ? decliningOnly(cost, salvage, life, start, end, factor) : switching(cost, salvage, life, start, end, factor);
    });
  }
}

/** Double-declining (or factor-declining) depreciation for one whole period. */
function ddb(cost: number, salvage: number, life: number, period: number, factor: number): number {
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
