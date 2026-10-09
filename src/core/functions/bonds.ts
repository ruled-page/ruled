import { ErrorType, FunctionArgumentType, type ImplementedFunctions } from 'hyperformula';
import { ExcelPlugin, error, type Call, type State } from './common';
import { BondError, Bonds, Calendar, vdb } from '../bondMath';

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

  /** Runs one of the shared bond functions, with #NUM! for arguments Excel refuses. */
  private bond<A extends unknown[]>(ast: Call, state: State, name: string, f: (b: Bonds, ...args: A) => number) {
    const bonds = new Bonds(new Calendar((this.config as unknown as { nullDate: { year: number; month: number; day: number } }).nullDate));
    return this.run(ast, state, name, (...args: A) => {
      try {
        return f(bonds, ...args);
      } catch (e) {
        if (e instanceof BondError) return error(ErrorType.NUM);
        throw e;
      }
    });
  }

  coupdaybs(ast: Call, state: State) {
    return this.bond(ast, state, 'COUPDAYBS', (b: Bonds, s: number, m: number, f: number, basis: number) => b.coupdaybs(s, m, f, basis));
  }

  coupdays(ast: Call, state: State) {
    return this.bond(ast, state, 'COUPDAYS', (b: Bonds, s: number, m: number, f: number, basis: number) => b.coupdays(s, m, f, basis));
  }

  coupdaysnc(ast: Call, state: State) {
    return this.bond(ast, state, 'COUPDAYSNC', (b: Bonds, s: number, m: number, f: number, basis: number) => b.coupdaysnc(s, m, f, basis));
  }

  coupncd(ast: Call, state: State) {
    return this.bond(ast, state, 'COUPNCD', (b: Bonds, s: number, m: number, f: number, basis: number) => b.coupncd(s, m, f, basis));
  }

  coupnum(ast: Call, state: State) {
    return this.bond(ast, state, 'COUPNUM', (b: Bonds, s: number, m: number, f: number, basis: number) => b.coupnum(s, m, f, basis));
  }

  couppcd(ast: Call, state: State) {
    return this.bond(ast, state, 'COUPPCD', (b: Bonds, s: number, m: number, f: number, basis: number) => b.couppcd(s, m, f, basis));
  }

  price(ast: Call, state: State) {
    return this.bond(ast, state, 'PRICE', (b: Bonds, s: number, m: number, rate: number, yld: number, redemption: number, f: number, basis: number) => b.price(s, m, rate, yld, redemption, f, basis));
  }

  yield(ast: Call, state: State) {
    return this.bond(ast, state, 'YIELD', (b: Bonds, s: number, m: number, rate: number, pr: number, redemption: number, f: number, basis: number) => b.yield(s, m, rate, pr, redemption, f, basis));
  }

  duration(ast: Call, state: State) {
    return this.bond(ast, state, 'DURATION', (b: Bonds, s: number, m: number, coupon: number, yld: number, f: number, basis: number) => b.duration(s, m, coupon, yld, f, basis));
  }

  mduration(ast: Call, state: State) {
    return this.bond(ast, state, 'MDURATION', (b: Bonds, s: number, m: number, coupon: number, yld: number, f: number, basis: number) => b.duration(s, m, coupon, yld, f, basis, true));
  }

  accrint(ast: Call, state: State) {
    return this.bond(ast, state, 'ACCRINT', (b: Bonds, issue: number, first: number, s: number, rate: number, par: number, f: number, basis: number, fromIssue: boolean) => b.accrint(issue, first, s, rate, par, f, basis, fromIssue));
  }

  accrintm(ast: Call, state: State) {
    return this.bond(ast, state, 'ACCRINTM', (b: Bonds, issue: number, s: number, rate: number, par: number, basis: number) => b.accrintm(issue, s, rate, par, basis));
  }

  disc(ast: Call, state: State) {
    return this.bond(ast, state, 'DISC', (b: Bonds, s: number, m: number, x: number, y: number, basis: number) => b.disc(s, m, x, y, basis));
  }

  intrate(ast: Call, state: State) {
    return this.bond(ast, state, 'INTRATE', (b: Bonds, s: number, m: number, x: number, y: number, basis: number) => b.intrate(s, m, x, y, basis));
  }

  received(ast: Call, state: State) {
    return this.bond(ast, state, 'RECEIVED', (b: Bonds, s: number, m: number, x: number, y: number, basis: number) => b.received(s, m, x, y, basis));
  }

  pricedisc(ast: Call, state: State) {
    return this.bond(ast, state, 'PRICEDISC', (b: Bonds, s: number, m: number, x: number, y: number, basis: number) => b.pricedisc(s, m, x, y, basis));
  }

  yielddisc(ast: Call, state: State) {
    return this.bond(ast, state, 'YIELDDISC', (b: Bonds, s: number, m: number, x: number, y: number, basis: number) => b.yielddisc(s, m, x, y, basis));
  }

  pricemat(ast: Call, state: State) {
    return this.bond(ast, state, 'PRICEMAT', (b: Bonds, s: number, m: number, issue: number, rate: number, yld: number, basis: number) => b.pricemat(s, m, issue, rate, yld, basis));
  }

  yieldmat(ast: Call, state: State) {
    return this.bond(ast, state, 'YIELDMAT', (b: Bonds, s: number, m: number, issue: number, rate: number, pr: number, basis: number) => b.yieldmat(s, m, issue, rate, pr, basis));
  }

  vdb(ast: Call, state: State) {
    return this.bond(ast, state, 'VDB', (_b: Bonds, cost: number, salvage: number, life: number, start: number, end: number, factor: number, noSwitch: boolean) => vdb(cost, salvage, life, start, end, factor, noSwitch));
  }
}
