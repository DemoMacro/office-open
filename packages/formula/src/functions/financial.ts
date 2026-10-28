import type { FormulaEvaluationContext } from "../context";
import { FormulaValueError } from "../errors";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { numbers, scalarNumber, scalarText } from "./utils";

type FinancialArgs = Parameters<FormulaFunctionImplementation>[0];
type FinancialArg = FinancialArgs[number];

const millisecondsPerDay = 86_400_000;

function number(value: FormulaFunctionArgument): number {
  const output = scalarNumber(value);
  if (!Number.isFinite(output)) throw new FormulaValueError("#NUM!");
  return output;
}

type FormulaFunctionArgument = Parameters<FormulaFunctionImplementation>[0][number];

function dateValue(value: FormulaFunctionArgument, context: FormulaEvaluationContext): number {
  const serial = number(value);
  return (context.dateSystem ?? context.workbook.dateSystem) === "1904" ? serial + 1462 : serial;
}

function integer(value: number, allowZero = true): number {
  const output = Math.trunc(value);
  if (output !== value || (!allowZero && output === 0)) throw new FormulaValueError("#NUM!");
  return output;
}

function requireBetween(value: number, min: number, max: number): number {
  if (value < min || value > max) throw new FormulaValueError("#NUM!");
  return value;
}

function rate(value: FormulaFunctionArgument): number {
  const output = number(value);
  if (output <= -1) throw new FormulaValueError("#NUM!");
  return output;
}

function solveByBisection(evaluate: (value: number) => number, low: number, high: number): number {
  let lower = low;
  let upper = high;
  let lowerValue = evaluate(lower);
  let upperValue = evaluate(upper);
  if (!Number.isFinite(lowerValue) || !Number.isFinite(upperValue)) return Number.NaN;
  if (lowerValue === 0) return lower;
  if (upperValue === 0) return upper;
  if (lowerValue > 0 === upperValue > 0) return Number.NaN;
  for (let iteration = 0; iteration < 200; iteration += 1) {
    const middle = (lower + upper) / 2;
    const middleValue = evaluate(middle);
    if (middleValue === 0 || (upper - lower) / 2 < 1e-12) return middle;
    if (lowerValue > 0 === middleValue > 0) {
      lower = middle;
      lowerValue = middleValue;
    } else {
      upper = middle;
      upperValue = middleValue;
    }
  }
  return (lower + upper) / 2;
}

function solveContinuous(evaluate: (value: number) => number, guess: number): number {
  let current = guess <= -0.999999 ? -0.99 : guess;
  for (let iteration = 0; iteration < 80; iteration += 1) {
    const currentFactor = evaluate(current);
    if (!Number.isFinite(currentFactor)) break;
    const delta = Math.max(Math.abs(current) * 1e-7, 1e-7);
    const derivative = (evaluate(current + delta) - currentFactor) / delta;
    if (!Number.isFinite(derivative) || derivative === 0) break;
    const next = current - currentFactor / derivative;
    if (next <= -1) {
      current = (current - 1) / 2;
      continue;
    }
    if (Math.abs(next - current) < 1e-10) return next;
    current = next;
  }

  const nearNegativeRates = [-0.999999, -0.99, -0.9, -0.75, -0.5, -0.25, -0.1, -0.01, -0.001];
  for (const low of nearNegativeRates) {
    const high = low === -0.001 ? 0 : low + (low === -0.999999 ? 0.009 : 0.01);
    const root = solveByBisection(evaluate, low, high);
    if (Number.isFinite(root)) return root;
  }
  for (let low = 0; low < 10; low += 0.05) {
    const root = solveByBisection(evaluate, low, low + 0.05);
    if (Number.isFinite(root)) return root;
  }
  for (let low = 10; low < 100; low += 0.5) {
    const root = solveByBisection(evaluate, low, low + 0.5);
    if (Number.isFinite(root)) return root;
  }
  throw new FormulaValueError("#NUM!");
}

function annuityPayment(
  rateValue: number,
  periods: number,
  present: number,
  future: number,
  type: number,
): number {
  if (periods === 0) return Number.NaN;
  if (rateValue === 0) return -(present + future) / periods;
  const growth = 1 + rateValue;
  return (
    (-(present * growth ** periods + future) * rateValue) /
    ((growth ** periods - 1) * (1 + rateValue * type))
  );
}

function annuityInterest(
  rateValue: number,
  period: number,
  periods: number,
  present: number,
  future: number,
  type: number,
): number {
  if (period < 1 || period > periods) return Number.NaN;
  const payment = annuityPayment(rateValue, periods, present, future, type);
  if (type === 1 && period === 1) return 0;
  const growth = 1 + rateValue;
  const exponent = type === 1 ? period - 1 : period;
  const balance = present * growth ** exponent + (payment * (growth ** exponent - 1)) / rateValue;
  return rateValue === 0 ? 0 : -balance * rateValue;
}

function toUtcDate(serial: number): Date {
  const date = new Date(Math.round((Math.trunc(serial) - 25569) * millisecondsPerDay));
  if (Number.isNaN(date.getTime()) || date.getUTCFullYear() < 1900 || date.getUTCFullYear() > 9999)
    throw new FormulaValueError("#NUM!");
  return date;
}

function dayDifference(start: number, end: number): number {
  return Math.round((toUtcDate(end).getTime() - toUtcDate(start).getTime()) / millisecondsPerDay);
}

function isLeapYear(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

function yearDays(year: number): number {
  return isLeapYear(year) ? 366 : 365;
}

function thirty360(start: number, end: number, european: boolean): number {
  const startDate = toUtcDate(start);
  const endDate = toUtcDate(end);
  const startDayNormalized = startDate.getUTCDate() === 31 ? 30 : startDate.getUTCDate();
  const endDay = european && endDate.getUTCDate() === 31 ? 30 : endDate.getUTCDate();
  const endDayNormalized = startDayNormalized === 30 && endDay === 31 ? 30 : endDay;
  return (
    (endDate.getUTCFullYear() - startDate.getUTCFullYear()) * 360 +
    (endDate.getUTCMonth() - startDate.getUTCMonth()) * 30 +
    endDayNormalized -
    startDayNormalized
  );
}

function yearFraction(start: number, end: number, basis: number): number {
  requireBetween(basis, 0, 4);
  if (start > end) throw new FormulaValueError("#NUM!");
  if (basis === 0 || basis === 4) return thirty360(start, end, basis === 4) / 360;
  const days = dayDifference(start, end);
  if (basis === 1) return actualActual(start, end);
  return basis === 2 ? days / 360 : days / 365;
}

function actualActual(startSerial: number, endSerial: number): number {
  const startDate = toUtcDate(startSerial);
  const endDate = toUtcDate(endSerial);
  if (endDate.getTime() <= startDate.getTime()) return 0;
  if (startDate.getUTCFullYear() === endDate.getUTCFullYear())
    return dayDifference(startSerial, endSerial) / yearDays(startDate.getUTCFullYear());
  let total = 0;
  let cursor = startDate;
  while (cursor.getUTCFullYear() < endDate.getUTCFullYear()) {
    const yearEnd = new Date(Date.UTC(cursor.getUTCFullYear() + 1, 0, 1));
    total +=
      (yearEnd.getTime() - cursor.getTime()) /
      millisecondsPerDay /
      yearDays(cursor.getUTCFullYear());
    cursor = yearEnd;
  }
  return (
    total +
    (endDate.getTime() - cursor.getTime()) / millisecondsPerDay / yearDays(endDate.getUTCFullYear())
  );
}

function addCouponMonths(date: Date, months: number): Date {
  const target = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + months, 1));
  target.setUTCDate(
    Math.min(
      date.getUTCDate(),
      new Date(Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0)).getUTCDate(),
    ),
  );
  return target;
}

function couponSchedule(settlement: number, maturity: number, frequency: number) {
  const settlementDate = toUtcDate(settlement);
  const maturityDate = toUtcDate(maturity);
  if (settlementDate.getTime() >= maturityDate.getTime()) throw new FormulaValueError("#NUM!");
  let previous = maturityDate;
  while (previous.getTime() > settlementDate.getTime()) {
    previous = addCouponMonths(previous, -12 / frequency);
  }
  const next = addCouponMonths(previous, 12 / frequency);
  let couponCount = 0;
  for (
    let coupon = next;
    coupon.getTime() < maturityDate.getTime();
    coupon = addCouponMonths(coupon, 12 / frequency)
  ) {
    couponCount += 1;
  }
  return { previous, next, couponCount };
}

function couponFraction(start: Date, end: Date, basis: number): number {
  const startSerial = start.getTime() / millisecondsPerDay + 25569;
  const endSerial = end.getTime() / millisecondsPerDay + 25569;
  if (basis === 0 || basis === 4) return thirty360(startSerial, endSerial, basis === 4) / 360;
  const days = dayDifference(startSerial, endSerial);
  if (basis === 2) return days / 360;
  if (basis === 3) return days / 365;
  return actualActual(startSerial, endSerial);
}

function couponValues(
  settlement: FinancialArg,
  maturity: FinancialArg,
  frequency: FinancialArg,
  basis?: FinancialArg,
): {
  settlement: number;
  maturity: number;
  frequency: number;
  basis: number;
} {
  const settlementValue = number(settlement);
  const maturityValue = number(maturity);
  const frequencyValue = integer(number(frequency), false);
  if (frequencyValue !== 1 && frequencyValue !== 2 && frequencyValue !== 4)
    throw new FormulaValueError("#NUM!");
  const basisValue = integer(number(basis ?? 0));
  requireBetween(basisValue, 0, 4);
  return {
    settlement: settlementValue,
    maturity: maturityValue,
    frequency: frequencyValue,
    basis: basisValue,
  };
}

function couponPrice(
  settlement: number,
  maturity: number,
  rateValue: number,
  yieldValue: number,
  redemption: number,
  frequency: number,
  basis: number,
): { price: number; accrued: number; fractionalPeriod: number; couponCount: number } {
  if (yieldValue <= -1) throw new FormulaValueError("#NUM!");
  const { previous, next, couponCount } = couponSchedule(settlement, maturity, frequency);
  const periodFraction = couponFraction(previous, next, basis);
  const fractionalPeriod = couponFraction(previous, toUtcDate(settlement), basis) / periodFraction;
  const coupon = (redemption * rateValue) / frequency;
  const periodicYield = yieldValue / frequency + 1;
  let presentValue = redemption / periodicYield ** (couponCount - 1 + fractionalPeriod);
  for (let index = 0; index < couponCount; index += 1) {
    presentValue += coupon / periodicYield ** (index + fractionalPeriod);
  }
  const accrued = coupon * fractionalPeriod;
  return { price: presentValue - accrued, accrued, fractionalPeriod, couponCount };
}

function matrixNumbers(value: FinancialArg): number[] {
  return numbers(value);
}

function flatSeries(args: FinancialArgs, start: number): number[] {
  return args.slice(start).flatMap(matrixNumbers);
}

function depreciationAmount(
  cost: number,
  salvage: number,
  life: number,
  start: number,
  end: number,
  factor: number,
  noSwitch: boolean,
): number {
  if (life <= 0 || start < 0 || end <= start || factor <= 0 || cost < salvage)
    throw new FormulaValueError("#NUM!");
  const rateValue = Math.min(factor / life, 1);
  let accumulated = 0;
  let total = 0;
  for (let period = 1; period <= Math.ceil(end); period += 1) {
    const remaining = Math.max(life - period + 1, 1);
    const depreciable = cost - salvage - accumulated;
    if (depreciable <= 0) break;
    const declining = Math.min((cost - accumulated) * rateValue, depreciable);
    const straightLine = depreciable / remaining;
    let amount = noSwitch ? declining : Math.max(declining, straightLine);
    const segmentStart = Math.max(period - 1, start);
    const segmentEnd = Math.min(period, end);
    if (segmentEnd > segmentStart) {
      const fraction = segmentEnd - segmentStart;
      const segmentAmount = Math.min(amount * fraction, cost - salvage - accumulated);
      total += segmentAmount;
      accumulated += segmentAmount;
    }
    amount = 0;
  }
  return total;
}

function accruedCouponInterest(
  issue: number,
  firstInterest: number,
  settlement: number,
  rateValue: number,
  par: number,
  frequency: number,
  basis: number,
  fromIssue: boolean,
): number {
  const coupon = (par * rateValue) / frequency;
  if (settlement <= firstInterest) {
    if (!fromIssue) return 0;
    return coupon * frequency * yearFraction(issue, settlement, basis);
  }
  const firstDate = toUtcDate(firstInterest);
  const settlementDate = toUtcDate(settlement);
  const maturityAnchor = addCouponMonths(firstDate, 12 / frequency);
  let previous = firstDate;
  let next = maturityAnchor;
  let fullCoupons = 0;
  while (next.getTime() <= settlementDate.getTime()) {
    previous = next;
    next = addCouponMonths(next, 12 / frequency);
    fullCoupons += 1;
  }
  const fraction =
    (fromIssue ? couponFraction(toUtcDate(issue), firstDate, basis) : 0) +
    fullCoupons +
    couponFraction(previous, settlementDate, basis);
  return coupon * fraction;
}

function pricemat(
  settlement: number,
  maturity: number,
  issue: number,
  rateValue: number,
  yieldValue: number,
  basis: number,
): number {
  const maturityDays = yearFraction(issue, maturity, basis);
  const settlementDays = yearFraction(issue, settlement, basis);
  const remainingDays = yearFraction(settlement, maturity, basis);
  const redemption = 100 + 100 * rateValue * maturityDays;
  return redemption / (1 + yieldValue * remainingDays) - 100 * rateValue * settlementDays;
}

function macaulayDuration(
  settlement: number,
  maturity: number,
  couponRate: number,
  yieldValue: number,
  frequency: number,
  basis: number,
): number {
  const market = couponPrice(settlement, maturity, couponRate, yieldValue, 100, frequency, basis);
  const { couponCount } = couponSchedule(settlement, maturity, frequency);
  const coupon = (100 * couponRate) / frequency;
  const periodicYield = yieldValue / frequency + 1;
  let weighted = 0;
  let presentValue = 0;
  for (let index = 0; index < couponCount; index += 1) {
    const time = index + market.fractionalPeriod;
    const discounted = coupon / periodicYield ** time;
    weighted += time * discounted;
    presentValue += discounted;
  }
  const maturityTime = couponCount - 1 + market.fractionalPeriod;
  const redemptionValue = 100 / periodicYield ** maturityTime;
  weighted += maturityTime * redemptionValue;
  presentValue += redemptionValue;
  return weighted / presentValue;
}

export const financialFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  [
    "FV",
    ([rateValue, periods, payment, present, type]) => {
      const growthRate = rate(rateValue!);
      const count = number(periods!);
      const paymentValue = number(payment!);
      const presentValue = present === undefined ? 0 : number(present);
      const typeValue = type === undefined ? 0 : integer(number(type));
      if (typeValue !== 0 && typeValue !== 1) throw new FormulaValueError("#NUM!");
      if (growthRate === 0) return -(presentValue + paymentValue * count);
      const growth = 1 + growthRate;
      return -(
        presentValue * growth ** count +
        paymentValue * ((growth ** count - 1) / growthRate) * (1 + growthRate * typeValue)
      );
    },
  ],
  [
    "PV",
    ([rateValue, periods, payment, future, type]) => {
      const growthRate = rate(rateValue!);
      const count = number(periods!);
      const paymentValue = number(payment!);
      const futureValue = future === undefined ? 0 : number(future);
      const typeValue = type === undefined ? 0 : integer(number(type));
      if (typeValue !== 0 && typeValue !== 1) throw new FormulaValueError("#NUM!");
      if (growthRate === 0) return -paymentValue * count - futureValue;
      const growth = 1 + growthRate;
      return (
        -paymentValue * ((1 - growth ** -count) / growthRate) * (1 + growthRate * typeValue) -
        futureValue * growth ** -count
      );
    },
  ],
  [
    "PMT",
    ([rateValue, periods, present, future, type]) => {
      const typeValue = type === undefined ? 0 : integer(number(type));
      if (typeValue !== 0 && typeValue !== 1) throw new FormulaValueError("#NUM!");
      return annuityPayment(
        rate(rateValue!),
        number(periods!),
        number(present!),
        future === undefined ? 0 : number(future),
        typeValue,
      );
    },
  ],
  [
    "IPMT",
    ([rateValue, period, periods, present, future, type]) => {
      const typeValue = type === undefined ? 0 : integer(number(type));
      if (typeValue !== 0 && typeValue !== 1) throw new FormulaValueError("#NUM!");
      return annuityInterest(
        rate(rateValue!),
        number(period!),
        number(periods!),
        number(present!),
        future === undefined ? 0 : number(future),
        typeValue,
      );
    },
  ],
  [
    "PPMT",
    ([rateValue, period, periods, present, future, type]) => {
      const growthRate = rate(rateValue!);
      const periodValue = number(period!);
      const count = number(periods!);
      const presentValue = number(present!);
      const futureValue = future === undefined ? 0 : number(future);
      const typeValue = type === undefined ? 0 : integer(number(type));
      if (typeValue !== 0 && typeValue !== 1) throw new FormulaValueError("#NUM!");
      return (
        annuityPayment(growthRate, count, presentValue, futureValue, typeValue) -
        annuityInterest(growthRate, periodValue, count, presentValue, futureValue, typeValue)
      );
    },
  ],
  [
    "NPER",
    ([rateValue, payment, present, future, type]) => {
      const growthRate = rate(rateValue!);
      const paymentValue = number(payment!);
      const presentValue = number(present!);
      const futureValue = future === undefined ? 0 : number(future);
      const typeValue = type === undefined ? 0 : integer(number(type));
      if (typeValue !== 0 && typeValue !== 1) throw new FormulaValueError("#NUM!");
      const discountedPayment = paymentValue * (1 + growthRate * typeValue);
      if (growthRate === 0) {
        if (paymentValue === 0) throw new FormulaValueError("#NUM!");
        return -(presentValue + futureValue) / paymentValue;
      }
      const numerator = discountedPayment / growthRate - futureValue;
      const denominator = presentValue + discountedPayment / growthRate;
      if (denominator === 0 || numerator <= 0 || denominator <= 0)
        throw new FormulaValueError("#NUM!");
      return Math.log(numerator / denominator) / Math.log(1 + growthRate);
    },
  ],
  [
    "RATE",
    ([periods, payment, present, future, type, guess]) => {
      const count = number(periods!);
      const paymentValue = number(payment!);
      const presentValue = number(present!);
      const futureValue = future === undefined ? 0 : number(future);
      const typeValue = type === undefined ? 0 : integer(number(type));
      if ((typeValue !== 0 && typeValue !== 1) || count <= 0) throw new FormulaValueError("#NUM!");
      return solveContinuous(
        (value) => {
          if (value <= -1) return Number.NaN;
          if (value === 0) return presentValue + paymentValue * count + futureValue;
          const growth = 1 + value;
          return (
            presentValue * growth ** count +
            paymentValue * ((growth ** count - 1) / value) * (1 + value * typeValue) +
            futureValue
          );
        },
        guess === undefined ? 0.1 : number(guess),
      );
    },
  ],
  [
    "NPV",
    ([rateValue, ...cashFlows]) => {
      const discountRate = rate(rateValue!);
      const flowValues = flatSeries(cashFlows, 0);
      return flowValues.reduce(
        (total, flow, index) => total + flow / (1 + discountRate) ** (index + 1),
        0,
      );
    },
  ],
  [
    "IRR",
    ([cashFlows, guess]) => {
      const flowValues = matrixNumbers(cashFlows!);
      if (
        flowValues.length < 2 ||
        !flowValues.some((flow) => flow > 0) ||
        !flowValues.some((flow) => flow < 0)
      )
        throw new FormulaValueError("#NUM!");
      return solveContinuous(
        (value) => {
          if (value <= -1) return Number.NaN;
          return flowValues.reduce((total, flow, index) => total + flow / (1 + value) ** index, 0);
        },
        guess === undefined ? 0.1 : number(guess),
      );
    },
  ],
  [
    "XIRR",
    ([cashFlows, dates, guess], context) => {
      const flowValues = matrixNumbers(cashFlows!);
      const dateValues = matrixNumbers(dates!).map((value) => {
        value = dateValue(value, context);
        const date = toUtcDate(value);
        return date.getTime() / millisecondsPerDay + 25569;
      });
      if (
        flowValues.length !== dateValues.length ||
        flowValues.length < 2 ||
        !flowValues.some((flow) => flow > 0) ||
        !flowValues.some((flow) => flow < 0)
      )
        throw new FormulaValueError("#NUM!");
      const baseDate = dateValues[0]!;
      const times = dateValues.map((value) => (value - baseDate) / 365);
      return solveContinuous(
        (value) => {
          if (value <= -1) return Number.NaN;
          return flowValues.reduce(
            (total, flow, index) => total + flow / (1 + value) ** times[index]!,
            0,
          );
        },
        guess === undefined ? 0.1 : number(guess),
      );
    },
  ],
  [
    "MIRR",
    ([cashFlows, financeRate, reinvestRate]) => {
      const flowValues = matrixNumbers(cashFlows!);
      if (
        flowValues.length < 2 ||
        !flowValues.some((flow) => flow > 0) ||
        !flowValues.some((flow) => flow < 0)
      )
        throw new FormulaValueError("#NUM!");
      const financeValue = number(financeRate!);
      const reinvestValue = number(reinvestRate!);
      const negativePresent = flowValues.reduce(
        (total, flow, index) => total + (flow < 0 ? flow / (1 + financeValue) ** index : 0),
        0,
      );
      const positiveFuture = flowValues.reduce(
        (total, flow, index) =>
          total + (flow > 0 ? flow * (1 + reinvestValue) ** (flowValues.length - 1 - index) : 0),
        0,
      );
      if (negativePresent === 0) throw new FormulaValueError("#NUM!");
      return (positiveFuture / -negativePresent) ** (1 / (flowValues.length - 1)) - 1;
    },
  ],
  ["SLN", ([cost, salvage, life]) => (number(cost!) - number(salvage!)) / number(life!)],
  [
    "SYD",
    ([cost, salvage, life, period]) => {
      const lifeValue = number(life!);
      const periodValue = number(period!);
      if (lifeValue <= 0 || periodValue < 1 || periodValue > lifeValue)
        throw new FormulaValueError("#NUM!");
      return (
        ((number(cost!) - number(salvage!)) * (lifeValue - periodValue + 1) * 2) /
        (lifeValue * (lifeValue + 1))
      );
    },
  ],
  [
    "DDB",
    ([cost, salvage, life, period, factor]) => {
      const costValue = number(cost!);
      const salvageValue = number(salvage!);
      const lifeValue = number(life!);
      const periodValue = number(period!);
      const factorValue = factor === undefined ? 2 : number(factor);
      if (
        costValue < salvageValue ||
        lifeValue <= 0 ||
        periodValue < 1 ||
        periodValue > lifeValue ||
        factorValue <= 0
      )
        throw new FormulaValueError("#NUM!");
      let accumulated = 0;
      let amount = 0;
      for (let index = 1; index <= Math.ceil(periodValue); index += 1) {
        amount = Math.min(
          (costValue - accumulated) * (factorValue / lifeValue),
          costValue - salvageValue - accumulated,
        );
        const fraction = Math.min(Math.max(periodValue - index + 1, 0), 1);
        accumulated += amount * fraction;
      }
      return amount;
    },
  ],
  [
    "DB",
    ([cost, salvage, life, period, month]) => {
      const costValue = number(cost!);
      const salvageValue = number(salvage!);
      const lifeValue = number(life!);
      const periodValue = number(period!);
      const monthValue = month === undefined ? 12 : integer(number(month));
      if (
        costValue <= 0 ||
        salvageValue < 0 ||
        lifeValue <= 0 ||
        periodValue < 1 ||
        periodValue > lifeValue + 1 ||
        monthValue < 1 ||
        monthValue > 12
      )
        throw new FormulaValueError("#NUM!");
      const rateValue =
        Math.round((1 - (salvageValue / costValue) ** (1 / lifeValue)) * 1000) / 1000;
      let previous = costValue * rateValue * (monthValue / 12);
      if (periodValue === 1) return previous;
      let accumulated = previous;
      for (let index = 2; index <= Math.ceil(periodValue); index += 1) {
        previous =
          index === lifeValue + 1
            ? previous * ((12 - monthValue) / 12) * rateValue
            : (costValue - accumulated) * rateValue;
        accumulated += previous;
      }
      return previous;
    },
  ],
  [
    "VDB",
    ([cost, salvage, life, start, end, factor, noSwitch]) => {
      const noSwitchValue =
        noSwitch === undefined
          ? false
          : scalarText(noSwitch).toUpperCase() === "TRUE" || scalarText(noSwitch) === "1";
      return depreciationAmount(
        number(cost!),
        number(salvage!),
        number(life!),
        number(start!),
        number(end!),
        factor === undefined ? 2 : number(factor),
        noSwitchValue,
      );
    },
  ],
  [
    "CUMIPMT",
    ([rateValue, periods, present, start, end, type]) => {
      const growthRate = rate(rateValue!);
      const count = number(periods!);
      const presentValue = number(present!);
      const startPeriod = integer(number(start!));
      const endPeriod = integer(number(end!));
      const typeValue = type === undefined ? 0 : integer(number(type));
      if (
        startPeriod < 1 ||
        endPeriod < startPeriod ||
        endPeriod > count ||
        (typeValue !== 0 && typeValue !== 1)
      )
        throw new FormulaValueError("#NUM!");
      let total = 0;
      for (let period = startPeriod; period <= endPeriod; period += 1) {
        total += annuityInterest(growthRate, period, count, presentValue, 0, typeValue);
      }
      return total;
    },
  ],
  [
    "CUMPRINC",
    ([rateValue, periods, present, start, end, type]) => {
      const growthRate = rate(rateValue!);
      const count = number(periods!);
      const presentValue = number(present!);
      const startPeriod = integer(number(start!));
      const endPeriod = integer(number(end!));
      const typeValue = type === undefined ? 0 : integer(number(type));
      if (
        startPeriod < 1 ||
        endPeriod < startPeriod ||
        endPeriod > count ||
        (typeValue !== 0 && typeValue !== 1)
      )
        throw new FormulaValueError("#NUM!");
      const payment = annuityPayment(growthRate, count, presentValue, 0, typeValue);
      let total = 0;
      for (let period = startPeriod; period <= endPeriod; period += 1) {
        total += payment - annuityInterest(growthRate, period, count, presentValue, 0, typeValue);
      }
      return total;
    },
  ],
  [
    "EFFECT",
    ([nominal, periods]) => {
      const nominalValue = number(nominal!);
      const frequency = integer(number(periods!), false);
      if (nominalValue < 0 || frequency < 1) throw new FormulaValueError("#NUM!");
      return (1 + nominalValue / frequency) ** frequency - 1;
    },
  ],
  [
    "NOMINAL",
    ([effective, periods]) => {
      const effectiveValue = number(effective!);
      const frequency = integer(number(periods!), false);
      if (effectiveValue < 0 || frequency < 1) throw new FormulaValueError("#NUM!");
      return frequency * ((1 + effectiveValue) ** (1 / frequency) - 1);
    },
  ],
  [
    "PRICE",
    ([settlement, maturity, couponRate, yieldValue, redemption, frequency, basis], context) => {
      const bond = couponValues(
        dateValue(settlement!, context),
        dateValue(maturity!, context),
        frequency!,
        basis,
      );
      return couponPrice(
        bond.settlement,
        bond.maturity,
        number(couponRate!),
        rate(yieldValue!),
        number(redemption!),
        bond.frequency,
        bond.basis,
      ).price;
    },
  ],
  [
    "YIELD",
    ([settlement, maturity, couponRate, price, redemption, frequency, basis], context) => {
      const bond = couponValues(
        dateValue(settlement!, context),
        dateValue(maturity!, context),
        frequency!,
        basis,
      );
      const couponRateValue = number(couponRate!);
      const redemptionValue = number(redemption!);
      const accrued = couponPrice(
        bond.settlement,
        bond.maturity,
        couponRateValue,
        0.1,
        redemptionValue,
        bond.frequency,
        bond.basis,
      ).accrued;
      return solveContinuous(
        (value) =>
          couponPrice(
            bond.settlement,
            bond.maturity,
            couponRateValue,
            value,
            redemptionValue,
            bond.frequency,
            bond.basis,
          ).price -
          (number(price!) + accrued),
        0.1,
      );
    },
  ],
  [
    "DISC",
    ([settlement, maturity, price, redemption, basis], context) => {
      requireBetween(integer(number(basis ?? 0)), 0, 4);
      const maturityValue = dateValue(maturity!, context);
      if (dateValue(settlement!, context) >= maturityValue) throw new FormulaValueError("#NUM!");
      return (
        (number(redemption!) - number(price!)) /
        number(redemption!) /
        yearFraction(dateValue(settlement!, context), maturityValue, integer(number(basis ?? 0)))
      );
    },
  ],
  [
    "ACCRINT",
    (
      [issue, firstInterest, settlement, couponRate, par, frequency, basis, calcMethod],
      context,
    ) => {
      const frequencyValue = integer(number(frequency!), false);
      if (frequencyValue !== 1 && frequencyValue !== 2 && frequencyValue !== 4)
        throw new FormulaValueError("#NUM!");
      const basisValue = integer(number(basis ?? 0));
      requireBetween(basisValue, 0, 4);
      const methodText = calcMethod === undefined ? "TRUE" : scalarText(calcMethod).toUpperCase();
      if (methodText !== "TRUE" && methodText !== "FALSE") throw new FormulaValueError("#VALUE!");
      return accruedCouponInterest(
        dateValue(issue!, context),
        dateValue(firstInterest!, context),
        dateValue(settlement!, context),
        number(couponRate!),
        number(par!),
        frequencyValue,
        basisValue,
        methodText !== "FALSE",
      );
    },
  ],
  [
    "ACCRINTM",
    ([issue, maturity, rateValue, par, basis], context) => {
      const basisValue = integer(number(basis ?? 0));
      requireBetween(basisValue, 0, 4);
      return (
        number(par!) *
        number(rateValue!) *
        yearFraction(dateValue(issue!, context), dateValue(maturity!, context), basisValue)
      );
    },
  ],
  [
    "RECEIVED",
    ([settlement, maturity, investment, discount, basis], context) => {
      const basisValue = integer(number(basis ?? 0));
      requireBetween(basisValue, 0, 4);
      const maturityValue = dateValue(maturity!, context);
      const investmentValue = number(investment!);
      const discountValue = number(discount!);
      if (
        dateValue(settlement!, context) >= maturityValue ||
        investmentValue <= 0 ||
        discountValue <= 0
      )
        throw new FormulaValueError("#NUM!");
      const denominator =
        1 -
        discountValue * yearFraction(dateValue(settlement!, context), maturityValue, basisValue);
      if (denominator <= 0) throw new FormulaValueError("#NUM!");
      return investmentValue / denominator;
    },
  ],
  [
    "INTRATE",
    ([settlement, maturity, investment, redemption, basis], context) => {
      const basisValue = integer(number(basis ?? 0));
      requireBetween(basisValue, 0, 4);
      const maturityValue = dateValue(maturity!, context);
      const investmentValue = number(investment!);
      if (dateValue(settlement!, context) >= maturityValue || investmentValue <= 0)
        throw new FormulaValueError("#NUM!");
      return (
        (number(redemption!) - investmentValue) /
        investmentValue /
        yearFraction(dateValue(settlement!, context), maturityValue, basisValue)
      );
    },
  ],
  [
    "PRICEMAT",
    ([settlement, maturity, issue, rateValue, yieldValue, basis], context) => {
      const basisValue = integer(number(basis ?? 0));
      requireBetween(basisValue, 0, 4);
      return pricemat(
        dateValue(settlement!, context),
        dateValue(maturity!, context),
        dateValue(issue!, context),
        number(rateValue!),
        rate(yieldValue!),
        basisValue,
      );
    },
  ],
  [
    "YIELDMAT",
    ([settlement, maturity, issue, rateValue, price, basis], context) => {
      const basisValue = integer(number(basis ?? 0));
      requireBetween(basisValue, 0, 4);
      const settlementValue = dateValue(settlement!, context);
      const maturityValue = dateValue(maturity!, context);
      const issueValue = dateValue(issue!, context);
      const fixedRate = number(rateValue!);
      return solveContinuous(
        (value) =>
          pricemat(settlementValue, maturityValue, issueValue, fixedRate, value, basisValue) -
          number(price!),
        0.1,
      );
    },
  ],
  [
    "DURATION",
    ([settlement, maturity, couponRate, yieldValue, frequency, basis], context) => {
      const bond = couponValues(
        dateValue(settlement!, context),
        dateValue(maturity!, context),
        frequency!,
        basis,
      );
      return macaulayDuration(
        bond.settlement,
        bond.maturity,
        number(couponRate!),
        rate(yieldValue!),
        bond.frequency,
        bond.basis,
      );
    },
  ],
  [
    "MDURATION",
    ([settlement, maturity, couponRate, yieldValue, frequency, basis], context) => {
      const bond = couponValues(
        dateValue(settlement!, context),
        dateValue(maturity!, context),
        frequency!,
        basis,
      );
      const yieldRate = rate(yieldValue!);
      return (
        macaulayDuration(
          bond.settlement,
          bond.maturity,
          number(couponRate!),
          yieldRate,
          bond.frequency,
          bond.basis,
        ) /
        (1 + yieldRate / bond.frequency)
      );
    },
  ],
]);
