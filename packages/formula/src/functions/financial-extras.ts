import type { FormulaEvaluationContext } from "../context";
import { FormulaValueError } from "../errors";
import type { FormulaValue } from "../value";
import { financialFunctions } from "./financial";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { scalarNumber } from "./utils";

type FinancialArgs = Parameters<FormulaFunctionImplementation>[0];
type FinancialArg = FinancialArgs[number];

function couponDaysNext(args: FinancialArgs, context: FormulaEvaluationContext): FormulaValue {
  const selectedBasis = basisValue(args[3]);
  requireBefore(dateValue(args[0]!, context), dateValue(args[1]!, context));
  const boundaries = couponBoundaries(
    dateValue(args[0]!, context),
    dateValue(args[1]!, context),
    frequencyValue(args[2]!),
  );
  return countDays(dateValue(args[0]!, context), serial(boundaries.next), selectedBasis);
}

const millisecondsPerDay = 86_400_000;
const serialEpochOffset = 25569;

function number(value: FinancialArg): number {
  const output = scalarNumber(value);
  if (!Number.isFinite(output)) throw new FormulaValueError("#NUM!");
  return output;
}

function aliasFunction(table: FormulaFunctionTable, name: string): FormulaFunctionImplementation {
  const implementation = table.get(name);
  if (!implementation) throw new Error(`unregistered financial alias target: ${name}`);
  return implementation;
}

function dateValue(value: FinancialArg, context: FormulaEvaluationContext): number {
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

function requirePositive(value: number, allowZero = false): number {
  if (!Number.isFinite(value) || (allowZero ? value < 0 : value <= 0))
    throw new FormulaValueError("#NUM!");
  return value;
}

function output(value: number): number {
  if (!Number.isFinite(value)) throw new FormulaValueError("#NUM!");
  return value;
}

function toUtcDate(serial: number): Date {
  const integerSerial = Math.trunc(serial);
  const date = new Date(Math.round((integerSerial - serialEpochOffset) * millisecondsPerDay));
  if (Number.isNaN(date.getTime()) || date.getUTCFullYear() < 1900 || date.getUTCFullYear() > 9999)
    throw new FormulaValueError("#NUM!");
  return date;
}

function serial(date: Date): number {
  const value = Math.round(date.getTime() / millisecondsPerDay + serialEpochOffset);
  toUtcDate(value);
  return value;
}

function serialValue(date: Date, context: FormulaEvaluationContext): number {
  const output =
    (context.dateSystem ?? context.workbook.dateSystem) === "1904"
      ? serial(date) - 1462
      : serial(date);
  return output;
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

function thirty360Days(start: number, end: number, european: boolean): number {
  const startDate = toUtcDate(start);
  const endDate = toUtcDate(end);
  const startDay = Math.min(startDate.getUTCDate(), 30);
  const endDay =
    !european && startDay === 30 && endDate.getUTCDate() === 31
      ? 30
      : Math.min(endDate.getUTCDate(), 30);
  return (
    (endDate.getUTCFullYear() - startDate.getUTCFullYear()) * 360 +
    (endDate.getUTCMonth() - startDate.getUTCMonth()) * 30 +
    endDay -
    startDay
  );
}

function countDays(start: number, end: number, basis: number): number {
  return basis === 0 || basis === 4
    ? thirty360Days(start, end, basis === 4)
    : dayDifference(start, end);
}

function yearFraction(start: number, end: number, basis: number): number {
  requireBetween(basis, 0, 4);
  const startDate = toUtcDate(start);
  const endDate = toUtcDate(end);
  if (endDate.getTime() < startDate.getTime()) throw new FormulaValueError("#NUM!");
  if (basis === 0 || basis === 4) return thirty360Days(start, end, basis === 4) / 360;
  const days = dayDifference(start, end);
  if (basis === 1) return days / yearDays(startDate.getUTCFullYear());
  return days / 365;
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

function frequencyValue(value: FinancialArg): number {
  const output = integer(number(value), false);
  if (output !== 1 && output !== 2 && output !== 4) throw new FormulaValueError("#NUM!");
  return output;
}

function basisValue(value: FinancialArg | undefined): number {
  return requireBetween(integer(number(value ?? 0)), 0, 4);
}

function requireBefore(start: number, end: number): void {
  if (start >= end) throw new FormulaValueError("#NUM!");
}

function couponBoundaries(
  settlement: number,
  anchor: number,
  frequency: number,
): { previous: Date; next: Date } {
  const settlementDate = toUtcDate(settlement);
  const anchorDate = toUtcDate(anchor);
  const months = 12 / frequency;
  let previous = anchorDate;
  while (previous.getTime() > settlementDate.getTime())
    previous = addCouponMonths(previous, -months);
  if (previous.getTime() === anchorDate.getTime()) throw new FormulaValueError("#NUM!");
  return { previous, next: addCouponMonths(previous, months) };
}

function forwardCouponBoundaries(
  settlement: number,
  anchor: number,
  frequency: number,
): { previous: Date; next: Date } {
  const settlementDate = toUtcDate(settlement);
  const months = 12 / frequency;
  let previous = toUtcDate(anchor);
  while (previous.getTime() <= settlementDate.getTime())
    previous = addCouponMonths(previous, months);
  return { previous: addCouponMonths(previous, -months), next: previous };
}

function couponCount(anchor: number, maturity: number, frequency: number): number {
  const maturityDate = toUtcDate(maturity);
  const months = 12 / frequency;
  let count = 0;
  for (
    let coupon = addCouponMonths(toUtcDate(anchor), months);
    coupon.getTime() <= maturityDate.getTime();
    coupon = addCouponMonths(coupon, months)
  ) {
    count += 1;
  }
  return count;
}

function periodFraction(
  start: Date,
  end: Date,
  periodStart: Date,
  periodEnd: Date,
  basis: number,
): number {
  const numerator = countDays(serial(start), serial(end), basis);
  const denominator = countDays(serial(periodStart), serial(periodEnd), basis);
  if (denominator <= 0) throw new FormulaValueError("#NUM!");
  return numerator / denominator;
}

function oddFirstPrice(
  settlement: number,
  maturity: number,
  issue: number,
  firstCoupon: number,
  couponRate: number,
  yieldValue: number,
  redemption: number,
  frequency: number,
  basis: number,
): number {
  requireBefore(settlement, maturity);
  requireBefore(settlement, firstCoupon);
  requireBefore(firstCoupon, maturity);
  requireBefore(issue, settlement);
  if (yieldValue < 0 || couponRate < 0 || redemption <= 0) throw new FormulaValueError("#NUM!");
  const boundaries = couponBoundaries(settlement, firstCoupon, frequency);
  const count = 1 + couponCount(firstCoupon, maturity, frequency);
  const firstDate = toUtcDate(firstCoupon);
  const issueDate = toUtcDate(issue);
  const settlementDate = toUtcDate(settlement);
  const months = 12 / frequency;
  let oddStart = firstDate;
  while (addCouponMonths(oddStart, -months).getTime() > issueDate.getTime())
    oddStart = addCouponMonths(oddStart, -months);
  const quasiPeriods: Array<{ start: Date; end: Date }> = [];
  for (let end = firstDate; ; end = addCouponMonths(end, -months)) {
    const start = addCouponMonths(end, -months);
    quasiPeriods.unshift({ start, end });
    if (start.getTime() <= issueDate.getTime()) break;
  }
  let oddCouponFactor = 0;
  let accruedFactor = 0;
  let completeQuasiCoupons = 0;
  for (const period of quasiPeriods) {
    const normalDays = countDays(serial(period.start), serial(period.end), basis);
    const couponEnd = period.end;
    const couponStart = period.start.getTime() < issueDate.getTime() ? issueDate : period.start;
    oddCouponFactor += countDays(serial(couponStart), serial(couponEnd), basis) / normalDays;
    const accrualStart = period.start.getTime() < issueDate.getTime() ? issueDate : period.start;
    const accrualEnd =
      period.end.getTime() > settlementDate.getTime() ? settlementDate : period.end;
    if (accrualEnd.getTime() > accrualStart.getTime())
      accruedFactor += countDays(serial(accrualStart), serial(accrualEnd), basis) / normalDays;
    if (period.end.getTime() <= settlementDate.getTime()) completeQuasiCoupons += 1;
  }
  const firstPeriodFraction = periodFraction(
    toUtcDate(settlement),
    boundaries.next,
    boundaries.previous,
    boundaries.next,
    basis,
  );
  const firstFraction = firstPeriodFraction + completeQuasiCoupons;
  const coupon = (redemption * couponRate) / frequency;
  const periodicGrowth = 1 + yieldValue / frequency;
  if (periodicGrowth <= 0) throw new FormulaValueError("#NUM!");
  let presentValue = 0;
  presentValue += (coupon * oddCouponFactor) / periodicGrowth ** firstFraction;
  for (let index = 1; index < count; index += 1)
    presentValue += coupon / periodicGrowth ** (index + firstFraction);
  presentValue += redemption / periodicGrowth ** (count - 1 + firstFraction);
  return presentValue - coupon * accruedFactor;
}

function oddLastPrice(
  settlement: number,
  maturity: number,
  lastInterest: number,
  couponRate: number,
  yieldValue: number,
  redemption: number,
  frequency: number,
  basis: number,
): number {
  requireBefore(settlement, maturity);
  requireBefore(lastInterest, settlement);
  if (yieldValue < 0 || couponRate < 0 || redemption <= 0) throw new FormulaValueError("#NUM!");
  const boundaries = forwardCouponBoundaries(settlement, lastInterest, frequency);
  const count = couponCount(lastInterest, maturity, frequency) + 1;
  if (count < 2) throw new FormulaValueError("#NUM!");
  const fractionalPeriod = periodFraction(
    toUtcDate(settlement),
    boundaries.next,
    toUtcDate(lastInterest),
    boundaries.next,
    basis,
  );
  const accruedFraction = periodFraction(
    toUtcDate(lastInterest),
    toUtcDate(settlement),
    boundaries.previous,
    boundaries.next,
    basis,
  );
  const coupon = (redemption * couponRate) / frequency;
  const periodicGrowth = 1 + yieldValue / frequency;
  if (periodicGrowth <= 0) throw new FormulaValueError("#NUM!");
  let presentValue = 0;
  for (let index = 0; index < count; index += 1)
    presentValue += coupon / periodicGrowth ** (index + fractionalPeriod);
  presentValue += redemption / periodicGrowth ** (count - 1 + fractionalPeriod);
  return presentValue - coupon * accruedFraction;
}

function solveForRate(evaluate: (rate: number) => number, guess: number): number {
  let current = Number.isFinite(guess) && guess > -0.999 ? guess : 0.1;
  for (let iteration = 0; iteration < 100; iteration += 1) {
    const currentPrice = evaluate(current);
    if (!Number.isFinite(currentPrice)) break;
    const delta = Math.max(Math.abs(current) * 1e-7, 1e-8);
    const derivative = (evaluate(current + delta) - currentPrice) / delta;
    if (!Number.isFinite(derivative) || derivative === 0) break;
    const next = current - currentPrice / derivative;
    if (next <= -0.999999) break;
    if (Math.abs(next - current) < 1e-10) return output(next);
    current = next;
  }

  for (let low = -0.999; low < 10; low += low < 0 ? 0.01 : 0.05) {
    const high = low + (low < 0 ? 0.01 : 0.05);
    const lowValue = evaluate(low);
    const highValue = evaluate(high);
    if (!Number.isFinite(lowValue) || !Number.isFinite(highValue)) continue;
    if (lowValue === 0) return low;
    if (highValue === 0) return high;
    if (lowValue > 0 === highValue > 0) continue;
    let lower = low;
    let upper = high;
    for (let iteration = 0; iteration < 100; iteration += 1) {
      const middle = (lower + upper) / 2;
      const middleValue = evaluate(middle);
      if (!Number.isFinite(middleValue)) break;
      if (middleValue === 0) return middle;
      if (evaluate(lower) > 0 === middleValue > 0) lower = middle;
      else upper = middle;
    }
    return output((lower + upper) / 2);
  }
  throw new FormulaValueError("#NUM!");
}

function dollarParts(value: number, fraction: number): number {
  requireBetween(integer(fraction), 1, Number.MAX_SAFE_INTEGER);
  const whole = Math.trunc(value);
  const remainder = Math.abs(value - whole);
  const sign = value < 0 ? -1 : 1;
  return sign * (Math.abs(whole) + (remainder * 100) / fraction);
}

function amortizationPeriodsPerYear(basis: number): number {
  return basis === 2 || basis === 3 ? 4 : 1;
}

function amordegrc(
  settlement: number,
  maturity: number,
  cost: number,
  salvage: number,
  period: number,
  depreciationRate: number,
  basis: number,
): number {
  requireBefore(settlement, maturity);
  requireBetween(basis, 0, 4);
  requirePositive(cost);
  requirePositive(salvage, true);
  if (salvage > cost || depreciationRate <= 0) throw new FormulaValueError("#NUM!");
  const integerPeriod = integer(period, false);
  const life = Math.round(1 / depreciationRate);
  if (life < 3) throw new FormulaValueError("#NUM!");
  const coefficient = life <= 4 ? 1.5 : life <= 6 ? 2 : 2.5;
  if (life * (cost - salvage) < cost) throw new FormulaValueError("#NUM!");
  const periodsPerYear = amortizationPeriodsPerYear(basis);
  const totalPeriods = Math.round(life * periodsPerYear);
  if (integerPeriod > totalPeriods) return 0;
  const annualAmount = cost * depreciationRate * coefficient;
  const firstFraction = yearFraction(settlement, maturity, basis);
  if (integerPeriod === 1) return output(annualAmount * firstFraction);
  if (integerPeriod === totalPeriods)
    return output(Math.min(annualAmount * (1 - firstFraction), cost - salvage));
  return output(Math.min(annualAmount / periodsPerYear, cost - salvage));
}

function amorlinc(
  settlement: number,
  maturity: number,
  cost: number,
  salvage: number,
  period: number,
  depreciationRate: number,
  basis: number,
): number {
  requireBefore(settlement, maturity);
  requireBetween(basis, 0, 4);
  requirePositive(cost);
  requirePositive(salvage, true);
  if (salvage > cost || depreciationRate <= 0) throw new FormulaValueError("#NUM!");
  const integerPeriod = integer(period, false);
  const firstFraction = yearFraction(settlement, maturity, basis);
  const totalPeriods = Math.ceil(1 / firstFraction);
  if (integerPeriod > totalPeriods) return 0;
  const annualAmount = cost * depreciationRate;
  let remaining = cost - salvage;
  for (let currentPeriod = 1; currentPeriod <= totalPeriods; currentPeriod += 1) {
    const amount =
      currentPeriod === 1
        ? annualAmount * firstFraction
        : currentPeriod === totalPeriods
          ? annualAmount * (1 - firstFraction)
          : annualAmount;
    if (currentPeriod === integerPeriod) return output(Math.max(Math.min(amount, remaining), 0));
    remaining -= amount;
  }
  throw new FormulaValueError("#NUM!");
}

function treasuryBillDays(settlement: number, maturity: number): number {
  requireBefore(settlement, maturity);
  const days = dayDifference(settlement, maturity);
  requireBetween(days, 1, 365);
  return days;
}

function numericSeries(value: FinancialArg): number[] {
  if (!Array.isArray(value)) return [number(value)];
  return value
    .flat()
    .filter((entry) => entry !== null)
    .map((entry) => number(entry));
}

function xnpv(rateValue: number, cashFlows: number[], dates: number[]): number {
  if (cashFlows.length !== dates.length || cashFlows.length === 0)
    throw new FormulaValueError("#NUM!");
  if (rateValue <= -1) throw new FormulaValueError("#NUM!");
  toUtcDate(dates[0]!);
  let total = 0;
  for (let index = 0; index < cashFlows.length; index += 1) {
    const elapsedDays = dayDifference(dates[0]!, dates[index]!);
    total += cashFlows[index]! / (1 + rateValue) ** (elapsedDays / 365);
  }
  return total;
}

export const financialExtraFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  ["EFFECT_ADD", aliasFunction(financialFunctions, "EFFECT")],
  [
    "AMORDEGRC",
    ([settlement, maturity, cost, salvage, period, depreciationRate, basis], context) =>
      amordegrc(
        dateValue(settlement!, context),
        dateValue(maturity!, context),
        number(cost!),
        number(salvage!),
        number(period!),
        number(depreciationRate!),
        basisValue(basis),
      ),
  ],
  [
    "AMORLINC",
    ([settlement, maturity, cost, salvage, period, depreciationRate, basis], context) =>
      amorlinc(
        dateValue(settlement!, context),
        dateValue(maturity!, context),
        number(cost!),
        number(salvage!),
        number(period!),
        number(depreciationRate!),
        basisValue(basis),
      ),
  ],
  [
    "COUPDAYBS",
    ([settlement, maturity, frequency, basis], context) => {
      basisValue(basis);
      requireBefore(dateValue(settlement!, context), dateValue(maturity!, context));
      const boundaries = couponBoundaries(
        dateValue(settlement!, context),
        dateValue(maturity!, context),
        frequencyValue(frequency!),
      );
      return countDays(
        serial(boundaries.previous),
        dateValue(settlement!, context),
        basisValue(basis),
      );
    },
  ],
  [
    "COUPDAYS",
    ([settlement, maturity, frequency, basis], context) => {
      const selectedBasis = basisValue(basis);
      requireBefore(dateValue(settlement!, context), dateValue(maturity!, context));
      const boundaries = couponBoundaries(
        dateValue(settlement!, context),
        dateValue(maturity!, context),
        frequencyValue(frequency!),
      );
      if (selectedBasis === 0 || selectedBasis === 4) return 360 / frequencyValue(frequency!);
      return countDays(serial(boundaries.previous), serial(boundaries.next), selectedBasis);
    },
  ],
  ["COUPDAYSNC", couponDaysNext],
  ["COUPDAYNC", couponDaysNext],
  [
    "COUPNCD",
    ([settlement, maturity, frequency, basis], context) => {
      basisValue(basis);
      requireBefore(dateValue(settlement!, context), dateValue(maturity!, context));
      return serialValue(
        couponBoundaries(
          dateValue(settlement!, context),
          dateValue(maturity!, context),
          frequencyValue(frequency!),
        ).next,
        context,
      );
    },
  ],
  [
    "COUPNUM",
    ([settlement, maturity, frequency, basis], context) => {
      basisValue(basis);
      requireBefore(dateValue(settlement!, context), dateValue(maturity!, context));
      const boundaries = couponBoundaries(
        dateValue(settlement!, context),
        dateValue(maturity!, context),
        frequencyValue(frequency!),
      );
      return (
        couponCount(
          serial(boundaries.next),
          dateValue(maturity!, context),
          frequencyValue(frequency!),
        ) + 1
      );
    },
  ],
  [
    "COUPPCD",
    ([settlement, maturity, frequency, basis], context) => {
      basisValue(basis);
      requireBefore(dateValue(settlement!, context), dateValue(maturity!, context));
      return serialValue(
        couponBoundaries(
          dateValue(settlement!, context),
          dateValue(maturity!, context),
          frequencyValue(frequency!),
        ).previous,
        context,
      );
    },
  ],
  [
    "DOLLARDE",
    ([fractionalDollar, fraction]) =>
      output(dollarParts(number(fractionalDollar!), integer(number(fraction!), false))),
  ],
  [
    "DOLLARFR",
    ([decimalDollar, fraction]) => {
      const denominator = integer(number(fraction!), false);
      requireBetween(denominator, 1, Number.MAX_SAFE_INTEGER);
      const value = number(decimalDollar!);
      const whole = Math.trunc(value);
      const sign = value < 0 ? -1 : 1;
      return output(sign * (Math.abs(whole) + (Math.abs(value - whole) * denominator) / 100));
    },
  ],
  [
    "ODDFPRICE",
    (
      [
        settlement,
        maturity,
        issue,
        firstCoupon,
        couponRate,
        yieldValue,
        redemption,
        frequency,
        basis,
      ],
      context,
    ) =>
      output(
        oddFirstPrice(
          dateValue(settlement!, context),
          dateValue(maturity!, context),
          dateValue(issue!, context),
          dateValue(firstCoupon!, context),
          number(couponRate!),
          number(yieldValue!),
          number(redemption!),
          frequencyValue(frequency!),
          basisValue(basis),
        ),
      ),
  ],
  [
    "ODDFYIELD",
    (
      [settlement, maturity, issue, firstCoupon, couponRate, price, redemption, frequency, basis],
      context,
    ) => {
      const settlementValue = dateValue(settlement!, context);
      const maturityValue = dateValue(maturity!, context);
      const issueValue = dateValue(issue!, context);
      const firstCouponValue = dateValue(firstCoupon!, context);
      const couponRateValue = number(couponRate!);
      const redemptionValue = number(redemption!);
      const priceValue = requirePositive(number(price!));
      const selectedFrequency = frequencyValue(frequency!);
      const selectedBasis = basisValue(basis);
      return solveForRate(
        (yieldValue) =>
          oddFirstPrice(
            settlementValue,
            maturityValue,
            issueValue,
            firstCouponValue,
            couponRateValue,
            yieldValue,
            redemptionValue,
            selectedFrequency,
            selectedBasis,
          ) - priceValue,
        0.1,
      );
    },
  ],
  [
    "ODDLPRICE",
    (
      [settlement, maturity, lastInterest, couponRate, yieldValue, redemption, frequency, basis],
      context,
    ) =>
      output(
        oddLastPrice(
          dateValue(settlement!, context),
          dateValue(maturity!, context),
          dateValue(lastInterest!, context),
          number(couponRate!),
          number(yieldValue!),
          number(redemption!),
          frequencyValue(frequency!),
          basisValue(basis),
        ),
      ),
  ],
  [
    "ODDLYIELD",
    (
      [settlement, maturity, lastInterest, couponRate, price, redemption, frequency, basis],
      context,
    ) => {
      const settlementValue = dateValue(settlement!, context);
      const maturityValue = dateValue(maturity!, context);
      const lastInterestValue = dateValue(lastInterest!, context);
      const couponRateValue = number(couponRate!);
      const redemptionValue = number(redemption!);
      const priceValue = requirePositive(number(price!));
      const selectedFrequency = frequencyValue(frequency!);
      const selectedBasis = basisValue(basis);
      return solveForRate(
        (yieldValue) =>
          oddLastPrice(
            settlementValue,
            maturityValue,
            lastInterestValue,
            couponRateValue,
            yieldValue,
            redemptionValue,
            selectedFrequency,
            selectedBasis,
          ) - priceValue,
        0.1,
      );
    },
  ],
  [
    "PRICEDISC",
    ([settlement, maturity, discount, redemption, basis], context) => {
      const selectedBasis = basisValue(basis);
      const settlementValue = dateValue(settlement!, context);
      const maturityValue = dateValue(maturity!, context);
      requireBefore(settlementValue, maturityValue);
      const discountValue = number(discount!);
      const redemptionValue = number(redemption!);
      if (discountValue <= 0 || redemptionValue <= 0) throw new FormulaValueError("#NUM!");
      return output(
        redemptionValue *
          (1 - discountValue * yearFraction(settlementValue, maturityValue, selectedBasis)),
      );
    },
  ],
  [
    "TBILLEQ",
    ([settlement, maturity, discount], context) => {
      const days = treasuryBillDays(dateValue(settlement!, context), dateValue(maturity!, context));
      const discountValue = requirePositive(number(discount!));
      return output((365 * discountValue) / (360 - discountValue * days));
    },
  ],
  [
    "TBILLPRICE",
    ([settlement, maturity, discount], context) => {
      const days = treasuryBillDays(dateValue(settlement!, context), dateValue(maturity!, context));
      const discountValue = requirePositive(number(discount!));
      return output(100 * (1 - (discountValue * days) / 360));
    },
  ],
  [
    "TBILLYIELD",
    ([settlement, maturity, price], context) => {
      const days = treasuryBillDays(dateValue(settlement!, context), dateValue(maturity!, context));
      const priceValue = requirePositive(number(price!));
      return output(((100 - priceValue) / priceValue) * (360 / days));
    },
  ],
  [
    "XNPV",
    ([rateValue, cashFlows, dates], context) =>
      output(
        xnpv(
          number(rateValue!),
          numericSeries(cashFlows!),
          numericSeries(dates!).map((value) => dateValue(value, context)),
        ),
      ),
  ],
  [
    "YIELDDISC",
    ([settlement, maturity, price, redemption, basis], context) => {
      const selectedBasis = basisValue(basis);
      const settlementValue = dateValue(settlement!, context);
      const maturityValue = dateValue(maturity!, context);
      requireBefore(settlementValue, maturityValue);
      const priceValue = requirePositive(number(price!));
      const redemptionValue = requirePositive(number(redemption!));
      const fraction = yearFraction(settlementValue, maturityValue, selectedBasis);
      if (fraction <= 0) throw new FormulaValueError("#NUM!");
      return output((redemptionValue / priceValue) ** (1 / fraction) - 1);
    },
  ],
]);
