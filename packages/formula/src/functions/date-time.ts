import type { FormulaEvaluationContext } from "../context";
import { FormulaValueError } from "../errors";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { scalarBoolean, scalarNumber, scalarText } from "./utils";

const millisecondsPerDay = 86_400_000;
const maximumDateSerials = {
  "1900": 2_958_465,
  "1904": 2_957_003,
} as const;

class Excel1900LeapDay extends Date {
  sourceLeapDay = true;

  override getUTCMonth(): number {
    return 1;
  }

  override getUTCDate(): number {
    return 29;
  }

  override getUTCDay(): number {
    return 3;
  }

  override setUTCDate(date: number): number {
    const timeOfDay = this.getTime() % millisecondsPerDay;
    return super.setTime(Date.UTC(1900, 2, date) + timeOfDay);
  }
}

function validDateSerial(serial: number, dateSystem: "1900" | "1904" = "1900"): number {
  if (!Number.isFinite(serial)) throw new FormulaValueError("#VALUE!");
  if (serial < 0 || serial > maximumDateSerials[dateSystem]) throw new FormulaValueError("#NUM!");
  return serial;
}

export function toUtcDate(serial: number, dateSystem: "1900" | "1904" = "1900"): Date {
  serial = validDateSerial(serial, dateSystem);
  if (dateSystem === "1904") return new Date((serial - 24_107) * millisecondsPerDay);
  const epoch = Date.UTC(1899, 11, 31);
  if (Math.trunc(serial) === 60) {
    const leapDay = new Excel1900LeapDay(Date.UTC(1900, 2, 1));
    leapDay.setTime(leapDay.getTime() + (serial - 60) * millisecondsPerDay);
    return leapDay;
  }
  const days = serial <= 59 ? serial : serial - 1;
  return new Date(epoch + days * millisecondsPerDay);
}

export function toSerialDate(date: Date, dateSystem: "1900" | "1904" = "1900"): number {
  if (dateSystem === "1904")
    return validDateSerial(date.getTime() / millisecondsPerDay + 24_107, dateSystem);
  if (date instanceof Excel1900LeapDay) {
    const elapsed = (date.getTime() - Date.UTC(1900, 2, 1)) / millisecondsPerDay;
    const serial =
      elapsed === 0 && date.sourceLeapDay ? 60 : elapsed < 0 ? 60 + elapsed : 61 + elapsed;
    return validDateSerial(serial, dateSystem);
  }
  const days = (date.getTime() - Date.UTC(1899, 11, 31)) / millisecondsPerDay;
  return validDateSerial(days <= 59 ? days : days + 1, dateSystem);
}

function civilDateFromSerial(
  serial: number,
  dateSystem: "1900" | "1904",
): { year: number; month: number; day: number; time: number } {
  serial = validDateSerial(serial, dateSystem);
  const whole = Math.trunc(serial);
  const time = serial - whole;
  if (dateSystem === "1900" && whole === 60) return { year: 1900, month: 2, day: 29, time };
  const date = toUtcDate(whole, dateSystem);
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    time,
  };
}

function daysInCivilMonth(year: number, month: number): number {
  if (year === 1900 && month === 2) return 29;
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

function serialFromCivilDate(
  year: number,
  month: number,
  day: number,
  time: number,
  dateSystem: "1900" | "1904",
): number {
  if (dateSystem === "1900" && year === 1900 && month === 2 && day === 29)
    return validDateSerial(60 + time, dateSystem);
  const date = new Date(Date.UTC(year, month - 1, day) + time * millisecondsPerDay);
  return toSerialDate(date, dateSystem);
}

function monthOffset(
  serial: number,
  months: number,
  dateSystem: "1900" | "1904",
): { year: number; month: number; maximumDay: number; day: number; time: number } {
  const { year, month, day, time } = civilDateFromSerial(serial, dateSystem);
  const offset = Math.trunc(months);
  const totalMonth = year * 12 + month - 1 + offset;
  const targetYear = Math.floor(totalMonth / 12);
  const targetMonth = totalMonth - targetYear * 12 + 1;
  const sourceMonth = year * 12 + month - 1;
  const sourcePseudoLeapDay = dateSystem === "1900" && month === 2 && day === 29;
  const dayOffset =
    sourcePseudoLeapDay && totalMonth < sourceMonth
      ? daysInCivilMonth(targetYear, targetMonth)
      : day;
  return {
    year: targetYear,
    month: targetMonth,
    maximumDay: daysInCivilMonth(targetYear, targetMonth),
    day: dayOffset,
    time,
  };
}

export function excelWeekday(serial: number, dateSystem: "1900" | "1904"): number {
  if (dateSystem === "1904") return toUtcDate(serial, dateSystem).getUTCDay();
  return (((validDateSerial(serial, dateSystem) - 1) % 7) + 7) % 7;
}

export function parseDateText(value: string): Date | undefined {
  const match = /^\s*(\d{4,})[-/](\d{1,2})[-/](\d{1,2})\s*$/u.exec(value);
  if (!match) return undefined;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (year === 1900 && month === 2 && day === 29) return new Excel1900LeapDay(Date.UTC(1900, 2, 1));
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
    ? date
    : undefined;
}

function dateParts(args: Parameters<FormulaFunctionImplementation>[0]): Date {
  const rawYear = scalarNumber(args[0]!);
  const year = rawYear + (rawYear >= 0 && rawYear <= 1899 ? 1900 : 0);
  if (year === 1900 && scalarNumber(args[1]!) === 2 && scalarNumber(args[2]!) === 29)
    return new Excel1900LeapDay(
      Date.UTC(
        1900,
        1,
        29,
        args[3] === undefined ? 0 : scalarNumber(args[3]),
        args[4] === undefined ? 0 : scalarNumber(args[4]),
        args[5] === undefined ? 0 : scalarNumber(args[5]),
      ),
    );
  return new Date(
    Date.UTC(
      year,
      scalarNumber(args[1]!) - 1,
      scalarNumber(args[2]!),
      args[3] === undefined ? 0 : scalarNumber(args[3]),
      args[4] === undefined ? 0 : scalarNumber(args[4]),
      args[5] === undefined ? 0 : scalarNumber(args[5]),
    ),
  );
}

function weekendMask(weekend: string): boolean[] {
  if (/^[01]{7}$/u.test(weekend)) return Array.from(weekend).map((entry) => entry === "1");
  const masks: Record<string, boolean[]> = {
    "1": [true, false, false, false, false, false, true],
    "2": [false, true, false, false, false, false, false],
    "3": [false, false, true, false, false, false, false],
    "4": [false, false, false, true, false, false, false],
    "5": [false, false, false, false, true, false, false],
    "6": [false, false, false, false, false, true, false],
    "7": [false, false, false, false, false, false, true],
    "11": [true, false, false, false, false, false, true],
    "12": [true, true, false, false, false, false, false],
    "13": [true, false, true, false, false, false, false],
    "14": [true, false, false, true, false, false, false],
    "15": [true, false, false, false, true, false, false],
    "16": [true, false, false, false, false, true, false],
    "17": [false, false, false, false, false, false, true],
  };
  const mask = masks[weekend];
  if (!mask) throw new FormulaValueError("#NUM!");
  return mask;
}

function parseHolidays(
  value: Parameters<FormulaFunctionImplementation>[0][number],
  dateSystem: "1900" | "1904",
): number[] {
  const toHolidaySerial = (entry: Parameters<FormulaFunctionImplementation>[0][number]): number => {
    return Math.trunc(toSerialDate(datePartsFromScalar(entry, dateSystem), dateSystem));
  };
  if (!Array.isArray(value)) {
    if (value === null) return [];
    return [toHolidaySerial(value)];
  }
  return value
    .flat()
    .filter((entry) => entry !== null)
    .map(toHolidaySerial);
}

function datePartsFromScalar(
  value: Parameters<FormulaFunctionImplementation>[0][number],
  dateSystem: "1900" | "1904",
): Date {
  if (typeof value === "number") return toUtcDate(value, dateSystem);
  const parsed = parseDateText(scalarText(value));
  if (!parsed) throw new FormulaValueError("#VALUE!");
  return parsed;
}

function workday(
  startSerial: number,
  days: number,
  weekend: string,
  holidays: number[],
  dateSystem: "1900" | "1904",
): number {
  startSerial = validDateSerial(startSerial, dateSystem);
  const mask = weekendMask(weekend);
  let current = Math.trunc(startSerial);
  let remaining = Math.abs(Math.trunc(days));
  if (!Number.isFinite(remaining)) throw new FormulaValueError("#VALUE!");
  const direction = Math.sign(days) || 1;
  if (mask.every(Boolean)) throw new FormulaValueError("#NUM!");
  const maximumSerial = maximumDateSerials[dateSystem];
  const firstSerial = current + direction;
  const boundarySerial = direction > 0 ? maximumSerial : 0;
  if (firstSerial < 0 || firstSerial > maximumSerial) throw new FormulaValueError("#NUM!");
  const calendarDays = Math.abs(boundarySerial - firstSerial) + 1;
  const workdaysPerWeek = mask.filter((day) => !day).length;
  const fullWeeks = Math.floor(calendarDays / 7);
  let availableWorkdays = fullWeeks * workdaysPerWeek;
  let weekday = excelWeekday(firstSerial, dateSystem);
  for (let day = 0; day < calendarDays % 7; day += 1) {
    if (!mask[weekday]) availableWorkdays += 1;
    weekday = (weekday + 1) % 7;
  }
  if (remaining > availableWorkdays) throw new FormulaValueError("#NUM!");

  const holidaySet = new Set(holidays);
  const sortedHolidays = [...holidaySet].sort((left, right) => left - right);
  const hasHoliday = (left: number, right: number): boolean => {
    let low = 0;
    let high = sortedHolidays.length - 1;
    while (low <= high) {
      const middle = (low + high) >> 1;
      const holiday = sortedHolidays[middle]!;
      if (holiday < left) low = middle + 1;
      else if (holiday > right) high = middle - 1;
      else return true;
    }
    return false;
  };
  while (remaining > 0) {
    const fullWeeks = Math.floor(remaining / workdaysPerWeek);
    if (fullWeeks > 0) {
      const horizon = current + direction * fullWeeks * 7;
      const left = Math.min(current + direction, horizon);
      const right = Math.max(current + direction, horizon);
      if (!hasHoliday(left, right)) {
        remaining -= fullWeeks * workdaysPerWeek;
        current = horizon;
        continue;
      }
    }
    current += direction;
    if (!mask[excelWeekday(current, dateSystem)] && !holidaySet.has(Math.trunc(current)))
      remaining -= 1;
  }
  return Math.trunc(current);
}

function networkDays(
  startSerial: number,
  endSerial: number,
  weekend: string,
  holidays: number[],
  dateSystem: "1900" | "1904",
): number {
  startSerial = validDateSerial(startSerial, dateSystem);
  endSerial = validDateSerial(endSerial, dateSystem);
  const mask = weekendMask(weekend);
  const holidaySet = new Set(holidays);
  let count = 0;
  const [begin, end] =
    startSerial <= endSerial
      ? [Math.trunc(startSerial), Math.trunc(endSerial)]
      : [Math.trunc(endSerial), Math.trunc(startSerial)];
  for (let serial = begin; serial <= end; serial += 1) {
    if (!mask[excelWeekday(serial, dateSystem)] && !holidaySet.has(serial)) count += 1;
  }
  return startSerial <= endSerial ? count : -count;
}

function isoWeekOfYear(serial: number, dateSystem: "1900" | "1904"): number {
  const date = toUtcDate(serial, dateSystem);
  const target = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  const dayNumber = (date.getUTCDay() + 6) % 7;
  target.setUTCDate(target.getUTCDate() - dayNumber + 3);
  const firstThursday = new Date(Date.UTC(target.getUTCFullYear(), 0, 4));
  const firstDayNumber = (firstThursday.getUTCDay() + 6) % 7;
  firstThursday.setUTCDate(firstThursday.getUTCDate() - firstDayNumber + 3);
  return 1 + Math.round((target.getTime() - firstThursday.getTime()) / millisecondsPerDay / 7);
}

function weekOfYear(
  serial: number,
  weekStart: number,
  dateSystem: "1900" | "1904" = "1900",
): number {
  const date = toUtcDate(serial, dateSystem);
  const yearStart = Date.UTC(date.getUTCFullYear(), 0, 1);
  const dayOffset = (date.getTime() - yearStart) / millisecondsPerDay;
  const startDay = excelWeekday(serial - dayOffset, dateSystem);
  return Math.floor((dayOffset + ((startDay - weekStart + 7) % 7)) / 7) + 1;
}

function excelWeekNumber(serial: number, returnType: number, dateSystem: "1900" | "1904"): number {
  if (returnType === 21 || returnType === 150) return isoWeekOfYear(serial, dateSystem);
  const type = returnType >= 100 && returnType <= 117 ? returnType - 99 : returnType;
  const starts: Record<number, number> = {
    1: 0,
    2: 1,
    11: 1,
    12: 2,
    13: 3,
    14: 4,
    15: 5,
    16: 6,
    17: 0,
  };
  const weekStart = starts[type];
  if (weekStart === undefined) throw new FormulaValueError("#NUM!");
  return weekOfYear(serial, weekStart, dateSystem);
}

function thirtyDay360(start: Date, end: Date, european: boolean): number {
  let day1 = start.getUTCDate();
  let day2 = end.getUTCDate();
  if (european) {
    day1 = Math.min(day1, 30);
    day2 = Math.min(day2, 30);
  } else {
    const lastDay1 = new Date(
      Date.UTC(start.getUTCFullYear(), start.getUTCMonth() + 1, 0),
    ).getUTCDate();
    const lastDay2 = new Date(
      Date.UTC(end.getUTCFullYear(), end.getUTCMonth() + 1, 0),
    ).getUTCDate();
    if (day1 === lastDay1 && day1 === 31) day1 = 30;
    if (day2 === lastDay2 && day2 === 31 && day1 >= 30) day2 = 30;
  }
  return (
    (end.getUTCFullYear() - start.getUTCFullYear()) * 360 +
    (end.getUTCMonth() - start.getUTCMonth()) * 30 +
    day2 -
    day1
  );
}

function yearFraction(
  start: number,
  end: number,
  basis: number,
  dateSystem: "1900" | "1904",
): number | "#NUM!" {
  const startDate = toUtcDate(start, dateSystem);
  const endDate = toUtcDate(end, dateSystem);
  if (start > end) throw new FormulaValueError("#NUM!");
  if (basis === 0) return thirtyDay360(startDate, endDate, false) / 360;
  if (basis === 4) return thirtyDay360(startDate, endDate, true) / 360;
  const days = end - start;
  if (basis === 2) return days / 360;
  if (basis === 3) return days / 365;
  if (basis === 1) {
    const firstYear = startDate.getUTCFullYear();
    const lastYear = endDate.getUTCFullYear();
    if (firstYear === lastYear) return days / (isLeapYear(firstYear) ? 366 : 365);
    let denominator = 0;
    for (let year = firstYear; year <= lastYear; year += 1)
      denominator += isLeapYear(year) ? 366 : 365;
    return days / (denominator / (lastYear - firstYear + 1));
  }
  throw new FormulaValueError("#NUM!");
}

function isLeapYear(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

function timeSerial(args: Parameters<FormulaFunctionImplementation>[0]): number {
  const parts = args.slice(0, 3).map(scalarNumber);
  if (parts.some((part) => !Number.isFinite(part) || part < 0 || part > 32_767))
    throw new FormulaValueError("#NUM!");
  const [hours, minutes, seconds] = parts.map(Math.trunc);
  const totalSeconds = hours! * 3600 + minutes! * 60 + seconds!;
  return (totalSeconds % 86_400) / 86_400;
}

function timeValue(value: Parameters<FormulaFunctionImplementation>[0][number]): number | string {
  const text = scalarText(value);
  const match = /^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*(AM|PM))?$/iu.exec(text);
  if (!match) return Number.isFinite(Number(text)) ? Number(text) : "#VALUE!";
  let hours = Number(match[1]);
  if (match[4]?.toUpperCase() === "PM" && hours < 12) hours += 12;
  if (match[4]?.toUpperCase() === "AM" && hours === 12) hours = 0;
  return (hours * 3600 + Number(match[2]) * 60 + Number(match[3] ?? 0)) / 86_400;
}

function currentDate(context: Parameters<FormulaFunctionImplementation>[1]): Date {
  const date = context.workbook.now?.() ?? new Date();
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
}

function dateSystem(context: FormulaEvaluationContext): "1900" | "1904" {
  return context.dateSystem ?? context.workbook.dateSystem ?? "1900";
}

function isMonthEnd(date: Date): boolean {
  return (
    date.getUTCDate() ===
    new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0)).getUTCDate()
  );
}

function addMonthsClamped(date: Date, months: number): Date {
  const result = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + months, 1));
  result.setUTCDate(
    Math.min(
      date.getUTCDate(),
      new Date(Date.UTC(result.getUTCFullYear(), result.getUTCMonth() + 1, 0)).getUTCDate(),
    ),
  );
  return result;
}

function fullMonths(start: Date, end: Date): number {
  const months =
    (end.getUTCFullYear() - start.getUTCFullYear()) * 12 + end.getUTCMonth() - start.getUTCMonth();
  return end.getUTCDate() < start.getUTCDate() && !isMonthEnd(end) ? months - 1 : months;
}

function addYearsClamped(date: Date, years: number): Date {
  return addMonthsClamped(date, years * 12);
}

function datedDifference(
  startDate: Date,
  endDate: Date,
  startSerial: number,
  endSerial: number,
  unit: string,
  dateSystem: "1900" | "1904",
): number | "#NUM!" {
  if (endSerial < startSerial) return "#NUM!";
  if (unit === "D") return Math.trunc(endSerial - startSerial);
  const months = fullMonths(startDate, endDate);
  if (unit === "M") return months;
  if (unit === "Y") return Math.floor(months / 12);
  if (unit === "YM") return months % 12;
  if (unit === "MD") {
    if (endDate.getUTCDate() >= startDate.getUTCDate())
      return endDate.getUTCDate() - startDate.getUTCDate();
    const previousMonthEnd = new Date(Date.UTC(endDate.getUTCFullYear(), endDate.getUTCMonth(), 0));
    return endDate.getUTCDate() + previousMonthEnd.getUTCDate() - startDate.getUTCDate();
  }
  if (unit === "YD") {
    const years = Math.floor(months / 12);
    const anniversary = addYearsClamped(startDate, years);
    const effective =
      anniversary.getTime() > endDate.getTime()
        ? addYearsClamped(startDate, years - 1)
        : anniversary;
    return Math.trunc(endSerial - toSerialDate(effective, dateSystem));
  }
  return "#NUM!";
}

export const dateTimeFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  ["DATE", (args, context) => toSerialDate(dateParts(args), dateSystem(context))],
  [
    "DATESTRING",
    ([value], context) => {
      const date = datePartsFromScalar(value, dateSystem(context));
      if (Number.isNaN(date.getTime())) throw new FormulaValueError("#VALUE!");
      const year = String(date.getUTCFullYear()).padStart(4, "0");
      const month = String(date.getUTCMonth() + 1).padStart(2, "0");
      const day = String(date.getUTCDate()).padStart(2, "0");
      return `${year}-${month}-${day}`;
    },
  ],
  [
    "DATEDIF",
    ([start, end, unit], context) => {
      const system = dateSystem(context);
      const startDate = toUtcDate(scalarNumber(start), system);
      const endDate = toUtcDate(scalarNumber(end), system);
      const type = scalarText(unit).toUpperCase();
      return datedDifference(
        startDate,
        endDate,
        scalarNumber(start),
        scalarNumber(end),
        type,
        system,
      );
    },
  ],
  [
    "DATEVALUE",
    ([value], context) => {
      const parsed = parseDateText(scalarText(value));
      return parsed ? toSerialDate(parsed, dateSystem(context)) : "#VALUE!";
    },
  ],
  ["DAY", ([value], context) => toUtcDate(scalarNumber(value), dateSystem(context)).getUTCDate()],
  [
    "EDATE",
    ([value, months], context) => {
      const system = dateSystem(context);
      const target = monthOffset(scalarNumber(value), scalarNumber(months), system);
      return serialFromCivilDate(
        target.year,
        target.month,
        Math.min(target.day, target.maximumDay),
        target.time,
        system,
      );
    },
  ],
  [
    "EOMONTH",
    ([value, months], context) => {
      const system = dateSystem(context);
      const target = monthOffset(scalarNumber(value), scalarNumber(months), system);
      return serialFromCivilDate(target.year, target.month, target.maximumDay, target.time, system);
    },
  ],
  ["HOUR", ([value], context) => toUtcDate(scalarNumber(value), dateSystem(context)).getUTCHours()],
  [
    "MINUTE",
    ([value], context) => toUtcDate(scalarNumber(value), dateSystem(context)).getUTCMinutes(),
  ],
  [
    "MONTH",
    ([value], context) => toUtcDate(scalarNumber(value), dateSystem(context)).getUTCMonth() + 1,
  ],
  [
    "NOW",
    (_args, context) => toSerialDate(context.workbook.now?.() ?? new Date(), dateSystem(context)),
  ],
  [
    "NETWORKDAYS",
    ([start, end, holidays], context) =>
      networkDays(
        scalarNumber(start),
        scalarNumber(end),
        "1",
        holidays === undefined ? [] : parseHolidays(holidays, dateSystem(context)),
        dateSystem(context),
      ),
  ],
  [
    "NETWORKDAYS.INTL",
    ([start, end, weekend, holidays], context) =>
      networkDays(
        scalarNumber(start),
        scalarNumber(end),
        weekend === undefined ? "1" : scalarText(weekend),
        holidays === undefined ? [] : parseHolidays(holidays, dateSystem(context)),
        dateSystem(context),
      ),
  ],
  [
    "NETWORKDAYSINTL",
    ([start, end, weekend, holidays], context) =>
      networkDays(
        scalarNumber(start),
        scalarNumber(end),
        weekend === undefined ? "1" : scalarText(weekend),
        holidays === undefined ? [] : parseHolidays(holidays, dateSystem(context)),
        dateSystem(context),
      ),
  ],
  [
    "SECOND",
    ([value], context) => toUtcDate(scalarNumber(value), dateSystem(context)).getUTCSeconds(),
  ],
  ["TIME", timeSerial],
  ["TIMEVALUE", ([value]) => timeValue(value)],
  ["TODAY", (_args, context) => toSerialDate(currentDate(context), dateSystem(context))],
  [
    "WEEKDAY",
    (args, context) => {
      if (args.length > 3) return "#VALUE!";
      const value = args[0]!;
      const type = args[1];
      const day = excelWeekday(scalarNumber(value), dateSystem(context)) + 1;
      const returnType = type === undefined ? 1 : scalarNumber(type);
      if (returnType === 1) return day;
      if (returnType === 2) return day === 1 ? 7 : day - 1;
      if (returnType === 3) return day === 1 ? 6 : day - 2;
      if (returnType === 11) return day === 1 ? 7 : day - 1;
      if (returnType >= 11 && returnType <= 17) {
        const startDay = returnType === 17 ? 1 : returnType - 9;
        return ((((day - startDay + 7) % 7) + 7) % 7) + 1;
      }
      throw new FormulaValueError("#NUM!");
    },
  ],
  [
    "WEEKNUM",
    ([value, returnType], context) =>
      excelWeekNumber(
        scalarNumber(value),
        returnType === undefined ? 1 : scalarNumber(returnType),
        dateSystem(context),
      ),
  ],
  [
    "WEEKNUMOOO",
    ([value, returnType], context) =>
      excelWeekNumber(
        scalarNumber(value),
        returnType === undefined ? 1 : scalarNumber(returnType),
        dateSystem(context),
      ),
  ],
  ["ISOWEEKNUM", ([value], context) => isoWeekOfYear(scalarNumber(value), dateSystem(context))],
  [
    "WORKDAY",
    ([start, days, holidays], context) =>
      workday(
        scalarNumber(start),
        scalarNumber(days),
        "1",
        holidays === undefined ? [] : parseHolidays(holidays, dateSystem(context)),
        dateSystem(context),
      ),
  ],
  [
    "WORKDAY.INTL",
    ([start, days, weekend, holidays], context) =>
      workday(
        scalarNumber(start),
        scalarNumber(days),
        weekend === undefined ? "1" : scalarText(weekend),
        holidays === undefined ? [] : parseHolidays(holidays, dateSystem(context)),
        dateSystem(context),
      ),
  ],
  [
    "WORKDAYINTL",
    ([start, days, weekend, holidays], context) =>
      workday(
        scalarNumber(start),
        scalarNumber(days),
        weekend === undefined ? "1" : scalarText(weekend),
        holidays === undefined ? [] : parseHolidays(holidays, dateSystem(context)),
        dateSystem(context),
      ),
  ],
  [
    "YEAR",
    ([value], context) => toUtcDate(scalarNumber(value), dateSystem(context)).getUTCFullYear(),
  ],
  [
    "YEARFRAC",
    ([start, end, basis], context) =>
      yearFraction(
        scalarNumber(start),
        scalarNumber(end),
        basis === undefined ? 0 : scalarNumber(basis),
        dateSystem(context),
      ),
  ],
  [
    "DAYS",
    ([end, start], context) => {
      const system = dateSystem(context);
      return (
        validDateSerial(scalarNumber(end), system) - validDateSerial(scalarNumber(start), system)
      );
    },
  ],
  [
    "DAYS360",
    ([start, end, method], context) => {
      const system = dateSystem(context);
      const startDate = toUtcDate(scalarNumber(start), system);
      const endDate = toUtcDate(scalarNumber(end), system);
      const european = method !== undefined && scalarBoolean(method);
      let day1 = startDate.getUTCDate();
      let day2 = endDate.getUTCDate();
      if (european) {
        day1 = Math.min(day1, 30);
        day2 = Math.min(day2, 30);
      } else {
        if (day1 === 31) day1 = 30;
        if (day2 === 31 && day1 === 30) day2 = 30;
      }
      return (
        (endDate.getUTCFullYear() - startDate.getUTCFullYear()) * 360 +
        (endDate.getUTCMonth() - startDate.getUTCMonth()) * 30 +
        day2 -
        day1
      );
    },
  ],
]);
