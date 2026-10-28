import { FormulaError, type FormulaScalar, type FormulaValue, isFormulaErrorValue } from "../value";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { scalarBoolean, scalarNumber, scalarText, values } from "./utils";

function firstScalar(value: FormulaValue): FormulaScalar {
  return Array.isArray(value) ? (value.flat()[0] ?? null) : value;
}

function optionalNumber(value: FormulaValue | undefined): number | undefined {
  return value === undefined ? undefined : scalarNumber(firstScalar(value));
}

function isFormulaError(value: FormulaValue): boolean {
  const scalar = firstScalar(value);
  return isFormulaErrorValue(scalar);
}

function sheetCount(context: { getSheetName(id: number): string | undefined }): number {
  let count = 0;
  while (context.getSheetName(count) !== undefined) count += 1;
  return count;
}

function formatDollar(value: number, decimals: number, locale: string): string {
  const amount = Number(value.toFixed(decimals));
  const text = amount.toLocaleString(locale, {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
  return amount < 0 ? `($${text.replace("-", "")})` : `$${text}`;
}

function typedValue(value: FormulaValue): FormulaValue {
  const scalar = firstScalar(value);
  if (typeof scalar === "number") return scalar;
  if (typeof scalar === "boolean") return scalar ? 1 : 0;
  if (scalar === null) return 0;
  if (typeof scalar === "object") return "#VALUE!";
  const parsed = Number(scalar.trim());
  if (scalar.trim() !== "" && Number.isFinite(parsed)) return parsed;
  return 0;
}

function typeCode(value: FormulaValue): number {
  if (isFormulaError(value)) return 16;
  if (Array.isArray(value)) return 64;
  const scalar = firstScalar(value);
  if (typeof scalar === "boolean") return 4;
  if (typeof scalar === "number") return 1;
  return 2;
}

export const coreExtraFunctions: FormulaFunctionTable = new Map<
  string,
  FormulaFunctionImplementation
>([
  [
    "BASE",
    ([value, radix, minLength]) => {
      const number = Math.trunc(scalarNumber(value));
      const radixValue = Math.trunc(scalarNumber(radix!));
      if (radixValue < 2 || radixValue > 36) return "#NUM!";
      return Math.abs(number)
        .toString(radixValue)
        .toUpperCase()
        .padStart(optionalNumber(minLength) ?? 1, "0");
    },
  ],
  [
    "DOLLAR",
    ([value, decimals], context) =>
      formatDollar(
        scalarNumber(value),
        optionalNumber(decimals) ?? 2,
        context.workbook.locale ?? "en-US",
      ),
  ],
  [
    "USDOLLAR",
    ([value, decimals]) =>
      formatDollar(scalarNumber(value), optionalNumber(decimals) ?? 2, "en-US"),
  ],
  ["EASTERSUNDAY", ([year]) => easterSunday(Math.trunc(scalarNumber(year)))],
  ["ENCODEURL", ([text]) => encodeURIComponent(scalarText(text))],
  ["ERROR.TYPE", ([value]) => errorValueNumber(firstFormulaError(value))],
  [
    "EUROCONVERT",
    ([value, from, to], context) => {
      if (!context.workbook.getEuroExchangeRate) return "#VALUE!";
      const fromRate = context.workbook.getEuroExchangeRate(scalarText(from).toUpperCase());
      const toRate = context.workbook.getEuroExchangeRate(scalarText(to).toUpperCase());
      if (fromRate === undefined || toRate === undefined || fromRate === 0) return "#VALUE!";
      return (scalarNumber(value) / fromRate) * toRate;
    },
  ],
  [
    "ISERR",
    ([value]) => {
      const error = firstFormulaError(value);
      return error !== undefined && error.value !== "#N/A";
    },
  ],
  ["ISBLANK", ([value]) => firstScalar(value) === null],
  ["ISERROR", ([value]) => isFormulaError(value)],
  ["ISEVEN", ([value]) => Math.trunc(scalarNumber(value)) % 2 === 0],
  ["ISEVEN_ADD", ([value]) => Math.trunc(scalarNumber(value)) % 2 === 0],
  ["ISLOGICAL", ([value]) => typeof firstScalar(value) === "boolean"],
  ["ISNONTEXT", ([value]) => typeof firstScalar(value) !== "string"],
  ["ISNA", ([value]) => firstFormulaError(value)?.value === "#N/A"],
  ["ISODD", ([value]) => Math.trunc(scalarNumber(value)) % 2 !== 0],
  ["ISODD_ADD", ([value]) => Math.trunc(scalarNumber(value)) % 2 !== 0],
  ["ISNUMBER", ([value]) => typeof firstScalar(value) === "number"],
  ["ISTEXT", ([value]) => typeof firstScalar(value) === "string"],
  [
    "ISPMT",
    ([rate, period, periods, presentValue]) => {
      const totalPeriods = scalarNumber(periods);
      const currentPeriod = scalarNumber(period);
      const presentValueNumber = scalarNumber(presentValue!);
      if (totalPeriods === 0) return "#DIV/0!";
      return presentValueNumber * scalarNumber(rate) * (currentPeriod / totalPeriods - 1);
    },
  ],
  ["N", ([value]) => typedValue(value)],
  ["NA", () => "#N/A"],
  ["NOT", ([value]) => !scalarBoolean(value)],
  [
    "PDURATION",
    ([rate, presentValue, futureValue]) => {
      const rateValue = scalarNumber(rate);
      if (rateValue <= 0 || scalarNumber(presentValue) <= 0 || scalarNumber(futureValue) <= 0)
        return "#NUM!";
      return (
        (Math.log(scalarNumber(futureValue)) - Math.log(scalarNumber(presentValue))) /
        Math.log(1 + rateValue)
      );
    },
  ],
  ["PHONETIC", (args) => args.flatMap(values).map(scalarText).join("")],
  [
    "QUERYSTRING",
    (args) =>
      args
        .filter((_, index) => index % 2 === 0)
        .map(
          (key, index) =>
            `${encodeURIComponent(scalarText(key))}=${encodeURIComponent(scalarText(args[index * 2 + 1] ?? null))}`,
        )
        .join("&"),
  ],
  ["RAWSUBTRACT", ([left, right]) => scalarNumber(left) - scalarNumber(right)],
  ["ROUNDSIG", ([value, digits]) => roundSignificant(scalarNumber(value), scalarNumber(digits!))],
  [
    "RRI",
    ([periods, presentValue, futureValue]) => {
      const periodsValue = scalarNumber(periods);
      const presentValueNumber = scalarNumber(presentValue!);
      const futureValueNumber = scalarNumber(futureValue!);
      if (periodsValue <= 0 || presentValueNumber <= 0 || futureValueNumber < 0) return "#NUM!";
      return (futureValueNumber / presentValueNumber) ** (1 / periodsValue) - 1;
    },
  ],
  [
    "SHEET",
    ([name], context) => {
      if (name === undefined) return (context.sheet ?? 0) + 1;
      const text = scalarText(name);
      for (let id = 0; id < sheetCount(context.workbook); id += 1)
        if (context.workbook.getSheetName(id)?.toUpperCase() === text.toUpperCase()) return id + 1;
      return "#VALUE!";
    },
  ],
  ["SHEETS", (_args, context) => sheetCount(context.workbook)],
  [
    "SINGLE",
    ([value]) => {
      const matrix = Array.isArray(value) ? value : [[value]];
      if (matrix.length === 1 && matrix[0]?.length === 1) return matrix[0]?.[0] ?? null;
      return "#VALUE!";
    },
  ],
  ["TYPE", ([value]) => typeCode(value)],
  ["TRUE", () => true],
  [
    "XOR",
    (args) =>
      args
        .flatMap(values)
        .filter((value) => value !== null && value !== "")
        .reduce((result, value) => result !== scalarBoolean(value), false),
  ],
  ["ZGZ", ([text]) => scalarText(text).replace(/\s+/gu, " ").trim()],
]);

function easterSunday(year: number): number {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return Date.UTC(year, month - 1, day) / 86_400_000 + 25569;
}

function firstFormulaError(value: FormulaValue): FormulaError | undefined {
  const scalar = firstScalar(value);
  return isFormulaErrorValue(scalar) ? scalar : undefined;
}

function errorValueNumber(value: FormulaError | undefined): number | string {
  switch (value?.value) {
    case "#NULL!":
      return 1;
    case "#DIV/0!":
      return 2;
    case "#VALUE!":
      return 3;
    case "#REF!":
      return 4;
    case "#NAME?":
      return 5;
    case "#NUM!":
      return 6;
    case "#N/A":
      return 7;
    case "#GETTING_DATA":
      return 8;
    default:
      return "#N/A";
  }
}

function roundSignificant(value: number, digits: number): number {
  if (value === 0 || !Number.isFinite(value)) return value;
  const exponent = Math.trunc(digits) - 1 - Math.floor(Math.log10(Math.abs(value)));
  const factor = 10 ** exponent;
  return Math.round(value * factor) / factor;
}
