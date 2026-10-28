import type { FormulaTextEncoding } from "../context";
import { FormulaValueError } from "../errors";
import { formulaError, isMatrix, type FormulaValue } from "../value";
import { excelWeekday, parseDateText, toSerialDate, toUtcDate } from "./date-time";
import type { FormulaFunctionImplementation } from "./types";
import type { FormulaFunctionTable } from "./utils";
import { assertFormulaCellCount, scalarBoolean, scalarNumber, scalarText, values } from "./utils";

function wildcardPattern(value: string): string {
  const characters = Array.from(value);
  let output = "";
  for (let index = 0; index < characters.length; index += 1) {
    const character = characters[index]!;
    if (character === "~") {
      const escaped = characters[index + 1];
      if (escaped === "*" || escaped === "?" || escaped === "~") {
        output += escapeRegex(escaped);
        index += 1;
      } else output += escapeRegex(character);
    } else if (character === "*") output += "[\\s\\S]*";
    else if (character === "?") output += "[\\s\\S]";
    else output += escapeRegex(character);
  }
  return output;
}

function wildcardIndex(needle: string, text: string, start: number): number {
  const characters = Array.from(text);
  if (start < 1 || start > characters.length + 1) return -1;
  const match = new RegExp(wildcardPattern(needle), "iu").exec(
    characters.slice(start - 1).join(""),
  );
  if (!match) return -1;
  const preceding = Array.from(text.slice(start - 1, start - 1 + match.index)).length;
  return start + preceding;
}

function isDbcsCodePage(codePage: number | undefined): boolean {
  return codePage === 932 || codePage === 936 || codePage === 949 || codePage === 950;
}

function byteCharacterSize(character: string, encoding: FormulaTextEncoding | undefined): number {
  if (encoding?.encode) return encoding.encode(character).length;
  if (isDbcsCodePage(encoding?.codePage)) {
    const code = character.codePointAt(0)!;
    return code > 127 ? 2 : 1;
  }
  return character.codePointAt(0)! > 255 ? 2 : 1;
}

function codePointIndexOf(text: string, value: string, start: number): number {
  const source = Array.from(text);
  const search = Array.from(value);
  if (search.length === 0) return start >= 1 && start <= source.length + 1 ? start - 1 : -1;
  for (let index = start - 1; index <= source.length - search.length; index += 1)
    if (search.every((character, offset) => source[index + offset] === character)) return index;
  return -1;
}

function byteOffsetAtCharacterIndex(
  text: string,
  characterIndex: number,
  encoding: FormulaTextEncoding | undefined,
): number {
  const characters = Array.from(text);
  let byteIndex = 0;
  for (let index = 0; index < characterIndex; index += 1)
    byteIndex += byteCharacterSize(characters[index]!, encoding);
  return byteIndex;
}

function characterIndexAtByteBoundary(
  text: string,
  byteStart: number,
  encoding: FormulaTextEncoding | undefined,
): number | undefined {
  const characters = Array.from(text);
  if (byteStart < 1 || byteStart > byteLength(text, encoding)) return undefined;
  if (byteStart === 1) return 0;
  let consumedBytes = 0;
  for (let index = 0; index < characters.length; index += 1) {
    consumedBytes += byteCharacterSize(characters[index]!, encoding);
    if (byteStart === consumedBytes + 1) return index + 1;
  }
  return undefined;
}

function fullWidth(value: string): string {
  return Array.from(value)
    .map((character) => {
      const code = character.codePointAt(0)!;
      if (code === 32) return "　";
      return code >= 33 && code <= 126 ? String.fromCodePoint(code + 0xfee0) : character;
    })
    .join("");
}

function byteLength(value: string, encoding: FormulaTextEncoding | undefined): number {
  return Array.from(value).reduce(
    (length, character) => length + byteCharacterSize(character, encoding),
    0,
  );
}

function substringBytes(
  value: string,
  start: number,
  count: number,
  encoding: FormulaTextEncoding | undefined,
): string {
  let byteIndex = 0;
  let output = "";
  for (const character of value) {
    const size = byteCharacterSize(character, encoding);
    if (byteIndex + 1 > start && byteIndex + size <= start + count) output += character;
    byteIndex += size;
  }
  return output;
}

function characterCode(value: string, encoding: FormulaTextEncoding | undefined): number {
  const character = Array.from(value)[0] ?? "";
  const bytes = encoding?.encode?.(character);
  if (bytes?.length) return bytes[0]!;
  return character.codePointAt(0) ?? 0;
}

function characterFromCode(value: number, encoding: FormulaTextEncoding | undefined): string {
  if (encoding?.decode) return encoding.decode(new Uint8Array([value & 0xff]));
  return String.fromCodePoint(value);
}

function compileRegex(pattern: string, caseInsensitive: boolean, global = true): RegExp {
  try {
    return new RegExp(pattern, `${regexFlags(caseInsensitive)}${global ? "g" : ""}`);
  } catch {
    throw new FormulaValueError("#VALUE!");
  }
}

function numberFormat(value: number, format: string, dateSystem: "1900" | "1904" = "1900"): string {
  if (format.toUpperCase() === "GENERAL") return String(value);
  const sections = splitFormatSections(format);
  const section = selectFormatSection(sections, value);
  const unconditioned = section
    .replace(/^\[(?:(?:black|blue|cyan|green|magenta|red|white|yellow)|color\s*\d+)\]/giu, "")
    .replace(/^\[[^!\]]+!/u, "");
  if (/y+|m+|d+|h+|s+|am\/pm|\[[hms]\]/iu.test(unconditioned))
    return formatDateSerial(Math.abs(value), unconditioned, dateSystem);
  const negativeSection =
    value < 0 &&
    splitFormatSections(format).length > 1 &&
    formatCondition(sections[0] ?? "") === undefined;
  return formatNumber(negativeSection ? Math.abs(value) : value, section);
}

function selectFormatSection(sections: string[], value: number): string {
  const withConditions = sections.map((format) => ({ format, condition: formatCondition(format) }));
  const conditional = withConditions.filter((section) => section.condition !== undefined);
  if (conditional.length > 0) {
    for (const section of conditional) {
      if (section.condition && conditionMatches(section.condition, value))
        return cleanFormatSection(section.format);
    }
    return cleanFormatSection(conditional[0]!.format);
  }
  return value > 0
    ? (sections[0] ?? "General")
    : value < 0
      ? (sections[1] ?? sections[0] ?? "General")
      : (sections[2] ?? sections[0] ?? "General");
}

function cleanFormatSection(format: string): string {
  return format
    .replace(/^\[(?:<>|>=|<=|>|<|=)[^[\]]+\]/u, "")
    .replace(/^\[(?:(?:black|blue|cyan|green|magenta|red|white|yellow)|color\s*\d+)\]/giu, "")
    .replace(/^\[[^!\]]+!/u, "");
}

function formatCondition(format: string): { operator: string; value: number } | undefined {
  const match = /^\[(<>|>=|<=|>|<|=)([^[\]]+)\]/u.exec(format);
  if (!match) return undefined;
  const target = Number(match[2]);
  if (!Number.isFinite(target)) return undefined;
  return { operator: match[1]!, value: target };
}

function conditionMatches(condition: { operator: string; value: number }, value: number): boolean {
  switch (condition.operator) {
    case ">":
      return value > condition.value;
    case "<":
      return value < condition.value;
    case ">=":
      return value >= condition.value;
    case "<=":
      return value <= condition.value;
    case "<>":
      return value !== condition.value;
    default:
      return value === condition.value;
  }
}

function splitFormatSections(format: string): string[] {
  const sections: string[] = [];
  let current = "";
  let quoted = false;
  for (let index = 0; index < format.length; index += 1) {
    const character = format[index]!;
    if (character === '"') {
      quoted = !quoted;
      current += character;
    } else if (character === "\\" && index + 1 < format.length) {
      current += character + format[index + 1]!;
      index += 1;
    } else if (character === ";" && !quoted) {
      sections.push(current);
      current = "";
    } else current += character;
  }
  sections.push(current);
  return sections;
}

function formatNumber(input: number, format: string): string {
  const fractionPattern = /([0#?]*)\s*([0#?]+)\/([0#?]+)/u.exec(format);
  if (fractionPattern) return formatFraction(input, fractionPattern);
  const scientificPattern = /([0#?]+)(?:\.([0#?]*))?E([+-])(0+)/iu.exec(format);
  if (scientificPattern) return formatScientific(input, scientificPattern);
  const literalParts: string[] = [];
  const literalMarker = (index: number): string => String.fromCharCode(0xe000 + index);
  let normalized = format.replace(/"([^"]*)"|\\(.)/gu, (_, quoted, escaped) => {
    const literal = quoted ?? escaped;
    literalParts.push(literal);
    return literalMarker(literalParts.length - 1);
  });
  const percentCount = (normalized.match(/%/gu) ?? []).length;
  normalized = normalized.replace(/%/gu, () => {
    literalParts.push("%");
    return literalMarker(literalParts.length - 1);
  });
  normalized = Array.from(normalized)
    .map((literal) => {
      const marker = literal.codePointAt(0)!;
      if (marker >= 0xe000 && marker <= 0xf8ff) return literal;
      if (/^[0#?,.;Ee\s]$/u.test(literal)) return literal;
      literalParts.push(literal);
      return literalMarker(literalParts.length - 1);
    })
    .join("");
  const scalingCommas =
    /^(?:[0#?]*,)*[0#?]*(?:\.[0#?]*)?(,)$/.exec(normalized.replace(/%/gu, ""))?.[1]?.length ?? 0;
  const commaScaled = (Math.abs(input) * 100 ** percentCount) / 1000 ** scalingCommas;
  const decimal = normalized.split(".")[1];
  const decimalPlaceholders = decimal?.match(/[0#?]/gu) ?? [];
  const decimals = decimalPlaceholders.length;
  const minimumDecimals = decimal?.match(/0/gu)?.length ?? 0;
  const fractionFormat = /([0#?]*)\s*([0#?]+)\/([0#?]+)/u.exec(normalized);
  if (fractionFormat) return formatFraction(input, fractionFormat);
  const minimumInteger = (normalized.split(".")[0]?.match(/0/gu) ?? []).length;
  const grouping = (normalized.split(".")[0] ?? "").includes(",");
  const fixed = commaScaled.toFixed(decimals);
  let [integer, fraction = ""] = fixed.split(".");
  if (integer === "0" && minimumInteger === 0) integer = "";
  while (integer!.length < minimumInteger) integer = `0${integer}`;
  if (grouping) integer = integer!.replace(/\B(?=(\d{3})+(?!\d))/gu, ",");
  fraction = fraction!.replace(/0+$/u, "").padEnd(minimumDecimals, "0");
  let numeric = decimals ? `${integer}.${fraction}` : integer;
  const scientific = /E([+-])(0+)/iu.exec(normalized);
  if (scientific) {
    const exponent = Math.floor(Math.log10(Math.abs(commaScaled || 1)));
    const mantissa = commaScaled / 10 ** exponent;
    const exponentDigits = Math.max(scientific[2]!.length, String(Math.abs(exponent)).length);
    numeric = `${mantissa.toFixed(Math.max(0, decimals))}E${exponent < 0 ? "-" : "+"}${String(
      Math.abs(exponent),
    ).padStart(exponentDigits, "0")}`;
  }
  const firstNumber = normalized.search(/[0#?]/u);
  const lastNumber = Math.max(
    normalized.lastIndexOf("0"),
    normalized.lastIndexOf("#"),
    normalized.lastIndexOf("?"),
  );
  const prefixText = literalParts
    .filter(
      (_, index) =>
        normalized.indexOf(literalMarker(index)) >= 0 &&
        normalized.indexOf(literalMarker(index)) < firstNumber,
    )
    .join("");
  const suffixText = literalParts
    .filter(
      (_, index) =>
        normalized.indexOf(literalMarker(index)) >= 0 &&
        normalized.indexOf(literalMarker(index)) > lastNumber,
    )
    .join("");
  return `${prefixText}${numeric}${suffixText}`;
}

function chineseNumber(value: number, uppercase: boolean): string {
  if (!Number.isFinite(value)) throw new FormulaValueError("#VALUE!");
  if (value === 0) return "零";
  const negative = value < 0;
  const source = Math.abs(value).toFixed(10).replace(/0+$/u, "");
  const [integer, fraction = ""] = source.split(".");
  const digits = uppercase ? "零壹贰叁肆伍陆柒捌玖" : "零一二三四五六七八九";
  const units = uppercase ? ["", "拾", "佰", "仟"] : ["", "十", "百", "千"];
  const groups = ["", "万", "亿", "兆"];
  let output = "";
  const chunks = integer!.match(/\d{1,4}(?=(\d{4})*$)/gu) ?? [];
  chunks.forEach((chunk, groupIndex) => {
    const groupUnits = groups[chunks.length - groupIndex - 1] ?? "";
    let group = "";
    let sawZero = false;
    Array.from(chunk).forEach((character, digitIndex) => {
      const digit = Number(character);
      const unit = units[chunk.length - digitIndex - 1]!;
      if (digit === 0) {
        sawZero = group !== "";
      } else {
        if (sawZero) group += digits[0]!;
        sawZero = false;
        group +=
          digit === 1 && unit === "十" && !uppercase && group === "" ? unit : digits[digit]! + unit;
      }
    });
    if (group) output += group + groupUnits;
  });
  if (fraction) output += "点";
  for (const character of fraction.replace(/0+$/u, "")) output += digits[Number(character)]!;
  return (negative ? "负" : "") + (output || digits[0]!);
}

function formatFraction(input: number, match: RegExpExecArray): string {
  const sign = input < 0 ? "-" : "";
  const value = Math.abs(input);
  const integerPattern = match[1] ?? "";
  const maximumDenominator = Number(match[3]!.replace(/\?/gu, "9")) || 1;
  let numerator = 0;
  let denominator = 1;
  let error = Number.POSITIVE_INFINITY;
  for (
    let candidateDenominator = 1;
    candidateDenominator <= maximumDenominator;
    candidateDenominator += 1
  ) {
    const candidateNumerator = Math.round(value * candidateDenominator);
    const candidateError = Math.abs(value - candidateNumerator / candidateDenominator);
    if (
      candidateError < error ||
      (candidateError === error && candidateDenominator < denominator)
    ) {
      numerator = candidateNumerator;
      denominator = candidateDenominator;
      error = candidateError;
    }
  }
  const whole = Math.floor(numerator / denominator);
  numerator %= denominator;
  const greatestCommonDivisor = (left: number, right: number): number =>
    right === 0 ? left : greatestCommonDivisor(right, left % right);
  const divisor = greatestCommonDivisor(numerator, denominator) || 1;
  const reducedNumerator = numerator / divisor;
  const reducedDenominator = denominator / divisor;
  const minimumInteger = (integerPattern.match(/0/gu) ?? []).length;
  const wholeText =
    whole > 0 || minimumInteger > 0 ? String(whole).padStart(minimumInteger, "0") : "";
  const fraction = reducedNumerator === 0 ? "" : `${reducedNumerator}/${reducedDenominator}`;
  const separator = /\s/u.test(match[0]) ? " " : "";
  return sign + [wholeText, fraction].filter(Boolean).join(separator);
}

function formatScientific(input: number, match: RegExpExecArray): string {
  const decimals = match[2]?.length ?? 0;
  let exponent = input === 0 ? 0 : Math.floor(Math.log10(Math.abs(input)));
  let mantissa = input / 10 ** exponent;
  mantissa = Number(mantissa.toFixed(decimals));
  if (Math.abs(mantissa) >= 10) {
    mantissa /= 10;
    exponent += 1;
  }
  const sign = match[3] === "+" || exponent < 0 ? (exponent < 0 ? "-" : "+") : "";
  return `${mantissa.toFixed(decimals)}E${sign}${String(Math.abs(exponent)).padStart(
    match[4]!.length,
    "0",
  )}`;
}

function formatDateSerial(
  serial: number,
  format: string,
  dateSystem: "1900" | "1904" = "1900",
): string {
  const date = toUtcDate(serial, dateSystem);
  const use12Hour = /am\/pm/iu.test(format);
  const originalHours = date.getUTCHours();
  const hours12 = originalHours % 12 === 0 ? 12 : originalHours % 12;
  const meridiem = originalHours < 12 ? "AM" : "PM";
  const weekdayNames = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
  ];
  const weekday = weekdayNames[excelWeekday(serial, dateSystem)]!;
  const replace = (pattern: string, offset: number): string => {
    const lower = pattern.toLowerCase();
    const before = format.slice(0, offset);
    const after = format.slice(offset + pattern.length);
    if (lower === "yyyy") return String(date.getUTCFullYear());
    if (lower === "yy") return String(date.getUTCFullYear() % 100).padStart(2, "0");
    if (lower === "mmmm") return date.toLocaleString("en-US", { month: "long", timeZone: "UTC" });
    if (lower === "mmm") return date.toLocaleString("en-US", { month: "short", timeZone: "UTC" });
    if (lower === "mm" || lower === "m") {
      const beforeTimeCodes = before.replace(/\[[^\]]+\]/gu, "");
      const value = /h/iu.test(beforeTimeCodes)
        ? date.getUTCHours()
        : /s/iu.test(after)
          ? date.getUTCMinutes()
          : date.getUTCMonth() + 1;
      return lower === "mm" ? String(value).padStart(2, "0") : String(value);
    }
    if (lower === "dddd") return weekday;
    if (lower === "ddd") return weekday.slice(0, 3);
    if (lower === "dd") return String(date.getUTCDate()).padStart(2, "0");
    if (lower === "d") return String(date.getUTCDate());
    if (lower === "hh") return String(use12Hour ? hours12 : originalHours).padStart(2, "0");
    if (lower === "h") return String(use12Hour ? hours12 : originalHours);
    if (lower === "ss") return String(date.getUTCSeconds()).padStart(2, "0");
    if (lower === "s") return String(date.getUTCSeconds());
    if (/^\[[hms]\]$/iu.test(pattern)) {
      if (pattern.toLowerCase() === "[h]") return String(Math.floor(serial * 24));
      if (pattern.toLowerCase() === "[m]") return String(Math.floor(serial * 1440));
      return String(Math.floor(serial * 86400));
    }
    return pattern;
  };
  return format
    .replace(/am\/pm/giu, meridiem)
    .replace(
      /\[[hms]\]|yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s/giu,
      (pattern, offset: number) => replace(pattern, offset),
    );
}

function regexFlags(caseInsensitive: boolean): string {
  return caseInsensitive ? "iu" : "u";
}

function expandRegexReplacement(match: RegExpExecArray, replacement: string): string {
  return replacement.replace(/\$\$|\$(?:<([^>]+)>|([1-9]\d*))/gu, (token, named, indexed) => {
    if (token === "$$") return "$";
    if (named !== undefined) return match.groups?.[named] ?? "";
    return match[Number(indexed)] ?? "";
  });
}

function strictText(value: FormulaValue): string {
  if (!isMatrix(value)) return scalarText(value);
  return `{${value
    .map((row) =>
      row
        .map((entry) =>
          typeof entry === "string"
            ? JSON.stringify(entry)
            : typeof entry === "boolean"
              ? entry
                ? "TRUE"
                : "FALSE"
              : scalarText(entry),
        )
        .join(","),
    )
    .join(";")}}`;
}

function delimiterMatches(
  text: string,
  delimiters: FormulaValue,
  matchMode: number,
  matchEnd: boolean,
): Array<{ start: number; end: number }> {
  const candidates = (isMatrix(delimiters) ? delimiters.flat() : [delimiters]).map(scalarText);
  const matches: Array<{ start: number; end: number }> = [];
  const source = foldUnicodeText(text, matchMode === 0);
  for (const delimiter of candidates) {
    if (!delimiter) continue;
    const target = foldUnicodeText(delimiter, matchMode === 0);
    let index = source.value.indexOf(target.value);
    while (index >= 0) {
      matches.push({
        start: source.boundaries[index]!,
        end: source.boundaries[index + target.value.length]!,
      });
      index = source.value.indexOf(target.value, index + 1);
    }
  }
  if (matchEnd && candidates.some((delimiter) => delimiter === "")) {
    matches.push({ start: 0, end: 0 }, { start: text.length, end: text.length });
  }
  return matches.sort((left, right) => left.start - right.start || right.end - left.end);
}

function foldUnicodeText(
  text: string,
  caseInsensitive: boolean,
): {
  value: string;
  boundaries: number[];
} {
  const chunks: string[] = [];
  const boundaries: number[] = [0];
  let index = 0;
  while (index < text.length) {
    const codePoint = text.codePointAt(index)!;
    const chunk = String.fromCodePoint(codePoint);
    chunks.push(
      caseInsensitive && chunk === "İ" ? "i" : caseInsensitive ? chunk.toLowerCase() : chunk,
    );
    const folded = chunks.at(-1)!;
    for (let offset = 1; offset < folded.length; offset += 1) boundaries.push(index);
    index += codePoint > 0xffff ? 2 : 1;
    boundaries.push(index);
  }
  return { value: chunks.join(""), boundaries };
}

function splitByDelimiters(
  text: string,
  delimiters: FormulaValue,
  ignoreEmpty: boolean,
  matchMode: number,
  matchEnd: boolean,
): string[] {
  const matches = delimiterMatches(text, delimiters, matchMode, matchEnd);
  const parts: string[] = [];
  let cursor = 0;
  for (const match of matches) {
    if (match.start < cursor) continue;
    parts.push(text.slice(cursor, match.start));
    cursor = match.end;
  }
  parts.push(text.slice(cursor));
  return ignoreEmpty ? parts.filter((part) => part !== "") : parts;
}

export const textFunctions: FormulaFunctionTable = new Map<string, FormulaFunctionImplementation>([
  [
    "ARRAYTOTEXT",
    ([value, format]) =>
      format !== undefined && scalarNumber(format) === 1
        ? strictText(value)
        : (isMatrix(value) ? value.flat() : [value]).map(scalarText).join(", "),
  ],
  ["ASC", ([value]) => scalarText(value).normalize("NFKC")],
  ["BAHTTEXT", ([value]) => `${scalarNumber(value).toLocaleString("th-TH")} บาท`],
  [
    "CHAR",
    ([value], context) => {
      const code = scalarNumber(value);
      if (!Number.isInteger(code) || code < 1 || code > 255) throw new FormulaValueError("#VALUE!");
      return characterFromCode(code, context.workbook.textEncoding);
    },
  ],
  [
    "CLEAN",
    ([value]) =>
      Array.from(scalarText(value))
        .filter((character) => {
          const code = character.codePointAt(0)!;
          return code >= 32 && code !== 127;
        })
        .join(""),
  ],
  ["CODE", ([value], context) => characterCode(scalarText(value), context.workbook.textEncoding)],
  ["CONCAT", (args) => args.flatMap(values).map(scalarText).join("")],
  ["CONCATENATE", (args) => args.map(scalarText).join("")],
  ["DBCS", ([value]) => fullWidth(scalarText(value))],
  ["EXACT", ([left, right]) => scalarText(left) === scalarText(right)],
  [
    "FIND",
    ([needleValue, text, start]) => {
      const source = scalarText(text);
      const startNumber = start === undefined ? 1 : scalarNumber(start);
      if (startNumber < 1 || startNumber > Array.from(source).length)
        throw new FormulaValueError("#VALUE!");
      const index = codePointIndexOf(source, scalarText(needleValue), startNumber);
      return index < 0 ? "#VALUE!" : index + 1;
    },
  ],
  [
    "FINDB",
    ([needle, text, start], context) => {
      const source = scalarText(text);
      const encoding = context.workbook.textEncoding;
      const startNumber = start === undefined ? 1 : scalarNumber(start);
      if (startNumber < 1 || startNumber > byteLength(source, encoding))
        throw new FormulaValueError("#VALUE!");
      const startIndex = characterIndexAtByteBoundary(source, startNumber, encoding);
      if (startIndex === undefined) throw new FormulaValueError("#VALUE!");
      const characterIndex = codePointIndexOf(source, scalarText(needle), startIndex + 1);
      return characterIndex < 0
        ? "#VALUE!"
        : byteOffsetAtCharacterIndex(source, characterIndex, encoding) + 1;
    },
  ],
  [
    "FIXED",
    ([value, decimals, noCommas]) =>
      scalarNumber(value).toLocaleString("en-US", {
        minimumFractionDigits: decimals === undefined ? 2 : scalarNumber(decimals),
        maximumFractionDigits: decimals === undefined ? 2 : scalarNumber(decimals),
        useGrouping: !(noCommas !== undefined && scalarBoolean(noCommas)),
      }),
  ],
  ["JIS", ([value]) => fullWidth(scalarText(value))],
  [
    "LEFT",
    ([value, count]) => {
      const length = count === undefined ? 1 : scalarNumber(count);
      if (length < 0) return "#VALUE!";
      return Array.from(scalarText(value)).slice(0, length).join("");
    },
  ],
  [
    "LEFTB",
    ([value, count], context) =>
      substringBytes(
        scalarText(value),
        0,
        count === undefined ? 1 : scalarNumber(count),
        context.workbook.textEncoding,
      ),
  ],
  ["LEN", ([value]) => Array.from(scalarText(value)).length],
  ["LENB", ([value], context) => byteLength(scalarText(value), context.workbook.textEncoding)],
  ["LOWER", ([value]) => scalarText(value).toLowerCase()],
  [
    "MID",
    ([value, start, count]) => {
      const startNumber = scalarNumber(start);
      const countNumber = scalarNumber(count);
      if (startNumber < 1 || countNumber < 0) throw new FormulaValueError("#VALUE!");
      return Array.from(scalarText(value))
        .slice(startNumber - 1, startNumber - 1 + countNumber)
        .join("");
    },
  ],
  [
    "MIDB",
    ([value, start, count], context) => {
      const startNumber = scalarNumber(start);
      const countNumber = scalarNumber(count);
      if (startNumber < 1 || countNumber < 0) throw new FormulaValueError("#VALUE!");
      return substringBytes(
        scalarText(value),
        startNumber - 1,
        countNumber,
        context.workbook.textEncoding,
      );
    },
  ],
  [
    "NUMBERVALUE",
    ([value, decimal, group]) => {
      let text = scalarText(value);
      if (group !== undefined)
        text = text.replace(new RegExp(escapeRegex(scalarText(group)), "gu"), "");
      if (decimal !== undefined) text = text.replace(scalarText(decimal), ".");
      const parsed = Number(text);
      return Number.isFinite(parsed) ? parsed : "#VALUE!";
    },
  ],
  [
    "PROPER",
    ([value]) =>
      scalarText(value).replace(
        /(^|[^\p{L}\p{N}])(\p{L})/gu,
        (_, separator: string, character: string) => separator + character.toUpperCase(),
      ),
  ],
  [
    "REGEX",
    ([text, pattern, mode]) =>
      applyRegex(
        scalarText(text),
        scalarText(pattern),
        mode === undefined ? "extract" : scalarText(mode),
        false,
      ),
  ],
  [
    "REGEXEXTRACT",
    ([text, pattern, mode, caseInsensitive]) => {
      const selected = mode === undefined ? 0 : scalarNumber(mode);
      return applyRegex(
        scalarText(text),
        scalarText(pattern),
        selected === 1 ? "all" : selected === 2 ? "groups" : "extract",
        caseInsensitive !== undefined && scalarNumber(caseInsensitive) === 1,
      );
    },
  ],
  [
    "REGEXREPLACE",
    ([text, pattern, replacement, occurrence, caseInsensitive]) => {
      const expression = compileRegex(
        scalarText(pattern),
        caseInsensitive !== undefined && scalarNumber(caseInsensitive) === 1,
      );
      const selected = occurrence === undefined ? 0 : scalarNumber(occurrence);
      const source = scalarText(text);
      if (selected === 0) return source.replace(expression, scalarText(replacement ?? ""));
      expression.lastIndex = 0;
      let count = 0;
      let match = expression.exec(source);
      while (match) {
        count += 1;
        if (count === selected)
          return (
            source.slice(0, match.index) +
            expandRegexReplacement(match, scalarText(replacement ?? "")) +
            source.slice(match.index + match[0].length)
          );
        if (!expression.global) break;
        match = expression.exec(source);
      }
      return source;
    },
  ],
  [
    "REGEXTEST",
    ([text, pattern, caseInsensitive]) =>
      compileRegex(
        scalarText(pattern),
        caseInsensitive !== undefined && scalarNumber(caseInsensitive) === 1,
        false,
      ).test(scalarText(text)),
  ],
  [
    "REPLACE",
    ([text, start, count, replacement]) => {
      const startNumber = scalarNumber(start);
      const countNumber = scalarNumber(count);
      if (startNumber < 1 || countNumber < 0) throw new FormulaValueError("#VALUE!");
      const characters = Array.from(scalarText(text));
      const index = startNumber - 1;
      return [
        ...characters.slice(0, index),
        ...Array.from(scalarText(replacement)),
        ...characters.slice(index + scalarNumber(count)),
      ].join("");
    },
  ],
  [
    "REPLACEB",
    ([text, start, count, replacement], context) => {
      const encoding = context.workbook.textEncoding;
      const startNumber = scalarNumber(start);
      const countNumber = scalarNumber(count);
      if (startNumber < 1 || countNumber < 0) throw new FormulaValueError("#VALUE!");
      const before = substringBytes(scalarText(text), 0, startNumber - 1, encoding);
      const after = substringBytes(
        scalarText(text),
        startNumber - 1 + countNumber,
        Number.MAX_SAFE_INTEGER,
        encoding,
      );
      return before + scalarText(replacement) + after;
    },
  ],
  [
    "REPT",
    ([value, count]) => {
      const repeat = scalarNumber(count);
      if (repeat < 0) throw new FormulaValueError("#VALUE!");
      return scalarText(value).repeat(repeat);
    },
  ],
  [
    "RIGHT",
    ([value, count]) => {
      const length = count === undefined ? 1 : scalarNumber(count);
      if (length < 0) return "#VALUE!";
      const characters = Array.from(scalarText(value));
      return characters.slice(Math.max(0, characters.length - length)).join("");
    },
  ],
  [
    "RIGHTB",
    ([value, count], context) => {
      const text = scalarText(value);
      const size = count === undefined ? 1 : scalarNumber(count);
      return substringBytes(
        text,
        Math.max(0, byteLength(text, context.workbook.textEncoding) - size),
        size,
        context.workbook.textEncoding,
      );
    },
  ],
  [
    "SEARCH",
    ([needleValue, text, start]) => {
      const startNumber = start === undefined ? 1 : scalarNumber(start);
      if (startNumber < 1 || startNumber > Array.from(scalarText(text)).length)
        throw new FormulaValueError("#VALUE!");
      const index = wildcardIndex(scalarText(needleValue), scalarText(text), startNumber);
      return index < 0 ? "#VALUE!" : index;
    },
  ],
  [
    "SEARCHB",
    ([needleValue, text, start], context) => {
      const source = scalarText(text);
      const encoding = context.workbook.textEncoding;
      const startNumber = start === undefined ? 1 : scalarNumber(start);
      if (startNumber < 1 || startNumber > byteLength(source, encoding))
        throw new FormulaValueError("#VALUE!");
      const startIndex = characterIndexAtByteBoundary(source, startNumber, encoding);
      if (startIndex === undefined) throw new FormulaValueError("#VALUE!");
      const index = wildcardIndex(scalarText(needleValue), source, startIndex + 1);
      return index < 0 ? "#VALUE!" : byteOffsetAtCharacterIndex(source, index - 1, encoding) + 1;
    },
  ],
  ["STRINGS", (args) => args.flatMap(values).map(scalarText).join("")],
  [
    "SUBSTITUTE",
    ([text, oldText, newText, instance]) => {
      const source = scalarText(text);
      const search = scalarText(oldText);
      if (!search) return source;
      const replacement = scalarText(newText);
      if (instance === undefined) return source.split(search).join(replacement);
      const occurrence = scalarNumber(instance);
      if (occurrence < 1) throw new FormulaValueError("#VALUE!");
      let index = 0;
      for (let count = 0; count < occurrence; count += 1) {
        index = source.indexOf(search, count === occurrence - 1 ? index : index + search.length);
        if (index < 0) return source;
      }
      return source.slice(0, index) + replacement + source.slice(index + search.length);
    },
  ],
  ["T", ([value]) => (Array.isArray(value) ? "" : typeof value === "string" ? value : "")],
  [
    "TRIM",
    ([value]) =>
      scalarText(value)
        .split(" ")
        .filter((part) => part !== "")
        .join(" "),
  ],
  [
    "TEXT",
    ([value, format], context) =>
      typeof value === "string" && !Number.isFinite(Number(value))
        ? scalarText(value)
        : numberFormat(
            scalarNumber(value),
            scalarText(format),
            context.dateSystem ?? context.workbook.dateSystem ?? "1900",
          ),
  ],
  [
    "TEXTAFTER",
    ([text, delimiter, instance, matchMode, matchEnd, notFound]) =>
      splitDelimiter(
        scalarText(text),
        delimiter,
        instance === undefined ? 1 : scalarNumber(instance),
        "after",
        matchMode === undefined ? 0 : scalarNumber(matchMode),
        matchEnd !== undefined && scalarBoolean(matchEnd),
        notFound ?? "#N/A",
      ),
  ],
  [
    "TEXTBEFORE",
    ([text, delimiter, instance, matchMode, matchEnd, notFound]) =>
      splitDelimiter(
        scalarText(text),
        delimiter,
        instance === undefined ? 1 : scalarNumber(instance),
        "before",
        matchMode === undefined ? 0 : scalarNumber(matchMode),
        matchEnd !== undefined && scalarBoolean(matchEnd),
        notFound ?? "#N/A",
      ),
  ],
  [
    "TEXTJOIN",
    ([delimiter, ignoreEmpty, ...groups]) => {
      const separator = scalarText(delimiter);
      const entries = groups.flatMap((group) => values(group)).map(scalarText);
      return (
        ignoreEmpty === undefined || scalarBoolean(ignoreEmpty)
          ? entries.filter((entry) => entry !== "")
          : entries
      ).join(separator);
    },
  ],
  [
    "TEXTSPLIT",
    ([text, columnDelimiter, rowDelimiter, ignoreEmpty, matchMode, pad]) => {
      const ignore = ignoreEmpty !== undefined && scalarBoolean(ignoreEmpty);
      const mode = matchMode === undefined ? 0 : scalarNumber(matchMode);
      const padValue = pad !== undefined && !isMatrix(pad) ? pad : formulaError("#N/A");
      const rows =
        rowDelimiter === undefined
          ? [splitByDelimiters(scalarText(text), columnDelimiter, ignore, mode, false)]
          : splitByDelimiters(scalarText(text), rowDelimiter, ignore, mode, false).map((row) =>
              splitByDelimiters(row, columnDelimiter, false, mode, false),
            );
      const columns = rows.reduce((maximum, row) => Math.max(maximum, row.length), 0);
      assertFormulaCellCount(rows.length, columns);
      return rows.map((row) => {
        return Array.from({ length: columns }, (_, index) => row[index] ?? padValue);
      });
    },
  ],
  [
    "UNICHAR",
    ([value]) => {
      const code = scalarNumber(value);
      if (!Number.isInteger(code) || code < 1 || code > 1114111)
        throw new FormulaValueError("#VALUE!");
      return String.fromCodePoint(code);
    },
  ],
  ["UNICODE", ([value]) => scalarText(value).codePointAt(0) ?? 0],
  ["UPPER", ([value]) => scalarText(value).toUpperCase()],
  [
    "NUMBERSTRING",
    ([value, type]) => {
      const selected = type === undefined ? 1 : scalarNumber(type);
      if (selected === 3) return Math.trunc(scalarNumber(value)).toString();
      if (selected !== 1 && selected !== 2) return "#VALUE!";
      return chineseNumber(scalarNumber(value), selected === 2);
    },
  ],
  [
    "VALUE",
    ([value], context) => {
      const source = scalarText(value);
      const percentCount = (source.match(/%/gu) ?? []).length;
      const text = source.replace(/[,$%\s]/gu, "");
      const number = Number(text);
      if (Number.isFinite(number)) return number / 100 ** percentCount;
      const parsed = parseDateText(text);
      if (!parsed) return "#VALUE!";
      return toSerialDate(parsed, context.dateSystem ?? context.workbook.dateSystem);
    },
  ],
  ["VALUETOTEXT", ([value]) => scalarText(value)],
]);

function applyRegex(
  text: string,
  pattern: string,
  mode: string,
  caseInsensitive: boolean,
): FormulaValue {
  let expression: RegExp;
  try {
    expression = new RegExp(pattern, `${regexFlags(caseInsensitive)}g`);
  } catch {
    return "#VALUE!";
  }
  const regexMatches = [...text.matchAll(expression)];
  const matches = regexMatches.map((match) => match[0]);
  if (mode === "test") return matches.length > 0;
  if (mode === "replace") return matches[0] ?? "";
  if (mode === "extractall" || mode === "all") return matches.map((match) => [match]);
  if (mode === "groups") return regexMatches.map((match) => [match[1] ?? ""]);
  return matches[0] ?? "#N/A";
}

function splitDelimiter(
  text: string,
  delimiter: FormulaValue,
  instance: number,
  side: "before" | "after",
  matchMode: number,
  matchEnd: boolean,
  notFound: FormulaValue,
): FormulaValue {
  const matches = delimiterMatches(text, delimiter, matchMode === 1 ? 0 : 1, matchEnd);
  const selected = instance > 0 ? matches[instance - 1] : matches.at(instance);
  if (!selected) return notFound;
  return side === "before" ? text.slice(0, selected.start) : text.slice(selected.end);
}

function escapeRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
}
