import type { FormulaDialect, FormulaSpan } from "./ast";
import type { FormulaErrorValue } from "./value";
import { formulaErrorValues } from "./value";

export type FormulaTokenKind =
  | "boolean"
  | "close-brace"
  | "close-paren"
  | "comma"
  | "end"
  | "error"
  | "identifier"
  | "number"
  | "open-brace"
  | "open-paren"
  | "operator"
  | "semicolon"
  | "spill"
  | "string"
  | "whitespace";

export interface FormulaToken extends FormulaSpan {
  kind: FormulaTokenKind;
  text: string;
  value?: string | number | boolean | FormulaErrorValue;
}

const errorLiterals: readonly FormulaErrorValue[] = formulaErrorValues;

const identifierPattern = /[\p{L}\p{N}_.\\`]/u;
function isDigit(character: string | undefined): boolean {
  return character !== undefined && character >= "0" && character <= "9";
}

export function tokenizeFormula(source: string, dialect: FormulaDialect = "excel"): FormulaToken[] {
  const tokens: FormulaToken[] = [];
  let start = 0;

  while (start < source.length) {
    const character = source[start]!;
    const end = start + 1;

    if (/\s/u.test(character)) {
      let current = start;
      while (current < source.length && /\s/u.test(source[current]!)) current += 1;
      tokens.push({ kind: "whitespace", text: source.slice(start, current), start, end: current });
      start = current;
      continue;
    }

    if (character === `"`) {
      let current = start + 1;
      let value = "";
      let closed = false;
      while (current < source.length) {
        if (source[current] !== `"`) {
          value += source[current]!;
          current += 1;
          continue;
        }
        if (source[current + 1] === `"`) {
          value += `"`;
          current += 2;
          continue;
        }
        current += 1;
        tokens.push({
          kind: "string",
          text: source.slice(start, current),
          start,
          end: current,
          value,
        });
        start = current;
        closed = true;
        break;
      }
      if (closed) continue;
      throw Object.assign(new Error("unterminated string literal"), {
        code: "invalid-string",
        start,
        end: source.length,
        dialect,
      });
    }

    if (character === "#") {
      const literal = errorLiterals.find((candidate) => source.startsWith(candidate, start));
      if (!literal) {
        tokens.push({ kind: "spill", text: "#", start, end });
        start = end;
        continue;
      }
      tokens.push({
        kind: "error",
        text: literal,
        start,
        end: start + literal.length,
        value: literal,
      });
      start += literal.length;
      continue;
    }

    if (isDigit(character) || (character === "." && isDigit(source[start + 1]))) {
      let current = start;
      while (isDigit(source[current])) current += 1;
      if (source[current] === ".") {
        current += 1;
        while (isDigit(source[current])) current += 1;
      }
      if (source[current] === "e" || source[current] === "E") {
        let exponent = current + 1;
        if (source[exponent] === "+" || source[exponent] === "-") exponent += 1;
        if (isDigit(source[exponent])) {
          current = exponent;
          while (isDigit(source[current])) current += 1;
        }
      }
      const text = source.slice(start, current);
      const value = Number(text);
      if (!Number.isFinite(value)) {
        throw Object.assign(new Error(`invalid number "${text}"`), {
          code: "invalid-number",
          start,
          end: current,
          dialect,
        });
      }
      tokens.push({ kind: "number", text, start, end: current, value });
      start = current;
      continue;
    }

    const twoCharacters = source.slice(start, start + 2);
    if (twoCharacters === "<=" || twoCharacters === ">=" || twoCharacters === "<>") {
      tokens.push({ kind: "operator", text: twoCharacters, start, end: start + 2 });
      start += 2;
      continue;
    }

    if (character === "(") {
      tokens.push({ kind: "open-paren", text: character, start, end });
      start = end;
      continue;
    }
    if (character === ")") {
      tokens.push({ kind: "close-paren", text: character, start, end });
      start = end;
      continue;
    }
    if (character === "{") {
      tokens.push({ kind: "open-brace", text: character, start, end });
      start = end;
      continue;
    }
    if (character === "}") {
      tokens.push({ kind: "close-brace", text: character, start, end });
      start = end;
      continue;
    }
    if (character === "," || character === ";") {
      tokens.push({
        kind: character === "," ? "comma" : "semicolon",
        text: character,
        start,
        end,
      });
      start = end;
      continue;
    }
    if ("+-*/^&=<>:%~!#@$|'[]".includes(character)) {
      tokens.push({ kind: "operator", text: character, start, end });
      start = end;
      continue;
    }

    if (identifierPattern.test(character)) {
      let current = start;
      while (current < source.length && identifierPattern.test(source[current]!)) current += 1;
      const text = source.slice(start, current);
      const upper = text.toUpperCase();
      const kind = upper === "TRUE" || upper === "FALSE" ? "boolean" : "identifier";
      tokens.push({
        kind,
        text,
        start,
        end: current,
        value: kind === "boolean" ? upper === "TRUE" : text,
      });
      start = current;
      continue;
    }

    throw Object.assign(new Error(`invalid character "${character}"`), {
      code: "invalid-character",
      start,
      end,
      dialect,
    });
  }

  tokens.push({ kind: "end", text: "", start: source.length, end: source.length });
  return tokens;
}
