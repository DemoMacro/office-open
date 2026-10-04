import { RtfParseError } from "../errors";

export type RtfToken =
  | { kind: "group-start"; position: number }
  | { kind: "group-end"; position: number }
  | { kind: "control"; word?: string; symbol?: string; param?: number; position: number }
  | { kind: "text"; value: string; position: number }
  | { kind: "hex"; value: string; position: number };

const CONTROL_WORD = /[a-zA-Z]/;
const HEX_DIGITS = /^[0-9a-fA-F]{2}/;

function readParameter(source: string, start: number): { value?: number; end: number } {
  let end = start;
  if (source[start] === "-") end += 1;
  const digitsStart = end;
  while (end < source.length) {
    const character = source[end];
    if (character === undefined || character < "0" || character > "9") break;
    end += 1;
  }
  if (end === digitsStart) return { end: start };
  return { value: Number(source.slice(start, end)), end };
}

export function tokenizeRtf(source: string): RtfToken[] {
  if (typeof source !== "string") throw new TypeError("RTF source must be a string");

  const tokens: RtfToken[] = [];
  let position = 0;
  let textStart: number | undefined;

  const flushText = (end: number) => {
    if (textStart === undefined) return;
    tokens.push({ kind: "text", value: source.slice(textStart, end), position: textStart });
    textStart = undefined;
  };

  while (position < source.length) {
    const character = source[position];
    if (character !== "{" && character !== "}" && character !== "\\") {
      textStart ??= position;
      position += 1;
      continue;
    }

    flushText(position);
    if (character === "{") {
      tokens.push({ kind: "group-start", position });
      position += 1;
      continue;
    }
    if (character === "}") {
      tokens.push({ kind: "group-end", position });
      position += 1;
      continue;
    }

    position += 1;
    const escaped = source[position];
    if (escaped === undefined) {
      throw new RtfParseError("a control sequence ends at end of input", position - 1, source);
    }
    if (escaped === "\r" || escaped === "\n") {
      position += escaped === "\r" && source[position + 1] === "\n" ? 2 : 1;
      continue;
    }
    if (CONTROL_WORD.test(escaped)) {
      const wordStart = position;
      position += 1;
      while (position < source.length) {
        const character = source[position];
        if (character === undefined || !CONTROL_WORD.test(character)) break;
        position += 1;
      }
      const word = source.slice(wordStart, position);
      let parameter: number | undefined;
      const nextCharacter = source[position];
      if (
        nextCharacter === "-" ||
        (nextCharacter !== undefined && nextCharacter >= "0" && nextCharacter <= "9")
      ) {
        const parameterResult = readParameter(source, position);
        parameter = parameterResult.value;
        position = parameterResult.end;
      }
      if (word === "bin") {
        if (parameter !== undefined && parameter < 0) {
          throw new RtfParseError(
            "\\bin requires a non-negative byte count",
            wordStart - 1,
            source,
          );
        }
        position += parameter ?? 0;
        if (position > source.length) {
          throw new RtfParseError("\\bin data extends past end of input", wordStart - 1, source);
        }
      }
      if (source[position] === " ") position += 1;
      tokens.push({ kind: "control", word, param: parameter, position: wordStart - 1 });
      continue;
    }

    if (escaped === "'") {
      const hex = source.slice(position + 1, position + 3);
      if (!HEX_DIGITS.test(hex)) {
        throw new RtfParseError("\\' requires two hexadecimal digits", position - 1, source);
      }
      tokens.push({ kind: "hex", value: hex, position: position - 1 });
      position += 3;
      continue;
    }

    tokens.push({ kind: "control", symbol: escaped, position: position - 1 });
    position += 1;
  }

  flushText(source.length);
  return tokens;
}
