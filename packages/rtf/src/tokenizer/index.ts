import { RtfParseError } from "../errors";

export type RtfToken =
  | { kind: "group-start"; position: number }
  | { kind: "group-end"; position: number }
  | { kind: "control"; word?: string; symbol?: string; param?: number; position: number }
  | { kind: "text"; value: string; position: number }
  | { kind: "hex"; value: string; position: number }
  | { kind: "binary"; byteLength: number; data: Uint8Array; position: number };

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

export function latin1Source(data: Uint8Array): string {
  const parts: string[] = [];
  const chunkSize = 0x8000;
  for (let offset = 0; offset < data.length; offset += chunkSize) {
    parts.push(String.fromCharCode(...data.subarray(offset, offset + chunkSize)));
  }
  return parts.join("");
}

export function tokenizeRtf(source: string | Uint8Array): RtfToken[] {
  if (typeof source !== "string" && !(source instanceof Uint8Array))
    throw new TypeError("RTF source must be a string or Uint8Array");

  const text = typeof source === "string" ? source : latin1Source(source);
  const bytes =
    typeof source === "string"
      ? Uint8Array.from(source, (character) => character.charCodeAt(0) & 0xff)
      : source;

  const tokens: RtfToken[] = [];
  let position = 0;
  let byteOffset = 0;
  let textStart: number | undefined;

  const flushText = (end: number) => {
    if (textStart === undefined) return;
    tokens.push({ kind: "text", value: text.slice(textStart, end), position: textStart });
    textStart = undefined;
  };

  while (position < text.length) {
    const character = text[position];
    if (character !== "{" && character !== "}" && character !== "\\") {
      textStart ??= position;
      position += 1;
      byteOffset += 1;
      continue;
    }

    flushText(position);
    if (character === "{") {
      tokens.push({ kind: "group-start", position });
      position += 1;
      byteOffset += 1;
      continue;
    }
    if (character === "}") {
      tokens.push({ kind: "group-end", position });
      position += 1;
      byteOffset += 1;
      continue;
    }

    position += 1;
    byteOffset += 1;
    const escaped = text[position];
    if (escaped === undefined) {
      throw new RtfParseError("a control sequence ends at end of input", position - 1, text);
    }
    if (escaped === "\r" || escaped === "\n") {
      position += escaped === "\r" && text[position + 1] === "\n" ? 2 : 1;
      byteOffset = position;
      continue;
    }
    if (CONTROL_WORD.test(escaped)) {
      const wordStart = position;
      position += 1;
      while (position < text.length) {
        const character = text[position];
        if (character === undefined || !CONTROL_WORD.test(character)) break;
        position += 1;
      }
      const word = text.slice(wordStart, position);
      byteOffset = position;
      let parameter: number | undefined;
      const nextCharacter = text[position];
      if (
        nextCharacter === "-" ||
        (nextCharacter !== undefined && nextCharacter >= "0" && nextCharacter <= "9")
      ) {
        const parameterResult = readParameter(text, position);
        parameter = parameterResult.value;
        byteOffset += parameterResult.end - position;
        position = parameterResult.end;
      }
      if (word === "bin") {
        if (parameter !== undefined && parameter < 0) {
          throw new RtfParseError("\\bin requires a non-negative byte count", wordStart - 1, text);
        }
        const byteLength = parameter ?? 0;
        const hasDelimiter = text[position] === " ";
        if (hasDelimiter) position += 1;
        byteOffset += hasDelimiter ? 1 : 0;
        if (byteOffset + byteLength > bytes.length) {
          throw new RtfParseError("\\bin data extends past end of input", wordStart - 1, text);
        }
        const data = bytes.slice(byteOffset, byteOffset + byteLength);
        byteOffset += byteLength;
        position += byteLength;
        tokens.push({ kind: "binary", byteLength, data, position: wordStart - 1 });
        continue;
      }
      if (text[position] === " ") position += 1;
      byteOffset = position;
      tokens.push({ kind: "control", word, param: parameter, position: wordStart - 1 });
      continue;
    }

    if (escaped === "'") {
      const hex = text.slice(position + 1, position + 3);
      if (!HEX_DIGITS.test(hex)) {
        throw new RtfParseError("\\' requires two hexadecimal digits", position - 1, text);
      }
      tokens.push({ kind: "hex", value: hex, position: position - 1 });
      position += 3;
      byteOffset += 3;
      continue;
    }

    tokens.push({ kind: "control", symbol: escaped, position: position - 1 });
    position += 1;
    byteOffset += 1;
  }

  flushText(text.length);
  return tokens;
}
