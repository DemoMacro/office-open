import type {
  FontOptions,
  RichTextOptions,
  RichTextRunOptions,
  RichTextRunPropertiesOptions,
} from "@office-open/xlsx";

import { XlsParseError } from "../errors";
import { RecordCode, requireBytes, type BiffVersion } from "../records";

interface DecodedBiffString {
  value: string;
  offset: number;
  richRuns?: { characterIndex: number; fontIndex: number }[];
}

function booleanLexicalValue(
  value: string | undefined,
): RichTextRunPropertiesOptions["boldValRaw"] {
  if (value === "true" || value === "false" || value === "1" || value === "0") return value;
  return undefined;
}
export function decodeBiffString(
  body: Uint8Array,
  view: DataView,
  offset: number,
  characterCount: number,
  flags: number,
  encoding = "windows-1252",
): DecodedBiffString {
  let cursor = offset;
  const highByte = (flags & 0x01) !== 0;
  const hasRichRuns = (flags & 0x08) !== 0;
  const hasExtension = (flags & 0x04) !== 0;
  requireBytes(body, cursor + (hasRichRuns ? 2 : 0) + (hasExtension ? 4 : 0), "string");
  const richRunCount = hasRichRuns ? view.getUint16(cursor, true) : 0;
  if (hasRichRuns) cursor += 2;
  const extensionLength = hasExtension ? view.getUint32(cursor, true) : 0;
  if (hasExtension) cursor += 4;
  const byteLength = characterCount * (highByte ? 2 : 1);
  requireBytes(body, cursor + byteLength + richRunCount * 4 + extensionLength, "string");
  const raw = body.subarray(cursor, cursor + byteLength);
  const value = highByte
    ? new TextDecoder("utf-16le").decode(raw)
    : new TextDecoder(encoding).decode(raw);
  cursor += byteLength;
  let richRuns: { characterIndex: number; fontIndex: number }[] | undefined;
  if (richRunCount > 0) {
    richRuns = Array.from({ length: richRunCount }, () => {
      const run = {
        characterIndex: view.getUint16(cursor, true),
        fontIndex: view.getUint16(cursor + 2, true),
      };
      cursor += 4;
      return run;
    });
  }
  cursor += extensionLength;
  return { value, offset: cursor, richRuns };
}

export function richTextValue(
  value: string,
  richRuns: readonly { characterIndex: number; fontIndex: number }[] | undefined,
  fonts: readonly FontOptions[],
): string | RichTextOptions {
  if (!richRuns?.length) return value;
  const runs: RichTextRunOptions[] = [];
  let previousIndex = -1;
  if (richRuns[0]!.characterIndex > 0) {
    runs.push({ text: value.slice(0, richRuns[0]!.characterIndex) });
  }
  for (const [index, richRun] of richRuns.entries()) {
    const start = richRun.characterIndex;
    const end = richRuns[index + 1]?.characterIndex ?? value.length;
    if (start > value.length || start <= previousIndex || end > value.length) {
      throw new XlsParseError("Invalid legacy XLS file: invalid rich-text run position", {
        part: "stream",
        path: "shared-strings/rich-text",
        reason: "invalid-rich-text-run-position",
      });
    }
    if (end > start) {
      const font = fonts[richRun.fontIndex];
      if (!font) {
        throw new XlsParseError("Invalid legacy XLS file: rich-text run references missing font", {
          part: "stream",
          path: "shared-strings/rich-text",
          reason: "missing-rich-text-font",
        });
      }
      runs.push({
        text: value.slice(start, end),
        properties: {
          font: font.font,
          charset: font.charset,
          family: font.family,
          bold: font.bold,
          boldValRaw: booleanLexicalValue(font.boldRaw),
          italic: font.italic,
          italicValRaw: booleanLexicalValue(font.italicRaw),
          underline: font.underline === true ? "single" : font.underline || undefined,
          strike: font.strike,
          strikeValRaw: booleanLexicalValue(font.strikeRaw),
          outline: font.outline,
          outlineValRaw: booleanLexicalValue(font.outlineRaw),
          shadow: font.shadow,
          shadowValRaw: booleanLexicalValue(font.shadowRaw),
          condense: font.condense,
          condenseValRaw: booleanLexicalValue(font.condenseRaw),
          extend: font.extend,
          extendValRaw: booleanLexicalValue(font.extendRaw),
          color: font.color,
          size: font.size,
          sizeValRaw: font.sizeRaw,
          vertAlign: font.vertAlign,
          scheme: font.scheme,
        },
      });
    }
    previousIndex = start;
  }
  return { runs };
}

export function readBiff8String(
  body: Uint8Array,
  offset = 0,
  encoding = "windows-1252",
): DecodedBiffString {
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  requireBytes(body, offset + 3, "string");
  const characterCount = view.getUint16(offset, true);
  const flags = body[offset + 2]!;
  return decodeBiffString(body, view, offset + 3, characterCount, flags, encoding);
}

export function readBiff5String(
  body: Uint8Array,
  offset = 0,
  encoding = "windows-1252",
): { value: string; offset: number } {
  requireBytes(body, offset + 1, "string");
  const characterCount = body[offset]!;
  requireBytes(body, offset + 1 + characterCount, "string");
  const raw = body.subarray(offset + 1, offset + 1 + characterCount);
  return {
    value: new TextDecoder(encoding).decode(raw),
    offset: offset + 1 + characterCount,
  };
}

export function readBiff35String(
  body: Uint8Array,
  offset = 0,
  encoding = "windows-1252",
): { value: string; offset: number } {
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  requireBytes(body, offset + 2, "string");
  const characterCount = view.getUint16(offset, true);
  requireBytes(body, offset + 2 + characterCount, "string");
  return {
    value: new TextDecoder(encoding).decode(body.subarray(offset + 2, offset + 2 + characterCount)),
    offset: offset + 2 + characterCount,
  };
}

export function readBiff2String(
  body: Uint8Array,
  offset = 0,
  encoding = "windows-1252",
): { value: string; offset: number } {
  return readBiff5String(body, offset, encoding);
}

export function readShortBiff8String(
  body: Uint8Array,
  offset = 0,
  encoding = "windows-1252",
): { value: string; offset: number } {
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  requireBytes(body, offset + 2, "string");
  return decodeBiffString(body, view, offset + 2, body[offset]!, body[offset + 1]!, encoding);
}

export function readNoCchBiff8String(
  body: Uint8Array,
  offset: number,
  characterCount: number,
  encoding = "windows-1252",
): { value: string; offset: number } {
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  requireBytes(body, offset + 1, "string");
  return decodeBiffString(body, view, offset + 1, characterCount, body[offset]!, encoding);
}

export function readNoCchBiff5String(
  body: Uint8Array,
  offset: number,
  characterCount: number,
  encoding = "windows-1252",
): { value: string; offset: number } {
  requireBytes(body, offset + characterCount, "string");
  return {
    value: new TextDecoder(encoding).decode(body.subarray(offset, offset + characterCount)),
    offset: offset + characterCount,
  };
}

export function encodingForCodepage(codepage: number): string {
  if (codepage === 1200) return "utf-16le";
  if (codepage === 65001) return "utf-8";
  try {
    return new TextDecoder(codepageLabel(codepage), { fatal: false }).encoding;
  } catch {
    return "windows-1252";
  }
}

export function codepageLabel(codepage: number): string {
  const labels: Record<number, string> = {
    932: "shift_jis",
    936: "gbk",
    949: "euc-kr",
    950: "big5",
    1250: "windows-1250",
    1251: "windows-1251",
    1252: "windows-1252",
    1253: "windows-1253",
    1254: "windows-1254",
    1255: "windows-1255",
    1256: "windows-1256",
    1257: "windows-1257",
    1258: "windows-1258",
  };
  return labels[codepage] ?? "windows-1252";
}

export function detectBiffVersion(stream: Uint8Array): {
  version: BiffVersion;
  workbookLayout: "standard" | "biff4w";
} {
  if (stream.byteLength < 6)
    throw new XlsParseError("Invalid legacy XLS file: workbook stream is truncated");
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  const code = view.getUint16(0, true);
  if (code === 0) throw new XlsParseError("Invalid legacy XLS file: workbook stream is truncated");
  if (code === RecordCode.EndOfFile) return { version: 8, workbookLayout: "standard" };
  if (code === 0x0009) return { version: 2, workbookLayout: "standard" };
  if (code === 0x0209)
    return {
      version: view.getUint16(4, true) === 0x0200 ? 3 : 2,
      workbookLayout: "standard",
    };
  if (code === 0x0409) {
    const bodyVersion = view.getUint16(4, true);
    const streamType = view.getUint16(6, true);
    if (bodyVersion !== 0x0400 && bodyVersion !== 0)
      return { version: 2, workbookLayout: "standard" };
    if (streamType !== 0x0010 && streamType !== 0x0100)
      return { version: 2, workbookLayout: "standard" };
    return { version: 4, workbookLayout: streamType === 0x0100 ? "biff4w" : "standard" };
  }
  if (code !== 0x0809) return { version: 8, workbookLayout: "standard" };
  const recordVersion = view.getUint16(4, true);
  if (recordVersion !== 0x0500 && recordVersion !== 0x0600)
    throw new XlsParseError(`Unsupported legacy XLS BIFF version: ${recordVersion}`);
  return { version: recordVersion === 0x0600 ? 8 : 5, workbookLayout: "standard" };
}

export const BEGIN_OF_FILE_CODES = new Set([0x0009, 0x0209, 0x0409, 0x0809]);
export const PLAINTEXT_ENCRYPTION_CODES = new Set([
  RecordCode.FilePass,
  ...BEGIN_OF_FILE_CODES,
  0x00e1,
  0x0193,
  0x0195,
  0x0196,
  0x0138,
]);
