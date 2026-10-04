import {
  CompoundFileReader,
  decryptLegacyRc4,
  decryptRc4CryptoApi,
  parseLegacyRc4Verifier,
  parseRc4CryptoApiHeader,
  verifyLegacyRc4Password,
  verifyRc4CryptoApiPassword,
} from "@office-open/core";
import type {
  CellOptions,
  ConditionalFormatRule,
  DefinedNameOptions,
  ExternalLinkOptions,
  FontOptions,
  FormulaOptions,
  HyperlinkOptions,
  RowOptions,
  StyleOptions,
  WorkbookOptions,
  WorksheetOptions,
} from "@office-open/xlsx";

import { readRequiredStream } from "../cfb/container";
import { LegacyExcelError } from "../errors";
import { readSummaryInformation } from "../parts/metadata";
import {
  RecordCode,
  isClearedFilePass,
  readRecord,
  recordHeaderSize,
  recordsFrom,
  requireBytes,
  type BiffRecord,
  type BiffVersion,
} from "../records";
import { escherImages, escherPictures } from "../records/escher";
import { decodeFormula } from "../records/formula";

interface BoundSheet {
  readonly position: number;
  readonly state: WorksheetOptions["state"];
  readonly name: string;
}

interface SheetState {
  rows: Map<number, RowOptions>;
  dimension?: string;
  extras: Partial<WorksheetOptions>;
}

interface WorkbookState {
  version: BiffVersion;
  workbookLayout: "standard" | "biff4w";
  stream: Uint8Array;
  sharedStrings: string[];
  sheets: readonly BoundSheet[];
  fonts: FontOptions[];
  numberFormats: Map<number, string>;
  palette: string[];
  date1904: boolean;
  codepage: number;
  externalLinks: ExternalLinkOptions[];
  definedNames: DefinedNameOptions[];
  externSheetNames: Map<number, string>;
  styleTable: Map<number, { fontIndex?: number; numberFormatId?: number }>;
  drawingImages: { data: Uint8Array; type: "png" | "jpg" | "wmf" | "emf" }[];
}

interface FormulaLocation {
  row: number;
  column: number;
  style: number;
  expression?: string;
}

interface TxoState {
  objectId: number;
  characterCount: number;
  runByteCount: number;
  textParts: TxoTextPart[];
  phase: "text" | "runs";
}

interface TxoTextPart {
  readonly encoding: "utf-16le" | "windows-1252";
  readonly data: Uint8Array;
}

interface TxoText {
  text: string;
}

interface PendingConditional {
  reference: string;
  remaining: number;
  rules: ConditionalFormatRule[];
}

function decodeBiffString(
  body: Uint8Array,
  view: DataView,
  offset: number,
  characterCount: number,
  flags: number,
  encoding = "windows-1252",
): { value: string; offset: number } {
  let cursor = offset;
  const highByte = (flags & 0x01) !== 0;
  const richRunCount = (flags & 0x08) !== 0 ? view.getUint16(cursor, true) : 0;
  if ((flags & 0x08) !== 0) cursor += 2;
  const hasExtension = (flags & 0x04) !== 0;
  if (hasExtension) cursor += 4;
  const byteLength = characterCount * (highByte ? 2 : 1);
  requireBytes(body, cursor + byteLength + richRunCount * 4, "string");
  const raw = body.subarray(cursor, cursor + byteLength);
  const value = highByte
    ? new TextDecoder("utf-16le").decode(raw)
    : new TextDecoder(encoding).decode(raw);
  cursor += byteLength + richRunCount * 4;
  if (hasExtension) {
    if (cursor + 4 > body.byteLength)
      throw new LegacyExcelError("Invalid legacy XLS file: truncated rich string");
    cursor += view.getUint32(cursor, true);
  }
  return { value, offset: cursor };
}

function readBiff8String(
  body: Uint8Array,
  offset = 0,
  encoding = "windows-1252",
): { value: string; offset: number } {
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  requireBytes(body, offset + 3, "string");
  const characterCount = view.getUint16(offset, true);
  const flags = body[offset + 2]!;
  return decodeBiffString(body, view, offset + 3, characterCount, flags, encoding);
}

function readBiff5String(
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

function readBiff35String(
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

function readBiff2String(
  body: Uint8Array,
  offset = 0,
  encoding = "windows-1252",
): { value: string; offset: number } {
  return readBiff5String(body, offset, encoding);
}

function readShortBiff8String(
  body: Uint8Array,
  offset = 0,
  encoding = "windows-1252",
): { value: string; offset: number } {
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  requireBytes(body, offset + 2, "string");
  return decodeBiffString(body, view, offset + 2, body[offset]!, body[offset + 1]!, encoding);
}

function encodingForCodepage(codepage: number): string {
  if (codepage === 1200) return "utf-16le";
  if (codepage === 65001) return "utf-8";
  try {
    return new TextDecoder(codepageLabel(codepage), { fatal: false }).encoding;
  } catch {
    return "windows-1252";
  }
}

function codepageLabel(codepage: number): string {
  const labels: Record<number, string> = {
    932: "windows-932",
    936: "windows-936",
    949: "windows-949",
    950: "windows-950",
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

function detectBiffVersion(stream: Uint8Array): {
  version: BiffVersion;
  workbookLayout: "standard" | "biff4w";
} {
  if (stream.byteLength < 6)
    throw new LegacyExcelError("Invalid legacy XLS file: workbook stream is truncated");
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  const code = view.getUint16(0, true);
  if (code === 0)
    throw new LegacyExcelError("Invalid legacy XLS file: workbook stream is truncated");
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
    throw new LegacyExcelError(`Unsupported legacy XLS BIFF version: ${recordVersion}`);
  return { version: recordVersion === 0x0600 ? 8 : 5, workbookLayout: "standard" };
}

const BEGIN_OF_FILE_CODES = new Set([0x0009, 0x0209, 0x0409, 0x0809]);
const PLAINTEXT_ENCRYPTION_CODES = new Set([
  RecordCode.FilePass,
  ...BEGIN_OF_FILE_CODES,
  0x00e1,
  0x0193,
  0x0195,
  0x0196,
  0x0138,
]);
/** Password for password-protected legacy XLS containers. */
export interface LegacyParseOptions {
  /** Password used to verify and decrypt legacy Office RC4 containers. */
  password?: string;
}

function readWorkbookGlobals(stream: Uint8Array): WorkbookState {
  const detected = detectBiffVersion(stream);
  if (detected.version <= 4) return readLegacyWorkbookGlobals(stream, detected);

  const workbookVersion = detected.version as 5 | 8;
  const sharedStrings: string[] = [];
  const sheets: BoundSheet[] = [];
  const fonts: FontOptions[] = [];
  const numberFormats = new Map<number, string>();
  const palette: string[] = [];
  const externalLinks: ExternalLinkOptions[] = [];
  const definedNames: DefinedNameOptions[] = [];
  const externSheetNames = new Map<number, string>();
  const styleTable = new Map<number, { fontIndex?: number; numberFormatId?: number }>();
  const drawingImages: { data: Uint8Array; type: "png" | "jpg" | "wmf" | "emf" }[] = [];
  const pendingSupbooks: { target: string; sheetNames: string[] }[] = [];
  let externSheetReferenceCount = 0;
  let codepage = 1252;
  let date1904 = false;
  let sharedStringParts: Uint8Array[] | undefined;
  let inGlobals = false;

  for (const record of recordsFrom(stream, 0, workbookVersion)) {
    if (record.code === RecordCode.FilePass && !isClearedFilePass(record))
      throw new LegacyExcelError("Encrypted legacy XLS files are not supported");
    if (record.code === RecordCode.Formula || record.code === RecordCode.String) {
      throw new LegacyExcelError(
        "Invalid legacy XLS file: worksheet record appears in workbook globals",
      );
    }
    if (record.code === 0x0809) {
      if (inGlobals) throw new LegacyExcelError("Invalid legacy XLS file: nested workbook globals");
      if (record.body.byteLength < 4)
        throw new LegacyExcelError("Invalid legacy XLS file: truncated Begin Of File");
      // The version token was validated by detectBiffVersion.
      inGlobals = true;
      continue;
    }
    if (!inGlobals) continue;

    if (sharedStringParts && record.code !== RecordCode.Continue) {
      sharedStrings.push(
        ...parseSharedStrings(
          sharedStringParts,
          workbookVersion === 8,
          encodingForCodepage(codepage),
        ),
      );
      sharedStringParts = undefined;
    }
    switch (record.code) {
      case RecordCode.SharedStringTable: {
        if (record.body.byteLength < 8)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated shared string table");
        sharedStringParts = [record.body];
        break;
      }
      case RecordCode.Continue: {
        sharedStringParts?.push(record.body);
        break;
      }
      case RecordCode.BoundSheet: {
        if (record.body.byteLength < 6)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated sheet definition");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const position = view.getUint32(0, true);
        const stateValue = view.getUint16(4, true);
        const state =
          (stateValue & 0x0003) === 1
            ? "hidden"
            : (stateValue & 0x0003) === 2
              ? "veryHidden"
              : "visible";
        if (workbookVersion === 8 && (stateValue >> 8 === 0x02 || stateValue >> 8 === 0x06)) {
          break;
        }
        let name: string;
        try {
          name =
            workbookVersion === 8
              ? readShortBiff8String(record.body, 6).value
              : readBiff5String(record.body, 6, encodingForCodepage(codepage)).value;
        } catch {
          throw new LegacyExcelError("Invalid legacy XLS file: truncated sheet name");
        }
        sheets.push({ position, state, name });
        break;
      }
      case RecordCode.Codepage: {
        if (record.body.byteLength < 2)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated CODEPAGE");
        codepage = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        ).getUint16(0, true);
        break;
      }
      case RecordCode.DateMode: {
        if (record.body.byteLength < 2)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated DATEMODE");
        date1904 =
          new DataView(
            record.body.buffer,
            record.body.byteOffset,
            record.body.byteLength,
          ).getUint16(0, true) === 1;
        break;
      }
      case RecordCode.Font: {
        const font = parseFont(record.body, workbookVersion === 8, encodingForCodepage(codepage));
        if (font) fonts.push(font);
        break;
      }
      case RecordCode.Format: {
        const format = parseFormat(
          record.body,
          workbookVersion === 8,
          encodingForCodepage(codepage),
        );
        if (format) numberFormats.set(format.id, format.code);
        break;
      }
      case RecordCode.Palette: {
        palette.push(...parsePalette(record.body));
        break;
      }
      case RecordCode.Supbook: {
        try {
          const supbook = parseSupbook(record.body, encodingForCodepage(codepage));
          if (supbook) pendingSupbooks.push(supbook);
        } catch {
          // Malformed external-workbook metadata is auxiliary and can be omitted.
        }
        break;
      }
      case RecordCode.ExternSheet: {
        const references = parseExternSheet(record.body);
        references.forEach((reference) => {
          const supbook = pendingSupbooks.at(-1);
          const sheetName = supbook?.sheetNames[reference.firstSheet] ?? "#REF";
          externSheetNames.set(++externSheetReferenceCount, sheetName);
          if (supbook) {
            externalLinks.push({
              externalBook: {
                target: supbook.target,
                sheetNames: [sheetName],
              },
            });
          }
        });
        break;
      }
      case RecordCode.Name: {
        try {
          const name = parseName(
            record.body,
            workbookVersion === 8,
            encodingForCodepage(codepage),
            externSheetNames,
          );
          definedNames.push(name);
        } catch {
          // Tolerate truncated NAME records (corrupt fixtures, BIFF5 edge
          // layouts) — defined names are auxiliary metadata.
        }
        break;
      }
      case RecordCode.Xf: {
        if (record.body.byteLength < 4) continue;
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        styleTable.set(styleTable.size, {
          fontIndex: view.getUint16(0, true),
          numberFormatId: view.getUint16(2, true),
        });
        break;
      }
      case RecordCode.MsoDrawingGroup: {
        drawingImages.push(...escherImages(record.body));
        break;
      }
      default:
        break;
    }
  }

  if (sharedStringParts)
    sharedStrings.push(
      ...parseSharedStrings(
        sharedStringParts,
        workbookVersion === 8,
        encodingForCodepage(codepage),
      ),
    );

  if (workbookVersion === undefined)
    throw new LegacyExcelError("Invalid legacy XLS file: missing Begin Of File");
  if (sheets.length === 0)
    throw new LegacyExcelError("Invalid legacy XLS file: workbook has no sheet definitions");
  return {
    version: workbookVersion,
    workbookLayout: "standard",
    sharedStrings,
    sheets,
    fonts,
    numberFormats,
    palette,
    date1904,
    codepage,
    externalLinks,
    definedNames,
    externSheetNames,
    styleTable,
    drawingImages,
    stream,
  };
}

function readLegacyWorkbookGlobals(
  stream: Uint8Array,
  detected: { version: BiffVersion; workbookLayout: "standard" | "biff4w" },
): WorkbookState {
  const sheets: BoundSheet[] = [];
  let sawBegin = false;
  let cursor = 0;

  while (cursor < stream.byteLength) {
    const record = readRecord(stream, cursor, detected.version);
    const nextCursor = cursor + recordHeaderSize(detected.version) + record.body.byteLength;
    if (record.code === RecordCode.FilePass && !isClearedFilePass(record))
      throw new LegacyExcelError("Encrypted legacy XLS files are not supported");
    if (!BEGIN_OF_FILE_CODES.has(record.code)) {
      cursor = nextCursor;
      continue;
    }
    if (sawBegin) throw new LegacyExcelError("Invalid legacy XLS file: nested workbook globals");
    sawBegin = true;
    if (detected.workbookLayout === "standard") {
      sheets.push({ position: 0, state: "visible", name: "Sheet 1" });
      return {
        ...detected,
        sharedStrings: [],
        sheets,
        fonts: [],
        numberFormats: new Map(),
        palette: [],
        date1904: false,
        codepage: 1252,
        externalLinks: [],
        definedNames: [],
        externSheetNames: new Map(),
        styleTable: new Map(),
        drawingImages: [],
        stream,
      };
    }

    cursor = nextCursor;
    while (cursor < stream.byteLength) {
      const workbookRecord = readRecord(stream, cursor, detected.version);
      cursor += recordHeaderSize(detected.version) + workbookRecord.body.byteLength;
      if (workbookRecord.code === RecordCode.FilePass && !isClearedFilePass(workbookRecord))
        throw new LegacyExcelError("Encrypted legacy XLS files are not supported");
      if (workbookRecord.code !== RecordCode.Sheethdr) continue;
      if (workbookRecord.body.byteLength < 5)
        throw new LegacyExcelError("Invalid legacy XLS file: truncated BIFF4W sheet header");
      const sheetLength = new DataView(
        workbookRecord.body.buffer,
        workbookRecord.body.byteOffset,
        workbookRecord.body.byteLength,
      ).getInt32(0, true);
      const sheetStart = cursor;
      const name = readBiff5String(workbookRecord.body, 4).value;
      sheets.push({ position: sheetStart, state: "visible", name });
      cursor = sheetStart + sheetLength;
    }
  }

  if (!sawBegin) throw new LegacyExcelError("Invalid legacy XLS file: missing Begin Of File");
  if (sheets.length === 0)
    throw new LegacyExcelError("Invalid legacy XLS file: workbook has no sheet definitions");
  return {
    ...detected,
    sharedStrings: [],
    sheets,
    fonts: [],
    numberFormats: new Map(),
    palette: [],
    date1904: false,
    codepage: 1252,
    externalLinks: [],
    definedNames: [],
    externSheetNames: new Map(),
    styleTable: new Map(),
    drawingImages: [],
    stream,
  };
}

function parseFont(body: Uint8Array, isBiff8: boolean, encoding: string): FontOptions | undefined {
  if (body.byteLength < 15) return undefined;
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  // MS-XLS 2.5.146: Font.name is a shortXLUnicodeString — a 1-byte cch (not
  // the 2-byte XLUnicodeString count), preceded by 14 fixed bytes.
  let fontName = "";
  try {
    const characterCount = body[14]!;
    fontName = isBiff8
      ? decodeBiffString(body, view, 16, characterCount, body[15] ?? 0, encoding).value
      : new TextDecoder(encoding).decode(body.subarray(15, 15 + characterCount));
  } catch {
    fontName = "";
  }
  const weight = view.getUint16(6, true);
  return {
    size: view.getUint16(0, true) / 20,
    bold: weight >= 700,
    italic: (view.getUint16(2, true) & 0x0002) !== 0,
    colorIndexed: view.getUint16(4, true),
    underline: body[10] === 1,
    font: fontName,
  };
}

function parseFormat(
  body: Uint8Array,
  isBiff8: boolean,
  encoding: string,
): { id: number; code: string } | undefined {
  if (body.byteLength < 4) return undefined;
  const id = new DataView(body.buffer, body.byteOffset, body.byteLength).getUint16(0, true);
  let code: string;
  try {
    code = (isBiff8 ? readBiff8String(body, 2, encoding) : readBiff5String(body, 2, encoding))
      .value;
  } catch {
    return undefined;
  }
  return {
    id,
    code,
  };
}

function parsePalette(body: Uint8Array): string[] {
  if (body.byteLength < 2) return [];
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const count = view.getUint16(0, true);
  if (body.byteLength < 2 + count * 4) return [];
  return Array.from({ length: count }, (_, index) => {
    const value = view.getUint32(2 + index * 4, true);
    const blue = (value & 0xff).toString(16).padStart(2, "0");
    const green = ((value >> 8) & 0xff).toString(16).padStart(2, "0");
    const red = ((value >> 16) & 0xff).toString(16).padStart(2, "0");
    return `FF${red}${green}${blue}`;
  });
}

function parseSupbook(
  body: Uint8Array,
  encoding: string,
): { target: string; sheetNames: string[] } | undefined {
  if (body.byteLength < 2) return undefined;
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const sheetCount = view.getUint16(0, true);
  if (body.byteLength >= 4 && view.getUint16(2, true) === 0x0401) return undefined;
  const target = readBiff8String(body, 2, encoding);
  const sheetNames: string[] = [];
  for (let index = 0; index < sheetCount; index++) {
    sheetNames.push(readShortBiff8String(body, target.offset, encoding).value);
  }
  return { target: target.value, sheetNames };
}

function parseExternSheet(body: Uint8Array): { firstSheet: number; lastSheet: number }[] {
  if (body.byteLength < 2)
    throw new LegacyExcelError("Invalid legacy XLS file: truncated EXTERNSHEET");
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const count = view.getUint16(0, true);
  if (body.byteLength < 2 + count * 6) return [];
  return Array.from({ length: count }, (_, index) => ({
    firstSheet: view.getUint16(4 + index * 6, true),
    lastSheet: view.getUint16(6 + index * 6, true),
  }));
}

function parseName(
  body: Uint8Array,
  isBiff8: boolean,
  encoding: string,
  externSheetNames: Map<number, string>,
): DefinedNameOptions {
  if (body.byteLength < 14) throw new LegacyExcelError("Invalid legacy XLS file: truncated NAME");
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const flags = view.getUint16(0, true);
  const formulaLength = view.getUint16(4, true);
  const localSheetId = view.getUint16(8, true) === 0 ? undefined : view.getUint16(8, true) - 1;
  const name = isBiff8
    ? readShortBiff8String(body, 14, encoding)
    : readBiff5String(body, 14, encoding);
  let value = name.value;
  try {
    const tokens = body.subarray(name.offset, name.offset + formulaLength);
    value = decodeFormula(tokens, {
      baseRow: 0,
      baseColumn: 0,
      sheetNameByExternIndex: (index) => externSheetNames.get(index),
    }).formula;
  } catch {
    value = name.value;
  }
  return {
    name: name.value,
    value,
    localSheetId,
    hidden: (flags & 0x0001) !== 0,
    function: (flags & 0x0002) !== 0,
    vbProcedure: (flags & 0x0004) !== 0,
  };
}

function readCString(body: Uint8Array): string {
  const end = body.indexOf(0);
  return new TextDecoder(encodingForCodepage(1252)).decode(
    body.subarray(0, end === -1 ? body.byteLength : end),
  );
}

function readNulTerminatedUnicode(
  body: Uint8Array,
  offset: number,
): { value: string; offset: number } {
  if (offset + 4 > body.byteLength)
    throw new LegacyExcelError("Invalid legacy XLS file: truncated unicode string");
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const byteLength = view.getUint32(offset, true) * 2;
  offset += 4;
  if (offset + byteLength > body.byteLength)
    throw new LegacyExcelError("Invalid legacy XLS file: truncated unicode string");
  const raw = body.subarray(offset, offset + byteLength - 2);
  return { value: new TextDecoder("utf-16le").decode(raw), offset: offset + byteLength };
}

function readByteCountUnicode(body: Uint8Array, offset: number): { value: string; offset: number } {
  if (offset + 4 > body.byteLength)
    throw new LegacyExcelError("Invalid legacy XLS file: truncated unicode string");
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const byteLength = view.getUint32(offset, true);
  offset += 4;
  if (offset + byteLength > body.byteLength)
    throw new LegacyExcelError("Invalid legacy XLS file: truncated unicode string");
  const raw = body.subarray(offset, offset + byteLength - 2);
  return { value: new TextDecoder("utf-16le").decode(raw), offset: offset + byteLength };
}

function parseHyperlink(body: Uint8Array): HyperlinkOptions {
  if (body.byteLength < 32) throw new LegacyExcelError("Invalid legacy XLS file: truncated HLINK");
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const row = view.getUint16(0, true);
  const column = view.getUint16(4, true);
  const options = view.getInt32(28, true);
  let offset = 32;
  let display: string | undefined;
  if ((options & 0x14) !== 0) {
    const value = readNulTerminatedUnicode(body, offset);
    display = value.value;
    offset = value.offset;
  }
  if ((options & 0x80) !== 0) offset = readNulTerminatedUnicode(body, offset).offset;
  let url: string | undefined;
  let location: string | undefined;
  if ((options & 0x0001) !== 0 && (options & 0x0100) === 0) {
    const moniker = body.subarray(offset, offset + 16);
    offset += 16;
    if (moniker[0] === 0xe0 && moniker[1] === 0xc9) {
      const value = readByteCountUnicode(body, offset);
      url = value.value;
    } else if ((options & 0x0008) !== 0) {
      const value = readNulTerminatedUnicode(body, offset);
      location = value.value;
    }
  } else if ((options & 0x0100) !== 0) {
    const value = readNulTerminatedUnicode(body, offset);
    location = value.value;
  }
  return {
    cell: reference(row, column),
    url,
    location,
    tooltip: display,
    display,
  };
}

function parseSharedStrings(
  parts: readonly Uint8Array[],
  isBiff8: boolean,
  encoding = "windows-1252",
): string[] {
  const first = parts[0]!;
  const firstView = new DataView(first.buffer, first.byteOffset, first.byteLength);
  const stringCount = firstView.getUint32(4, true);
  const strings: string[] = [];
  let partIndex = 0;
  let offset = 8;
  let highByte = false;

  const takeByte = (): number => {
    ensureByte();
    return parts[partIndex]![offset++]!;
  };
  const ensureByte = (): void => {
    const part = parts[partIndex];
    if (part && offset < part.byteLength) return;
    partIndex++;
    offset = 0;
    if (!parts[partIndex])
      throw new LegacyExcelError("Invalid legacy XLS file: truncated shared string table");
  };
  const beginCharacter = (): void => {
    ensureByte();
    if (isBiff8 && partIndex > 0 && offset === 0) highByte = (takeByte() & 0x01) !== 0;
  };
  const takeUint16 = (): number => takeByte() | (takeByte() << 8);
  const takeUint32 = (): number => takeUint16() | (takeUint16() << 16);
  const skipBytes = (length: number): void => {
    for (let index = 0; index < length; index++) takeByte();
  };

  for (let index = 0; index < stringCount; index++) {
    const currentPart = parts[partIndex];
    if (!currentPart || (offset >= currentPart.byteLength && partIndex + 1 >= parts.length)) {
      break;
    }
    const characterCount = takeUint16();
    const flags = takeByte();
    highByte = isBiff8 && (flags & 0x01) !== 0;
    const richRunCount = (flags & 0x08) !== 0 ? takeUint16() : 0;
    const extensionLength = (flags & 0x04) !== 0 ? takeUint32() : 0;
    const characterBytes: number[] = [];
    for (let characterIndex = 0; characterIndex < characterCount; characterIndex++) {
      beginCharacter();
      const firstByte = takeByte();
      const secondByte = highByte ? takeByte() : 0;
      characterBytes.push(firstByte);
      if (highByte) characterBytes.push(secondByte);
    }
    skipBytes(richRunCount * 4 + extensionLength);
    strings.push(
      new TextDecoder(highByte ? "utf-16le" : encoding).decode(new Uint8Array(characterBytes)),
    );
  }
  return strings;
}

function decodeRk(value: number): number {
  const divideBy100 = (value & 0x01) !== 0;
  const isInteger = (value & 0x02) !== 0;
  let result: number;
  if (isInteger) {
    result = value >> 2;
  } else {
    const bytes = new Uint8Array(8);
    new DataView(bytes.buffer).setUint32(4, value & 0xfffffffc, true);
    result = new DataView(bytes.buffer).getFloat64(0, true);
  }
  return divideBy100 ? result / 100 : result;
}

function columnLabel(column: number): string {
  let value = column;
  let label = "";
  do {
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26) - 1;
  } while (value >= 0);
  return label;
}

function concatBytes(parts: readonly Uint8Array[]): Uint8Array {
  const total = parts.reduce((length, part) => length + part.byteLength, 0);
  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const part of parts) {
    bytes.set(part, offset);
    offset += part.byteLength;
  }
  return bytes;
}

function decodeParts(parts: readonly TxoTextPart[]): string {
  return parts.map(({ encoding, data }) => new TextDecoder(encoding).decode(data)).join("");
}

function reference(row: number, column: number): string {
  return `${columnLabel(column)}${row + 1}`;
}

function ensureRow(sheet: SheetState, rowNumber: number): RowOptions {
  const existing = sheet.rows.get(rowNumber);
  if (existing) return existing;
  const row: RowOptions = { rowNumber: rowNumber + 1, cells: [] };
  sheet.rows.set(rowNumber, row);
  return row;
}

function cellStyle(style: number, state: WorkbookState): number | StyleOptions {
  const xf = state.styleTable.get(style);
  if (!xf) return style;
  const font = state.fonts[xf.fontIndex ?? -1];
  const format =
    xf.numberFormatId === undefined ? undefined : state.numberFormats.get(xf.numberFormatId);
  return { font, numFmt: format };
}

function makeCell(
  row: number,
  column: number,
  style: number,
  value: CellOptions["value"],
  state?: WorkbookState,
): CellOptions {
  return {
    reference: reference(row, column),
    value,
    style: state ? cellStyle(style, state) : style,
  };
}

function formulaOptions(
  expression?: string,
  overrides?: Partial<FormulaOptions>,
): FormulaOptions | undefined {
  return expression ? { formula: expression, ...overrides } : undefined;
}

function errorLiteral(code: number): string {
  const literals: { [code: number]: string } = {
    0x00: "#NULL!",
    0x07: "#DIV/0!",
    0x0f: "#VALUE!",
    0x17: "#REF!",
    0x1d: "#NAME?",
    0x24: "#NUM!",
    0x2a: "#N/A",
    0x2b: "#GETTING_DATA",
  };
  const literal = literals[code];
  if (!literal)
    throw new LegacyExcelError(
      `Invalid legacy XLS file: unknown formula error code 0x${code.toString(16)}`,
    );
  return literal;
}

function isStringResult(value: Uint8Array): boolean {
  return (
    value.byteLength === 8 &&
    value.subarray(0, 6).every((byte) => byte === 0) &&
    value[6] === 0xff &&
    value[7] === 0xff
  );
}

function cachedFormulaResult(value: Uint8Array): CellOptions["value"] | undefined {
  if (isStringResult(value)) return undefined;
  const view = new DataView(value.buffer, value.byteOffset, value.byteLength);
  if (!value.subarray(0, 6).every((byte) => byte === 0)) return view.getFloat64(0, true);
  if (value[6] === 0x00) return view.getFloat64(0, true);
  if (value[6] === 0x01) return value[7] !== 0;
  if (value[6] === 0x02) return errorLiteral(value[7]!);
  if (value[6] === 0x03) return "";
  if (value[6] === 0xff) return undefined;
  return view.getFloat64(0, true);
}

function parseWorksheetStream(
  stream: Uint8Array,
  position: number,
  state: WorkbookState,
): SheetState {
  const sheet: SheetState = { rows: new Map(), extras: {} };
  let formula: FormulaLocation | undefined;
  let depth = 0;
  let pendingFormula: FormulaLocation | undefined;
  let currentDrawing: Uint8Array[] | undefined;
  let txo: TxoState | undefined;
  let conditional: PendingConditional | undefined;
  const txoTextByObjectId = new Map<number, TxoText>();
  let noteObjectId: number | undefined;
  let sharedFormulaIndex = 0;
  const worksheetExtras: Partial<WorksheetOptions> = {};

  for (const record of recordsFrom(stream, position, state.version, false)) {
    if (BEGIN_OF_FILE_CODES.has(record.code)) {
      depth++;
      continue;
    }
    if (depth === 0) continue;
    if (record.code === RecordCode.EndOfFile) {
      depth--;
      if (depth === 0) return { ...sheet, extras: worksheetExtras };
      continue;
    }
    if (depth !== 1) continue;

    if (record.code === RecordCode.FilePass && !isClearedFilePass(record))
      throw new LegacyExcelError("Encrypted legacy XLS files are not supported");
    if (
      currentDrawing &&
      record.code !== RecordCode.Continue &&
      record.code !== RecordCode.MsoDrawing
    ) {
      worksheetExtras.images = escherPictures(concatBytes(currentDrawing), state.drawingImages);
      currentDrawing = undefined;
    }
    if (currentDrawing && record.code === RecordCode.Continue) {
      currentDrawing.push(record.body);
      continue;
    }
    if (txo && record.code === RecordCode.Continue) {
      if (txo.phase === "text") {
        const highByte = record.body[0] === 1;
        const data = record.body.subarray(1, 1 + txo.characterCount * (highByte ? 2 : 1));
        const characterCount = highByte ? Math.floor(data.byteLength / 2) : data.byteLength;
        txo.textParts.push({ encoding: highByte ? "utf-16le" : "windows-1252", data });
        txo.characterCount -= characterCount;
        if (txo.characterCount <= 0) txo.phase = "runs";
      } else {
        txo.runByteCount -= record.body.byteLength;
      }
      if (txo.runByteCount <= 0) {
        txoTextByObjectId.set(txo.objectId, {
          text: decodeParts(txo.textParts),
        });
        txo = undefined;
      }
      continue;
    }
    if (record.code === RecordCode.String && pendingFormula) {
      let text: string;
      try {
        text =
          state.version === 8
            ? readBiff8String(record.body).value
            : state.version === 2
              ? readBiff2String(record.body).value
              : readBiff35String(record.body).value;
      } catch {
        throw new LegacyExcelError("Invalid legacy XLS file: truncated formula string");
      }
      ensureRow(sheet, pendingFormula.row).cells!.push({
        ...makeCell(
          pendingFormula.row,
          pendingFormula.column,
          pendingFormula.style,
          undefined,
          state,
        ),
        value: text,
        formula: formulaOptions(pendingFormula.expression),
      });
      pendingFormula = undefined;

      continue;
    }
    if (record.code !== RecordCode.SharedFormula && record.code !== RecordCode.ArrayFormula) {
      pendingFormula = undefined;
    }

    switch (record.code) {
      case RecordCode.Dimensions: {
        const minimumLength = state.version === 8 ? 14 : state.version === 2 ? 8 : 10;
        if (record.body.byteLength < minimumLength) {
          throw new LegacyExcelError("Invalid legacy XLS file: truncated dimensions");
        }
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const firstRow = state.version === 8 ? view.getUint32(0, true) : view.getUint16(0, true);
        const lastRow = state.version === 8 ? view.getUint32(4, true) : view.getUint16(2, true);
        const firstColumn =
          state.version === 8
            ? view.getUint16(8, true)
            : state.version === 2
              ? record.body[4]!
              : view.getUint16(4, true);
        const lastColumn =
          state.version === 8
            ? view.getUint16(10, true)
            : state.version === 2
              ? record.body[5]!
              : view.getUint16(6, true);
        sheet.dimension = `${reference(firstRow, firstColumn)}:${reference(lastRow, lastColumn)}`;
        break;
      }
      case RecordCode.Row: {
        if (record.body.byteLength < 16)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated row");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = ensureRow(sheet, view.getUint16(0, true));
        row.spans = `${view.getUint16(2, true) + 1}:${view.getUint16(4, true)}`;
        row.height = view.getUint16(6, true) / 20;
        const flags = view.getUint16(12, true);
        row.outlineLevel = flags & 0x0007;
        row.collapsed = (flags & 0x0010) !== 0;
        row.hidden = (flags & 0x0020) !== 0;
        row.customFormat = (flags & 0x0080) !== 0;
        row.thickTop = (flags & 0x0100) !== 0;
        row.thickBot = (flags & 0x0200) !== 0;
        row.phonetic = (flags & 0x0400) !== 0;
        if (row.customFormat) row.style = view.getUint16(14, true) & 0x0fff;
        break;
      }
      case RecordCode.Blank:
      case RecordCode.Biff2Blank: {
        if (record.body.byteLength < 6)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated blank cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        ensureRow(sheet, row).cells!.push({
          reference: reference(row, column),
          value: null,
          style: view.getUint16(4, true),
        });
        break;
      }
      case RecordCode.BoolErr:
      case RecordCode.Biff2BoolErr: {
        if (record.body.byteLength < 8)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated boolean/error cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        const style = view.getUint16(4, true);
        const value = record.body[6]!;
        const isError = record.body[7]! !== 0;
        ensureRow(sheet, row).cells!.push(
          makeCell(row, column, style, isError ? errorLiteral(value) : value !== 0, state),
        );
        break;
      }
      case RecordCode.MulBlank: {
        if (record.body.byteLength < 10)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated multiple blank cells");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const firstColumn = view.getUint16(2, true);
        const lastColumn = view.getUint16(record.body.byteLength - 2, true);
        const count = lastColumn - firstColumn + 1;
        if (record.body.byteLength < 6 + count * 2) {
          throw new LegacyExcelError(
            "Invalid legacy XLS file: multiple blank cell range is truncated",
          );
        }
        const rowOptions = ensureRow(sheet, row);
        for (let index = 0; index < count; index++) {
          const column = firstColumn + index;
          const style = view.getUint16(4 + index * 2, true);
          rowOptions.cells!.push({ reference: reference(row, column), value: null, style });
        }
        break;
      }
      case RecordCode.Biff2Number:
      case RecordCode.Number: {
        if (state.version === 2) {
          if (record.body.byteLength < 15)
            throw new LegacyExcelError("Invalid legacy XLS file: truncated BIFF2 number cell");
          const view = new DataView(
            record.body.buffer,
            record.body.byteOffset,
            record.body.byteLength,
          );
          ensureRow(sheet, view.getUint16(0, true)).cells!.push(
            makeCell(view.getUint16(0, true), view.getUint16(2, true), 0, view.getFloat64(7, true)),
          );
          break;
        }
        if (record.body.byteLength < 14)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated number cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        ensureRow(sheet, row).cells!.push(
          makeCell(row, column, view.getUint16(4, true), view.getFloat64(6, true), state),
        );
        break;
      }
      case RecordCode.Biff2Integer: {
        if (record.body.byteLength < 9)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated BIFF2 integer cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        ensureRow(sheet, view.getUint16(0, true)).cells!.push(
          makeCell(
            view.getUint16(0, true),
            view.getUint16(2, true),
            0,
            view.getUint16(7, true),
            state,
          ),
        );
        break;
      }
      case RecordCode.Label: {
        if (record.body.byteLength < 7)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated label cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        const style = view.getUint16(4, true);
        let text: { value: string; offset: number };
        try {
          text =
            state.version === 8
              ? readBiff8String(record.body, 6)
              : readBiff35String(record.body, 6);
        } catch {
          throw new LegacyExcelError("Invalid legacy XLS file: truncated label cell");
        }
        ensureRow(sheet, row).cells!.push(makeCell(row, column, style, text.value, state));
        break;
      }
      case RecordCode.Biff2Label: {
        if (record.body.byteLength < 8)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated BIFF2 label cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        let text: { value: string; offset: number };
        try {
          text = readBiff2String(record.body, 7);
        } catch {
          throw new LegacyExcelError("Invalid legacy XLS file: truncated BIFF2 label cell");
        }
        ensureRow(sheet, view.getUint16(0, true)).cells!.push(
          makeCell(view.getUint16(0, true), view.getUint16(2, true), 0, text.value, state),
        );
        break;
      }
      case RecordCode.RichLabel: {
        if (record.body.byteLength < 7)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated rich label cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        const style = view.getUint16(4, true);
        let text: { value: string; offset: number };
        try {
          text = readBiff8String(record.body, 6);
        } catch {
          throw new LegacyExcelError("Invalid legacy XLS file: truncated rich label cell");
        }
        ensureRow(sheet, row).cells!.push(makeCell(row, column, style, text.value, state));
        break;
      }
      case RecordCode.LabelSst: {
        if (record.body.byteLength < 10)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated shared string cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        const style = view.getUint16(4, true);
        const index = view.getUint32(6, true);
        const value = state.sharedStrings[index];
        if (value === undefined)
          throw new LegacyExcelError(`Invalid legacy XLS file: missing shared string ${index}`);
        ensureRow(sheet, row).cells!.push(makeCell(row, column, style, value, state));
        break;
      }
      case RecordCode.Rk: {
        if (record.body.byteLength < 10)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated RK cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        ensureRow(sheet, row).cells!.push(
          makeCell(row, column, view.getUint16(4, true), decodeRk(view.getInt32(6, true)), state),
        );
        break;
      }
      case RecordCode.MulRk: {
        if (record.body.byteLength < 10)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated multiple RK cells");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const firstColumn = view.getUint16(2, true);
        const lastColumn = view.getUint16(record.body.byteLength - 2, true);
        const count = lastColumn - firstColumn + 1;
        if (record.body.byteLength < 6 + count * 6) {
          throw new LegacyExcelError(
            "Invalid legacy XLS file: multiple RK cell range is truncated",
          );
        }
        const rowOptions = ensureRow(sheet, row);
        for (let index = 0; index < count; index++) {
          const offset = 4 + index * 6;
          const column = firstColumn + index;
          const style = view.getUint16(offset, true);
          const value = decodeRk(view.getInt32(offset + 2, true));
          rowOptions.cells!.push(makeCell(row, column, style, value, state));
        }
        break;
      }
      case RecordCode.Formula: {
        if (record.body.byteLength < 22)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated formula cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        const cached = record.body.subarray(6, 14);
        const result = cachedFormulaResult(cached);
        const tokenLength = view.getUint16(20, true);
        let expression: string | undefined;
        if (tokenLength > 0) {
          const tokens = record.body.subarray(22, 22 + tokenLength);
          try {
            expression = decodeFormula(tokens, {
              baseRow: row,
              baseColumn: column,
              sheetNameByExternIndex: (index) => state.externSheetNames.get(index),
            }).formula;
          } catch {
            expression = undefined;
          }
        }
        if (result === undefined) {
          formula = { row, column, style: view.getUint16(4, true), expression };
          pendingFormula = formula;

          continue;
        }
        ensureRow(sheet, row).cells!.push({
          ...makeCell(row, column, view.getUint16(4, true), result, state),
          formula: formulaOptions(expression),
        });
        break;
      }
      case RecordCode.SharedFormula: {
        if (record.body.byteLength < 10 || !formula) break;
        const tokens = record.body.subarray(10);
        try {
          const decoded = decodeFormula(tokens, {
            baseRow: formula.row,
            baseColumn: formula.column,
            sheetNameByExternIndex: (index) => state.externSheetNames.get(index),
          });
          ensureRow(sheet, formula.row).cells!.push({
            ...makeCell(
              formula.row,
              formula.column,
              formula.style,
              cachedFormulaResult(new Uint8Array(8)),
              state,
            ),
            formula: {
              formula: decoded.formula,
              type: "shared",
              sharedIndex: sharedFormulaIndex++,
            },
          });
        } catch {
          break;
        }
        break;
      }
      case RecordCode.ArrayFormula: {
        if (record.body.byteLength < 14 || !formula) break;
        const firstRow = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        ).getUint16(0, true);
        const firstColumn = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        ).getUint16(4, true);
        const lastRow = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        ).getUint16(2, true);
        const lastColumn = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        ).getUint16(6, true);
        const referenceText = `${reference(firstRow, firstColumn)}:${reference(lastRow, lastColumn)}`;
        try {
          const decoded = decodeFormula(record.body.subarray(14), {
            baseRow: formula.row,
            baseColumn: formula.column,
            sheetNameByExternIndex: (index) => state.externSheetNames.get(index),
          });
          ensureRow(sheet, formula.row).cells!.push({
            ...makeCell(
              formula.row,
              formula.column,
              formula.style,
              cachedFormulaResult(new Uint8Array(8)),
              state,
            ),
            formula: { formula: decoded.formula, type: "array", reference: referenceText },
          });
        } catch {
          break;
        }
        break;
      }
      case RecordCode.ColInfo: {
        if (record.body.byteLength < 10)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated COLINFO");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const flags = view.getUint16(8, true);
        worksheetExtras.columns = [
          ...(worksheetExtras.columns ?? []),
          {
            min: view.getUint16(0, true) + 1,
            max: view.getUint16(2, true) + 1,
            width: view.getUint16(4, true) / 256,
            customWidth: true,
            hidden: (flags & 0x0001) !== 0,
            outlineLevel: (flags >> 8) & 0x0007,
            collapsed: (flags & 0x1000) !== 0,
          },
        ];
        break;
      }
      case RecordCode.MergedCells: {
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const available =
          record.body.byteLength < 2 ? 0 : Math.floor((record.body.byteLength - 2) / 8);
        const count = record.body.byteLength < 2 ? 0 : Math.min(view.getUint16(0, true), available);
        worksheetExtras.mergeCells = Array.from({ length: count }, (_, index) => {
          const offset = 2 + index * 8;
          const firstRow = view.getUint16(offset, true);
          const lastRow = view.getUint16(offset + 2, true);
          const firstColumn = view.getUint16(offset + 4, true);
          const lastColumn = view.getUint16(offset + 6, true);
          return { ref: `${reference(firstRow, firstColumn)}:${reference(lastRow, lastColumn)}` };
        });
        break;
      }
      case RecordCode.Hyperlink: {
        try {
          worksheetExtras.hyperlinks = [
            ...(worksheetExtras.hyperlinks ?? []),
            parseHyperlink(record.body),
          ];
        } catch {
          // Malformed hyperlink metadata is auxiliary and can be omitted.
        }
        break;
      }
      case RecordCode.Obj: {
        if (record.body.byteLength >= 8) {
          const view = new DataView(
            record.body.buffer,
            record.body.byteOffset,
            record.body.byteLength,
          );
          const objectId = view.getUint16(6, true);
          if (view.getUint16(4, true) === 0x0019) {
            noteObjectId = objectId;
          }
        }
        break;
      }
      case RecordCode.Txo: {
        if (record.body.byteLength < 18) {
          txo = undefined;
          break;
        }
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const objectId = noteObjectId ?? 0;
        txo = {
          objectId,
          characterCount: view.getUint16(10, true),
          runByteCount: view.getUint16(12, true),
          textParts: [],
          phase: view.getUint16(10, true) === 0 ? "runs" : "text",
        };
        break;
      }
      case RecordCode.Note: {
        if (record.body.byteLength < 8)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated NOTE");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        const objectId = view.getUint16(6, true);
        const author = readShortBiff8String(
          record.body,
          8,
          encodingForCodepage(state.codepage),
        ).value;
        const noteText = txoTextByObjectId.get(objectId)?.text ?? "";
        worksheetExtras.comments = [
          ...(worksheetExtras.comments ?? []),
          {
            cell: reference(row, column),
            author,
            text: noteText,
            visible: (view.getUint16(4, true) & 0x0002) !== 0,
            size: { width: 108, height: 59.25 },
            anchor: { from: { col: column, row }, to: { col: column + 1, row: row + 1 } },
          },
        ];
        break;
      }
      case RecordCode.MsoDrawing:
      case RecordCode.MsoDrawingGroup: {
        currentDrawing = currentDrawing ? [...currentDrawing, record.body] : [record.body];
        break;
      }
      case RecordCode.CondFmt: {
        if (record.body.byteLength < 14)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated CONDFMT");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const rangeCount = view.getUint16(12, true);
        if (record.body.byteLength < 14 + rangeCount * 8)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated CONDFMT ranges");
        const range = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        conditional = {
          reference: `${reference(range.getUint16(4, true), range.getUint16(8, true))}:${reference(
            range.getUint16(6, true),
            range.getUint16(10, true),
          )}`,
          remaining: view.getUint16(0, true),
          rules: [],
        };
        break;
      }
      case RecordCode.ConditionalFormat: {
        if (!conditional || record.body.byteLength < 12)
          throw new LegacyExcelError(
            "Invalid legacy XLS file: conditional format rule outside CONDFMT",
          );
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const type = record.body[0]!;
        const operatorCode = record.body[1]!;
        const firstLength = view.getUint16(2, true);
        const secondLength = view.getUint16(4, true);
        const flags = view.getUint32(6, true);
        let offset = 12;
        if ((flags & 0x04000000) !== 0) offset += 118;
        if ((flags & 0x10000000) !== 0) offset += 8;
        if ((flags & 0x20000000) !== 0) offset += 4;
        const decodeRuleFormula = (length: number): string | undefined => {
          if (length === 0) return undefined;
          const bytes = record.body.subarray(offset, offset + length);
          offset += length;
          try {
            return decodeFormula(bytes, { baseRow: 0, baseColumn: 0 }).formula;
          } catch {
            return `#FORMULA_${length}`;
          }
        };
        const formulas = [decodeRuleFormula(firstLength), decodeRuleFormula(secondLength)].filter(
          (formula): formula is string => formula !== undefined,
        );
        conditional.rules.push({
          type: type === 1 ? "expression" : "cellIs",
          operator:
            operatorCode === 3
              ? "equal"
              : operatorCode === 4
                ? "notEqual"
                : operatorCode === 5
                  ? "greaterThan"
                  : operatorCode === 6
                    ? "lessThan"
                    : operatorCode === 7
                      ? "greaterThanOrEqual"
                      : operatorCode === 8
                        ? "lessThanOrEqual"
                        : "between",
          formulas,
          priority: conditional.rules.length + 1,
        });
        conditional.remaining--;
        if (conditional.remaining === 0) {
          worksheetExtras.conditionalFormats = [
            ...(worksheetExtras.conditionalFormats ?? []),
            { sqref: conditional.reference, rules: conditional.rules },
          ];
          conditional = undefined;
        }
        break;
      }
      case RecordCode.Header: {
        worksheetExtras.headerFooter = {
          ...worksheetExtras.headerFooter,
          oddHeader: readCString(record.body),
        };
        break;
      }
      case RecordCode.Footer: {
        worksheetExtras.headerFooter = {
          ...worksheetExtras.headerFooter,
          oddFooter: readCString(record.body),
        };
        break;
      }
      case RecordCode.LeftMargin:
      case RecordCode.RightMargin:
      case RecordCode.TopMargin:
      case RecordCode.BottomMargin: {
        if (record.body.byteLength < 8)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated page margin");
        const value = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        ).getFloat64(0, true);
        worksheetExtras.pageMargins = {
          ...worksheetExtras.pageMargins,
          ...(record.code === RecordCode.LeftMargin
            ? { left: value }
            : record.code === RecordCode.RightMargin
              ? { right: value }
              : record.code === RecordCode.TopMargin
                ? { top: value }
                : { bottom: value }),
        };
        break;
      }
      case RecordCode.PageSetup: {
        if (record.body.byteLength < 32)
          throw new LegacyExcelError("Invalid legacy XLS file: truncated PAGESETUP");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const flags = view.getUint16(10, true);
        worksheetExtras.pageSetup = {
          paperSize: view.getUint16(0, true),
          scale: view.getUint16(2, true),
          fitToWidth: view.getUint16(6, true),
          fitToHeight: view.getUint16(8, true),
          pageOrder: (flags & 1) === 0 ? "downThenOver" : "overThenDown",
          orientation: (flags & 2) === 0 ? "landscape" : "portrait",
        };
        worksheetExtras.pageMargins = {
          ...worksheetExtras.pageMargins,
          header: view.getFloat64(16, true),
          footer: view.getFloat64(24, true),
        };
        break;
      }
      default:
        break;
    }
  }

  if (depth !== 0)
    throw new LegacyExcelError("Invalid legacy XLS file: sheet stream has no Begin Of File");
  return { ...sheet, extras: worksheetExtras };
}

function readWorkbookStream(data: Uint8Array): WorkbookState {
  if (data.byteLength > 8 && data[0] === 0x09) {
    const globals = readWorkbookGlobals(data);
    return { ...globals, stream: data };
  }
  const reader = new CompoundFileReader(data);
  const stream = readRequiredStream(reader, reader.entry("Workbook") ? "Workbook" : "Book");
  if (stream.byteLength < 8)
    throw new LegacyExcelError("Invalid legacy XLS file: workbook stream is truncated");
  const globals = readWorkbookGlobals(stream);
  return { ...globals, stream };
}

export function parseWorkbook(data: Uint8Array, options?: LegacyParseOptions): WorkbookOptions {
  if (!(data instanceof Uint8Array)) throw new LegacyExcelError("XLS data must be a Uint8Array");
  let workbookData: Uint8Array;
  try {
    workbookData = decryptWorkbookContainer(data, options?.password) ?? data;
  } catch (error) {
    throw toLegacyExcelError(error);
  }
  const state = (() => {
    try {
      return readWorkbookStream(workbookData);
    } catch (error) {
      throw toLegacyExcelError(error);
    }
  })();
  const metadata = readSummaryInformation(data);
  const worksheets: WorksheetOptions[] = state.sheets.map((sheet, index) => {
    const parsed = parseWorksheetStream(state.stream, sheet.position, state);
    return {
      name: sheet.name,
      sheetId: index + 1,
      state: sheet.state,
      rows: [...parsed.rows.values()].sort(
        (left, right) => (left.rowNumber ?? 0) - (right.rowNumber ?? 0),
      ),
      dimension: parsed.dimension,
      ...parsed.extras,
    };
  });
  return {
    worksheets,
    ...metadata,
    fonts: state.fonts.length > 0 ? state.fonts : undefined,
    numFmts:
      state.numberFormats.size > 0
        ? [...state.numberFormats].map(([numFmtId, formatCode]) => ({ numFmtId, formatCode }))
        : undefined,
    colors:
      state.palette.length > 0
        ? { indexedColors: state.palette.map((rgb) => ({ rgb: rgb as `FF${string}` })) }
        : undefined,
    definedNames: state.definedNames.length > 0 ? state.definedNames : undefined,
    externalLinks: state.externalLinks.length > 0 ? state.externalLinks : undefined,
    properties: { date1904: state.date1904 },
  };
}

function toLegacyExcelError(error: unknown): LegacyExcelError {
  if (error instanceof LegacyExcelError) return error;
  return new LegacyExcelError(error instanceof Error ? error.message : String(error));
}

function decryptWorkbookContainer(data: Uint8Array, password?: string): Uint8Array | undefined {
  let stream: Uint8Array;
  if (data.byteLength > 8 && data[0] === 0x09) {
    stream = data;
  } else {
    const reader = new CompoundFileReader(data);
    const entry = reader.entry("Workbook") ?? reader.entry("Book");
    if (!entry || entry.type !== "stream") {
      throw new LegacyExcelError("Invalid legacy XLS file: missing Workbook or Book stream");
    }
    stream = reader.read(entry.path);
  }
  const detected = detectBiffVersion(stream);
  let filePass: { offset: number; body: Uint8Array } | undefined;
  let cursor = 0;
  while (cursor < stream.byteLength) {
    const record = readRecord(stream, cursor, detected.version);
    if (record.code === RecordCode.FilePass) {
      filePass = { offset: cursor, body: record.body };
      break;
    }
    if (record.code === RecordCode.EndOfFile) break;
    cursor += recordHeaderSize(detected.version) + record.body.byteLength;
  }
  if (!filePass) return undefined;
  if (password === undefined) {
    throw new LegacyExcelError("Encrypted legacy XLS files are not supported");
  }
  const body = filePass.body;
  if (body.byteLength < 2 || body[0] !== 1 || body[1] !== 0) {
    throw new LegacyExcelError("Encrypted legacy XLS files are not supported");
  }
  const version = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const decrypt = (() => {
    const majorVersion = body.byteLength >= 4 ? version.getUint16(2, true) : 0;
    const minorVersion = body.byteLength >= 6 ? version.getUint16(4, true) : 0;
    if (majorVersion === 1 && minorVersion !== 1) {
      throw new LegacyExcelError("Encrypted legacy XLS files are not supported");
    }
    if (majorVersion !== 1 && (minorVersion !== 0x0002 || majorVersion < 2 || majorVersion > 4)) {
      throw new LegacyExcelError("Encrypted legacy XLS files are not supported");
    }
    if (majorVersion === 1) {
      const verifier = parseLegacyRc4Verifier(body, 6);
      if (!verifyLegacyRc4Password(password, verifier)) {
        throw new LegacyExcelError("Invalid legacy XLS password");
      }
      return (data: Uint8Array): Uint8Array =>
        decryptLegacyRc4(data, password, verifier.salt, 1024);
    }
    const { keySizeBits, verifier } = parseRc4CryptoApiHeader(body, 6);
    if (!verifyRc4CryptoApiPassword(password, verifier, keySizeBits)) {
      throw new LegacyExcelError("Invalid legacy XLS password");
    }
    return (data: Uint8Array): Uint8Array =>
      decryptRc4CryptoApi(data, password, verifier.salt, keySizeBits, 1024);
  })();

  const records: BiffRecord[] = [];
  cursor = 0;
  while (cursor < stream.byteLength) {
    const record = readRecord(stream, cursor, detected.version);
    records.push(record);
    cursor += recordHeaderSize(detected.version) + record.body.byteLength;
  }
  const headerSize = recordHeaderSize(detected.version);
  const encryptedParts: Uint8Array[] = [];
  for (const record of records) {
    const part = new Uint8Array(headerSize + record.body.byteLength);
    if (record.code !== RecordCode.FilePass && !PLAINTEXT_ENCRYPTION_CODES.has(record.code)) {
      if (record.code === RecordCode.BoundSheet) {
        part.set(record.body.subarray(4), headerSize + 4);
      } else {
        part.set(record.body, headerSize);
      }
    }
    encryptedParts.push(part);
  }
  const decryptedBodies = decrypt(concatBytes(encryptedParts));
  const result = new Uint8Array(
    records.reduce(
      (total, record) => total + recordHeaderSize(detected.version) + record.body.byteLength,
      0,
    ),
  );
  let resultOffset = 0;
  let decryptedOffset = 0;
  for (const record of records) {
    result.set(stream.subarray(record.start, record.start + headerSize), resultOffset);
    if (record.code === RecordCode.FilePass) {
      result[resultOffset] = 0;
      result[resultOffset + 1] = 0;
    }
    resultOffset += headerSize;
    const decryptedBodyStart = decryptedOffset + headerSize;
    if (record.code === RecordCode.FilePass) {
      result.fill(0, resultOffset, resultOffset + record.body.byteLength);
    } else if (PLAINTEXT_ENCRYPTION_CODES.has(record.code)) {
      result.set(record.body, resultOffset);
    } else if (record.code === RecordCode.BoundSheet) {
      result.set(record.body.subarray(0, 4), resultOffset);
      result.set(
        decryptedBodies.subarray(
          decryptedBodyStart + 4,
          decryptedBodyStart + record.body.byteLength,
        ),
        resultOffset + 4,
      );
    } else {
      result.set(
        decryptedBodies.subarray(decryptedBodyStart, decryptedBodyStart + record.body.byteLength),
        resultOffset,
      );
    }
    decryptedOffset += headerSize + record.body.byteLength;
    resultOffset += record.body.byteLength;
  }
  return result;
}
