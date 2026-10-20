import type {
  DefinedNameOptions,
  ExternalLinkOptions,
  FontOptions,
  HyperlinkOptions,
} from "@office-open/xlsx";

import { XlsParseError } from "../errors";
import {
  RecordCode,
  isClearedFilePass,
  readRecord,
  recordHeaderSize,
  recordsFrom,
  type BiffRecord,
  type BiffVersion,
} from "../records";
import { assertSupportedBiffRecord, biffRecordContext } from "../records/biff-records";
import { escherImages } from "../records/escher";
import { decodeFormula } from "../records/formula";
import {
  BEGIN_OF_FILE_CODES,
  detectBiffVersion,
  encodingForCodepage,
  readBiff5String,
  readNoCchBiff8String,
  readNoCchBiff5String,
  readShortBiff8String,
} from "./biff";
import { reference } from "./cells";
import type { BoundSheet, WorkbookState, XfStyle } from "./models";
import { parseFont, parseFormat, parsePalette, parseXf } from "./styles";
export function readWorkbookGlobals(stream: Uint8Array): WorkbookState {
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
  const styleTable = new Map<number, XfStyle>();
  const drawingImages: { data: Uint8Array; type: "png" | "jpg" | "wmf" | "emf" }[] = [];
  let saveExternalLinkValues: boolean | undefined;
  let refreshAllConnections: boolean | undefined;
  let builtInGroupCount: number | undefined;
  const pendingSupbooks: Supbook[] = [];
  const pendingExternSheetReferences:
    | ReturnType<typeof parseExternSheet>
    | { supbookIndex: number; firstSheet: number; lastSheet: number }[] = [];
  const pendingNameRecords: BiffRecord[] = [];
  let externSheetReferenceCount = 0;
  let codepage = 1252;
  let date1904 = false;
  let sharedStringParts: Uint8Array[] | undefined;
  let inGlobals = false;

  for (const record of recordsFrom(stream, 0, workbookVersion)) {
    if (record.code === RecordCode.FilePass && !isClearedFilePass(record))
      throw new XlsParseError("Encrypted legacy XLS files are not supported");
    if (record.code === RecordCode.Formula || record.code === RecordCode.String) {
      throw new XlsParseError(
        "Invalid legacy XLS file: worksheet record appears in workbook globals",
      );
    }
    if (record.code === 0x0809) {
      if (inGlobals) throw new XlsParseError("Invalid legacy XLS file: nested workbook globals");
      if (record.body.byteLength < 4)
        throw new XlsParseError("Invalid legacy XLS file: truncated Begin Of File");
      // The version token was validated by detectBiffVersion.
      inGlobals = true;
      continue;
    }
    if (!inGlobals) continue;

    assertSupportedBiffRecord(record);
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
          throw new XlsParseError("Invalid legacy XLS file: truncated shared string table");
        sharedStringParts = [record.body];
        break;
      }
      case RecordCode.Continue: {
        sharedStringParts?.push(record.body);
        break;
      }
      case RecordCode.BoundSheet: {
        if (record.body.byteLength < 6)
          throw new XlsParseError("Invalid legacy XLS file: truncated sheet definition");
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
          throw new XlsParseError("Invalid legacy XLS file: truncated sheet name");
        }
        sheets.push({ position, state, name });
        break;
      }
      case RecordCode.Codepage: {
        if (record.body.byteLength < 2)
          throw new XlsParseError("Invalid legacy XLS file: truncated CODEPAGE");
        codepage = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        ).getUint16(0, true);
        break;
      }
      case RecordCode.DateMode: {
        if (record.body.byteLength < 2)
          throw new XlsParseError("Invalid legacy XLS file: truncated DATEMODE");
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
      case RecordCode.FunctionGroupCount: {
        if (record.body.byteLength < 2)
          throw new XlsParseError("Invalid legacy XLS file: truncated function group count");
        builtInGroupCount = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        ).getUint16(0, true);
        break;
      }
      case RecordCode.RefreshAll: {
        if (record.body.byteLength < 2)
          throw new XlsParseError("Invalid legacy XLS file: truncated refresh-all flag");
        refreshAllConnections =
          (new DataView(
            record.body.buffer,
            record.body.byteOffset,
            record.body.byteLength,
          ).getUint16(0, true) &
            1) ===
          1;
        break;
      }
      case RecordCode.BookBool: {
        if (record.body.byteLength < 2)
          throw new XlsParseError("Invalid legacy XLS file: truncated workbook boolean");
        const value = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        ).getUint16(0, true);
        if (value !== 0 && value !== 1)
          throw new XlsParseError("Invalid legacy XLS file: invalid workbook boolean", {
            ...biffRecordContext(record, "invalid-boolean-value"),
          });
        saveExternalLinkValues = value === 1;
        break;
      }
      case RecordCode.Supbook: {
        if (workbookVersion === 8)
          pendingSupbooks.push(parseSupbook(record, encodingForCodepage(codepage)));
        break;
      }
      case RecordCode.ExternSheet: {
        if (workbookVersion === 8) {
          pendingExternSheetReferences.push(...parseExternSheet(record));
        } else {
          const supbookIndex = pendingSupbooks.length;
          pendingSupbooks.push(parseBiff5ExternSheet(record, encodingForCodepage(codepage)));
          pendingExternSheetReferences.push({ supbookIndex, firstSheet: 0, lastSheet: 0 });
        }
        break;
      }
      case RecordCode.Name: {
        pendingNameRecords.push(record);
        break;
      }
      case RecordCode.Xf: {
        styleTable.set(styleTable.size, parseXf(record.body, workbookVersion, palette));
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
    throw new XlsParseError("Invalid legacy XLS file: missing Begin Of File");
  if (sheets.length === 0)
    throw new XlsParseError("Invalid legacy XLS file: workbook has no sheet definitions");

  for (const reference of pendingExternSheetReferences) {
    const supbook = pendingSupbooks[reference.supbookIndex];
    if (!supbook)
      throw new XlsParseError("Invalid legacy XLS file: EXTERNSHEET references a missing SUPBOOK", {
        path: "workbook-globals/EXTERNSHEET",
        reason: "missing-supbook",
      });
    const sheetName =
      reference.firstSheet === 0xffff ||
      reference.lastSheet === 0xffff ||
      ("sheetNames" in supbook &&
        Array.isArray(supbook.sheetNames) &&
        (reference.firstSheet >= supbook.sheetNames.length ||
          reference.lastSheet >= supbook.sheetNames.length))
        ? "#REF!"
        : workbookVersion === 5
          ? supbook.kind === "current"
            ? ""
            : supbook.sheetNames?.[reference.firstSheet]
          : "sheetNames" in supbook
            ? supbook.sheetNames?.[reference.firstSheet]
            : sheets[reference.firstSheet]?.name;
    if (!sheetName)
      throw new XlsParseError("Invalid legacy XLS file: EXTERNSHEET references a missing sheet", {
        path: "workbook-globals/EXTERNSHEET",
        reason: "missing-external-sheet",
      });
    if (workbookVersion === 5) {
      externSheetReferenceCount += 1;
      externSheetNames.set(externSheetReferenceCount, sheetName);
      externSheetNames.set(-externSheetReferenceCount, sheetName);
    } else {
      externSheetNames.set(++externSheetReferenceCount, sheetName);
    }
    if (supbook.kind === "external") {
      externalLinks.push({ externalBook: { target: supbook.target, sheetNames: [sheetName] } });
    }
    if (supbook.kind === "dde") {
      const [ddeService = supbook.target, ddeTopic = ""] = supbook.target.split("|");
      externalLinks.push({ ddeLink: { ddeService, ddeTopic, ddeItems: [{ name: sheetName }] } });
    }
    if (supbook.kind === "ole") {
      externalLinks.push({ oleLink: { progId: supbook.target, oleItems: [{ name: sheetName }] } });
    }
  }

  for (const record of pendingNameRecords) {
    definedNames.push(
      parseName(record, workbookVersion === 8, encodingForCodepage(codepage), externSheetNames),
    );
  }

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
    saveExternalLinkValues,
    refreshAllConnections,
    builtInGroupCount,
    stream,
  };
}

export function readLegacyWorkbookGlobals(
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
      throw new XlsParseError("Encrypted legacy XLS files are not supported");
    if (!BEGIN_OF_FILE_CODES.has(record.code)) {
      cursor = nextCursor;
      continue;
    }
    if (sawBegin) throw new XlsParseError("Invalid legacy XLS file: nested workbook globals");
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
        throw new XlsParseError("Encrypted legacy XLS files are not supported");
      if (workbookRecord.code !== RecordCode.Sheethdr) continue;
      if (workbookRecord.body.byteLength < 5)
        throw new XlsParseError("Invalid legacy XLS file: truncated BIFF4W sheet header");
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

  if (!sawBegin) throw new XlsParseError("Invalid legacy XLS file: missing Begin Of File");
  if (sheets.length === 0)
    throw new XlsParseError("Invalid legacy XLS file: workbook has no sheet definitions");
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

export type Supbook =
  | { kind: "current" }
  | { kind: "internal"; sheetNames?: string[] }
  | { kind: "external"; target: string; sheetNames: string[] }
  | { kind: "dde"; target: string; sheetNames: string[] }
  | { kind: "ole"; target: string; sheetNames: string[] };

export function parseSupbook(record: BiffRecord, encoding: string): Supbook {
  const body = record.body;
  if (body.byteLength < 4)
    throw new XlsParseError(
      "Invalid legacy XLS file: truncated SUPBOOK",
      biffRecordContext(record, "truncated-supbook-count"),
    );
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const sheetCount = view.getUint16(0, true);
  const characterCount = view.getUint16(2, true);
  if (characterCount === 0x0401) return { kind: "internal" };
  let target: { value: string; offset: number };
  try {
    target = readNoCchBiff8String(body, 4, characterCount, encoding);
  } catch (error) {
    throw new XlsParseError("Invalid legacy XLS file: truncated SUPBOOK path", {
      ...biffRecordContext(record, "truncated-supbook-path"),
      cause: error,
    });
  }
  const sheetNames: string[] = [];
  for (let index = 0; index < sheetCount; index++) {
    try {
      const value = readShortBiff8String(body, target.offset, encoding);
      sheetNames.push(value.value);
      target.offset = value.offset;
    } catch (error) {
      if (sheetNames.length === 0)
        throw new XlsParseError("Invalid legacy XLS file: truncated SUPBOOK sheet name", {
          ...biffRecordContext(record, "truncated-supbook-sheet-name"),
          cause: error,
        });
      console.warn(
        `xls parse diagnostic: truncated SUPBOOK sheet list (${JSON.stringify({
          ...biffRecordContext(record, "truncated-sheet-list"),
          expectedSheetCount: sheetCount,
          parsedSheetCount: sheetNames.length,
        })})`,
      );
      break;
    }
  }
  if (target.value.includes("|")) return { kind: "dde", target: target.value, sheetNames };
  if (/^(ole:|package:)/i.test(target.value)) {
    return {
      kind: "ole",
      target: target.value.replace(/^(ole:|package:)/i, ""),
      sheetNames,
    };
  }
  return { kind: "external", target: target.value, sheetNames };
}

export function parseExternSheet(
  record: BiffRecord,
): { supbookIndex: number; firstSheet: number; lastSheet: number }[] {
  const body = record.body;
  if (body.byteLength < 2)
    throw new XlsParseError(
      "Invalid legacy XLS file: truncated EXTERNSHEET",
      biffRecordContext(record, "truncated-externsheet-count"),
    );
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const count = view.getUint16(0, true);
  if (body.byteLength < 2 + count * 6)
    throw new XlsParseError(
      "Invalid legacy XLS file: truncated EXTERNSHEET",
      biffRecordContext(record, "truncated-externsheet-reference"),
    );
  return Array.from({ length: count }, (_, index) => ({
    supbookIndex: view.getUint16(2 + index * 6, true),
    firstSheet: view.getUint16(4 + index * 6, true),
    lastSheet: view.getUint16(6 + index * 6, true),
  }));
}

function decodeBiff5EncodedReference(
  body: Uint8Array,
  encoding: string,
): { kind: "current" | "internal" | "external"; target: string; sheetName?: string } {
  const length = body[0]!;
  if (body.byteLength < 1 + length)
    throw new XlsParseError("Invalid legacy XLS file: truncated EXTERNSHEET name", {
      path: "workbook-globals/EXTERNSHEET",
      reason: "truncated-externsheet-name",
    });
  const decoder = new TextDecoder(encoding);
  const raw = decoder.decode(body.subarray(1, 1 + length));
  const marker = raw.charCodeAt(0);
  if (marker === 0x02) return { kind: "current", target: "" };
  if (marker === 0x03) return { kind: "internal", target: "", sheetName: raw.slice(1) };
  if (marker !== 0x01) {
    const separator = raw.indexOf("]");
    return separator >= 0
      ? { kind: "external", target: raw.slice(0, separator), sheetName: raw.slice(separator + 1) }
      : { kind: "external", target: raw };
  }
  let target = "";
  let cursor = 1;
  while (cursor < raw.length) {
    const control = raw.charCodeAt(cursor);
    cursor += 1;
    if (control === 0x01) {
      const value = raw[cursor]!;
      cursor += 1;
      target += value === "@" ? `\\${raw[cursor] ?? ""}` : `${value}:`;
      if (value === "@") cursor += 1;
    } else if (control === 0x03) {
      target += "\\";
    } else if (control === 0x04) {
      target += "../";
    } else if (control === 0x05) {
      const urlLength = raw.charCodeAt(cursor);
      cursor += 1;
      target += raw.slice(cursor, cursor + urlLength);
      cursor += urlLength;
    } else if (control === 0x09) {
      return { kind: "current", target: "" };
    } else if (control < 0x20) {
      continue;
    } else {
      target += raw[cursor - 1]!;
    }
  }
  const separator = target.lastIndexOf("]");
  return separator >= 0
    ? {
        kind: "external",
        target: target.slice(0, separator),
        sheetName: target.slice(separator + 1),
      }
    : { kind: "external", target };
}

export function parseBiff5ExternSheet(record: BiffRecord, encoding: string): Supbook {
  const reference = decodeBiff5EncodedReference(record.body, encoding);
  if (reference.kind === "current") return { kind: "current" };
  if (reference.kind === "internal")
    return { kind: "internal", sheetNames: [reference.sheetName ?? ""] };
  return {
    kind: "external",
    target: reference.target,
    sheetNames: [reference.sheetName ?? ""],
  };
}

export function parseName(
  record: BiffRecord,
  isBiff8: boolean,
  encoding: string,
  externSheetNames: Map<number, string>,
): DefinedNameOptions {
  const body = record.body;
  if (body.byteLength < 14)
    throw new XlsParseError(
      "Invalid legacy XLS file: truncated NAME",
      biffRecordContext(record, "truncated-defined-name"),
    );
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const flags = view.getUint16(0, true);
  const formulaLength = view.getUint16(4, true);
  const localSheetId = view.getUint16(8, true) === 0 ? undefined : view.getUint16(8, true) - 1;
  const name = isBiff8
    ? readNoCchBiff8String(body, 14, body[3]!, encoding)
    : readNoCchBiff5String(body, 14, body[3]!, encoding);
  let value = name.value;
  const tokens = body.subarray(name.offset, name.offset + formulaLength);
  const decodeNameFormula = (formulaVersion: number, formulaData: Uint8Array) =>
    decodeFormula(formulaData, {
      baseRow: 0,
      baseColumn: 0,
      formulaVersion,
      sheetNameByExternIndex: (index) => externSheetNames.get(index),
    });
  if (isBiff8) {
    value = decodeNameFormula(8, tokens).formula;
  } else {
    try {
      value = decodeNameFormula(8, tokens).formula;
    } catch {
      try {
        value = decodeNameFormula(8, tokens.subarray(0, -1)).formula;
      } catch {
        value = decodeNameFormula(5, tokens).formula;
      }
    }
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

export function readCString(body: Uint8Array): string {
  const end = body.indexOf(0);
  return new TextDecoder(encodingForCodepage(1252)).decode(
    body.subarray(0, end === -1 ? body.byteLength : end),
  );
}

export function readNulTerminatedUnicode(
  body: Uint8Array,
  offset: number,
): { value: string; offset: number } {
  if (offset + 4 > body.byteLength)
    throw new XlsParseError("Invalid legacy XLS file: truncated unicode string");
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const byteLength = view.getUint32(offset, true) * 2;
  offset += 4;
  if (offset + byteLength > body.byteLength)
    throw new XlsParseError("Invalid legacy XLS file: truncated unicode string");
  const raw = body.subarray(offset, offset + byteLength - 2);
  return { value: new TextDecoder("utf-16le").decode(raw), offset: offset + byteLength };
}

export function readByteCountUnicode(
  body: Uint8Array,
  offset: number,
): { value: string; offset: number } {
  if (offset + 4 > body.byteLength)
    throw new XlsParseError("Invalid legacy XLS file: truncated unicode string");
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const byteLength = view.getUint32(offset, true);
  offset += 4;
  if (offset + byteLength > body.byteLength)
    throw new XlsParseError("Invalid legacy XLS file: truncated unicode string");
  const raw = body.subarray(offset, offset + byteLength - 2);
  return { value: new TextDecoder("utf-16le").decode(raw), offset: offset + byteLength };
}

export function parseHyperlink(body: Uint8Array): HyperlinkOptions {
  if (body.byteLength < 32) throw new XlsParseError("Invalid legacy XLS file: truncated HLINK");
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

export function parseSharedStrings(
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
      throw new XlsParseError("Invalid legacy XLS file: truncated shared string table");
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
