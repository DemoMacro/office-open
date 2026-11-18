import type {
  DefinedNameOptions,
  ExternalLinkOptions,
  FileSharingOptions,
  FontOptions,
  CalculationPropertiesOptions,
  HyperlinkOptions,
  RichTextOptions,
  CellStyleXfOptions,
  CustomCellStyleOptions,
  WorkbookProtectionOptions,
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
  readBiff8String,
  richTextValue,
  readNoCchBiff8String,
  readNoCchBiff5String,
  readShortBiff8String,
} from "./biff";
import { reference } from "./cells";
import type { BoundSheet, WorkbookState, XfStyle } from "./models";
import {
  applyXfExt,
  parseFont,
  parseFormat,
  parsePalette,
  parseStyle,
  parseStyleExt,
  parseXf,
  styleXfOptions,
  type LegacyStyleExtension,
  type LegacyStyleRecord,
} from "./styles";
export function readWorkbookGlobals(stream: Uint8Array): WorkbookState {
  const detected = detectBiffVersion(stream);
  if (detected.version <= 4) return readLegacyWorkbookGlobals(stream, detected);

  const workbookVersion = detected.version as 5 | 8;
  const sharedStrings: (string | RichTextOptions)[] = [];
  const sheets: BoundSheet[] = [];
  const fonts: FontOptions[] = [];
  const numberFormats = new Map<number, string>();
  const palette: string[] = [];
  const externalLinks: ExternalLinkOptions[] = [];
  const definedNames: DefinedNameOptions[] = [];
  const externSheetNames = new Map<number, string>();
  const styleTable = new Map<number, XfStyle>();
  const pendingStyleRecords: LegacyStyleRecord[] = [];
  const pendingStyleExtensions: LegacyStyleExtension[] = [];
  let lastStyleRecordIndex = -1;
  const drawingImages: { data: Uint8Array; type: "png" | "jpg" | "wmf" | "emf" }[] = [];
  let saveExternalLinkValues: boolean | undefined;
  let refreshAllConnections: boolean | undefined;
  let builtInGroupCount: number | undefined;
  let codeName: string | undefined;
  let calculation: CalculationPropertiesOptions | undefined;
  let workbookProtection: WorkbookProtectionOptions | undefined;
  let fileSharing: FileSharingOptions | undefined;
  let lastModifiedBy: string | undefined;
  let backupFile: boolean | undefined;
  const pendingSupbooks: Supbook[] = [];
  const pendingExternSheetReferences:
    | ReturnType<typeof parseExternSheet>
    | { supbookIndex: number; firstSheet: number; lastSheet: number }[] = [];
  const pendingNameRecords: BiffRecord[] = [];
  const pendingXfExtRecords: BiffRecord[] = [];
  let externSheetReferenceCount = 0;
  let codepage = 1252;
  let date1904 = false;
  let sharedStringParts: Uint8Array[] | undefined;
  let inGlobals = false;

  for (const record of recordsFrom(stream, 0, workbookVersion)) {
    if (record.code === RecordCode.FilePass && !isClearedFilePass(record))
      throw new XlsParseError("Encrypted legacy XLS files are not supported", {
        part: "stream",
        recordType: record.code,
        recordName: "FilePass",
        offset: record.start,
        reason: "encrypted-unsupported",
      });
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
          fonts,
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
      case RecordCode.CalculationCount: {
        const value = readGlobalUint16(record, "calculation count");
        if (value < 1 || value > 32767)
          throw invalidGlobalValue(record, "calculation count", "invalid-calculation-count");
        calculation = { ...calculation, iterateCount: value };
        break;
      }
      case RecordCode.CalculationMode: {
        const value = readGlobalUint16(record, "calculation mode");
        const calcMode: CalculationPropertiesOptions["calcMode"] =
          value === 0
            ? "manual"
            : value === 1
              ? "auto"
              : value === 0xffff
                ? "autoNoTable"
                : undefined;
        if (calcMode === undefined)
          throw invalidGlobalValue(record, "calculation mode", "invalid-calculation-mode");
        calculation = { ...calculation, calcMode };
        break;
      }
      case RecordCode.CalculationPrecision: {
        const value = readGlobalBoolean(record, "calculation precision");
        calculation = { ...calculation, fullPrecision: value };
        break;
      }
      case RecordCode.CalculationReferenceMode: {
        const value = readGlobalBoolean(record, "calculation reference mode");
        calculation = { ...calculation, refMode: value ? "A1" : "R1C1" };
        break;
      }
      case RecordCode.CalculationDelta: {
        if (record.body.byteLength !== 8)
          throw invalidGlobalLength(record, "calculation delta", "invalid-calculation-delta");
        const value = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        ).getFloat64(0, true);
        if (!Number.isFinite(value) || value < 0)
          throw invalidGlobalValue(record, "calculation delta", "invalid-calculation-delta");
        calculation = { ...calculation, iterateDelta: value };
        break;
      }
      case RecordCode.Iteration: {
        calculation = { ...calculation, iterate: readGlobalBoolean(record, "iteration") };
        break;
      }
      case RecordCode.CalculationSaveRecalc: {
        calculation = {
          ...calculation,
          calcOnSave: readGlobalBoolean(record, "calculation save recalc"),
        };
        break;
      }
      case RecordCode.Protect: {
        workbookProtection = {
          ...workbookProtection,
          lockStructure: readGlobalBoolean(record, "workbook protection"),
        };
        break;
      }
      case RecordCode.Password: {
        const value = readGlobalUint16(record, "workbook password");
        workbookProtection = {
          ...workbookProtection,
          workbookPasswordHashRaw: value.toString(16).padStart(4, "0").toUpperCase(),
        };
        break;
      }
      case RecordCode.WindowProtect: {
        workbookProtection = {
          ...workbookProtection,
          lockWindows: readGlobalBoolean(record, "window protection"),
        };
        break;
      }
      case RecordCode.RevisionProtection: {
        workbookProtection = {
          ...workbookProtection,
          lockRevision: readGlobalBoolean(record, "revision protection"),
        };
        break;
      }
      case RecordCode.RevisionProtectionPassword: {
        const value = readGlobalUint16(record, "revision password");
        workbookProtection = {
          ...workbookProtection,
          revisionsPasswordHashRaw: value.toString(16).padStart(4, "0").toUpperCase(),
        };
        break;
      }
      case RecordCode.Backup: {
        backupFile = readGlobalBoolean(record, "backup");
        break;
      }
      case RecordCode.FileSharing: {
        fileSharing = parseFileSharing(
          record,
          encodingForCodepage(codepage),
          workbookVersion === 8,
        );
        break;
      }
      case RecordCode.WriteAccess: {
        lastModifiedBy = parseWriteAccess(record, encodingForCodepage(codepage));
        break;
      }
      case RecordCode.CodeName: {
        if (workbookVersion !== 8 || codeName !== undefined)
          throw new XlsParseError("Invalid legacy XLS file: misplaced or duplicate CODENAME", {
            ...biffRecordContext(record, "misplaced-or-duplicate-codename"),
          });
        let decoded: { value: string; offset: number };
        try {
          decoded = readShortBiff8String(record.body, 0, encodingForCodepage(codepage));
        } catch (error) {
          throw new XlsParseError("Invalid legacy XLS file: malformed CODENAME", {
            ...biffRecordContext(record, "malformed-codename"),
            cause: error,
          });
        }
        if (decoded.offset !== record.body.byteLength || decoded.value.length > 31)
          throw new XlsParseError("Invalid legacy XLS file: invalid CODENAME length", {
            ...biffRecordContext(record, "invalid-codename-length"),
          });
        codeName = decoded.value;
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
      case RecordCode.Style: {
        pendingStyleRecords.push(
          parseStyle(record.body, workbookVersion === 8, encodingForCodepage(codepage), record),
        );
        lastStyleRecordIndex = pendingStyleRecords.length - 1;
        break;
      }
      case RecordCode.XfExt: {
        pendingXfExtRecords.push(record);
        break;
      }
      case RecordCode.StyleExt: {
        pendingStyleExtensions.push({
          ...parseStyleExt(record.body, palette, record),
          styleRecordIndex: lastStyleRecordIndex >= 0 ? lastStyleRecordIndex : undefined,
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
        fonts,
      ),
    );

  for (const record of pendingXfExtRecords) applyXfExt(record, styleTable, palette);

  const styleXfIndexes = [...styleTable].filter(([, xf]) => xf.isStyle).map(([index]) => index);
  const styleXfIndex = new Map(styleXfIndexes.map((index, styleIndex) => [index, styleIndex]));
  const cellStyleXfs: CellStyleXfOptions[] = styleXfIndexes.map((index) =>
    styleXfOptions(styleTable.get(index)!, fonts, numberFormats),
  );
  const cellStyles: CustomCellStyleOptions[] = [];
  const styleExtensions = new Map<number, LegacyStyleExtension>();
  const matchedStyleRecords = new Set<number>();
  for (const extension of pendingStyleExtensions) {
    const directIndex = extension.styleRecordIndex;
    const directRecord = directIndex === undefined ? undefined : pendingStyleRecords[directIndex];
    const index =
      directIndex !== undefined &&
      directRecord !== undefined &&
      !matchedStyleRecords.has(directIndex) &&
      styleRecordMatchesExtension(directRecord, extension)
        ? directIndex
        : pendingStyleRecords.findIndex(
            (styleRecord, styleRecordIndex) =>
              !matchedStyleRecords.has(styleRecordIndex) &&
              styleRecordMatchesExtension(styleRecord, extension),
          );
    if (index < 0) throw invalidStyleExtension(extension);
    matchedStyleRecords.add(index);
    styleExtensions.set(index, extension);
  }
  for (const [styleRecordIndex, styleRecord] of pendingStyleRecords.entries()) {
    const xf = styleTable.get(styleRecord.xfIndex);
    const xfId = styleXfIndex.get(styleRecord.xfIndex);
    if (!xf || xfId === undefined)
      throw new XlsParseError("Invalid legacy XLS file: Style references a missing XF", {
        path: "Workbook/globals/Style",
        recordName: "Style",
        recordType: RecordCode.Style,
        reason: "missing-style-xf",
      });
    const styleXf = cellStyleXfs[xfId]!;
    const extension = styleExtensions.get(styleRecordIndex);
    if (extension) mergeStyleXf(styleXf, extension.style);
    const name = extension?.name || styleRecord.name;
    if (name === undefined) continue;
    cellStyles.push({
      name,
      xfId,
      ...((extension?.builtinId ?? styleRecord.builtinId) === undefined
        ? {}
        : { builtinId: extension?.builtinId ?? styleRecord.builtinId }),
      ...(extension?.customBuiltin ? { customBuiltin: true } : {}),
      ...(extension?.hidden ? { hidden: true } : {}),
      ...(extension ? { category: extension.category } : {}),
      ...(styleRecord.iLevel === undefined ? {} : { iLevel: styleRecord.iLevel }),
    });
  }

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
    const resolvedSheetName = sheetName ?? "#REF!";
    if (workbookVersion === 5) {
      externSheetReferenceCount += 1;
      externSheetNames.set(externSheetReferenceCount, resolvedSheetName);
      externSheetNames.set(-externSheetReferenceCount, resolvedSheetName);
    } else {
      externSheetNames.set(++externSheetReferenceCount, resolvedSheetName);
    }
    if (supbook.kind === "external") {
      externalLinks.push({
        externalBook: { target: supbook.target, sheetNames: [resolvedSheetName] },
      });
    }
    if (supbook.kind === "dde") {
      const [ddeService = supbook.target, ddeTopic = ""] = supbook.target.split("|");
      externalLinks.push({
        ddeLink: { ddeService, ddeTopic, ddeItems: [{ name: resolvedSheetName }] },
      });
    }
    if (supbook.kind === "ole") {
      externalLinks.push({
        oleLink: { progId: supbook.target, oleItems: [{ name: resolvedSheetName }] },
      });
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
    cellStyles,
    cellStyleXfs,
    drawingImages,
    saveExternalLinkValues,
    refreshAllConnections,
    builtInGroupCount,
    codeName,
    calculation,
    workbookProtection,
    fileSharing,
    lastModifiedBy,
    backupFile,
    stream,
  };
}

function readGlobalUint16(record: BiffRecord, name: string): number {
  if (record.body.byteLength !== 2)
    throw invalidGlobalLength(record, name, "invalid-global-length");
  return new DataView(record.body.buffer, record.body.byteOffset, 2).getUint16(0, true);
}

function readGlobalBoolean(record: BiffRecord, name: string): boolean {
  const value = readGlobalUint16(record, name);
  if (value !== 0 && value !== 1) throw invalidGlobalValue(record, name, "invalid-global-boolean");
  return value === 1;
}

function invalidGlobalLength(record: BiffRecord, name: string, reason: string): XlsParseError {
  return new XlsParseError(`Invalid legacy XLS file: invalid ${name} record length`, {
    ...biffRecordContext(record, reason),
  });
}

function invalidGlobalValue(record: BiffRecord, name: string, reason: string): XlsParseError {
  return new XlsParseError(`Invalid legacy XLS file: invalid ${name} value`, {
    ...biffRecordContext(record, reason),
  });
}

function parseFileSharing(
  record: BiffRecord,
  encoding: string,
  isBiff8: boolean,
): FileSharingOptions {
  if (record.body.byteLength < 6)
    throw invalidGlobalLength(record, "file sharing", "truncated-file-sharing");
  const view = new DataView(record.body.buffer, record.body.byteOffset, record.body.byteLength);
  const readOnly = view.getUint16(0, true);
  if (readOnly !== 0 && readOnly !== 1)
    throw invalidGlobalValue(record, "file sharing read-only flag", "invalid-file-sharing-flag");
  const options: FileSharingOptions = {
    readOnlyRecommended: readOnly === 1,
    reservationPasswordHashRaw: view.getUint16(2, true).toString(16).padStart(4, "0").toUpperCase(),
  };
  const characterCount = view.getUint16(4, true);
  if (characterCount === 0) return options;
  try {
    const decoded = isBiff8
      ? readBiff8String(record.body, 4, encoding)
      : readNoCchBiff5String(record.body, 6, characterCount, encoding);
    if (decoded.offset > record.body.byteLength)
      throw new Error("file sharing username exceeds record");
    if (decoded.value.length > 0) options.userName = decoded.value;
  } catch (error) {
    throw new XlsParseError("Invalid legacy XLS file: malformed file sharing username", {
      ...biffRecordContext(record, "malformed-file-sharing-username"),
      cause: error,
    });
  }
  return options;
}

function parseWriteAccess(record: BiffRecord, encoding: string): string {
  if (record.body.byteLength < 3 || record.body.byteLength > 112)
    throw invalidGlobalLength(record, "write access", "invalid-write-access-length");
  const view = new DataView(record.body.buffer, record.body.byteOffset, record.body.byteLength);
  const characterCount = view.getUint16(0, true);
  const flags = record.body[2]!;
  if ((flags & 0xfe) !== 0 || characterCount > 109) {
    if (record.body.byteLength !== 32 && record.body.byteLength !== 112)
      throw invalidGlobalValue(record, "write access", "invalid-write-access-header");
    const highByte = record.body[1] === 0;
    const fallbackEncoding = encoding === "utf-16le" ? "windows-1252" : encoding;
    return trimWriteAccess(
      new TextDecoder(highByte ? "utf-16le" : fallbackEncoding).decode(record.body),
    );
  }
  const highByte = (flags & 1) !== 0;
  const byteLength = characterCount * (highByte ? 2 : 1);
  if (3 + byteLength > record.body.byteLength)
    throw invalidGlobalLength(record, "write access", "truncated-write-access");
  const value = new TextDecoder(highByte ? "utf-16le" : encoding).decode(
    record.body.subarray(3, 3 + byteLength),
  );
  return trimWriteAccess(value);
}

function trimWriteAccess(value: string): string {
  let end = value.length;
  while (end > 0 && (value.charCodeAt(end - 1) === 0 || value[end - 1] === " ")) end--;
  return value.slice(0, end);
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
      throw new XlsParseError("Encrypted legacy XLS files are not supported", {
        part: "stream",
        recordType: record.code,
        recordName: "FilePass",
        offset: record.start,
        reason: "encrypted-unsupported",
      });
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
        cellStyles: [],
        cellStyleXfs: [],
        drawingImages: [],
        stream,
      };
    }

    cursor = nextCursor;
    while (cursor < stream.byteLength) {
      const workbookRecord = readRecord(stream, cursor, detected.version);
      cursor += recordHeaderSize(detected.version) + workbookRecord.body.byteLength;
      if (workbookRecord.code === RecordCode.FilePass && !isClearedFilePass(workbookRecord))
        throw new XlsParseError("Encrypted legacy XLS files are not supported", {
          part: "stream",
          recordType: workbookRecord.code,
          recordName: "FilePass",
          offset: workbookRecord.start,
          reason: "encrypted-unsupported",
        });
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
    cellStyles: [],
    cellStyleXfs: [],
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
      throw new XlsParseError("Invalid legacy XLS file: truncated SUPBOOK sheet name", {
        ...biffRecordContext(record, "truncated-supbook-sheet-name"),
        cause: error,
      });
    }
  }
  let targetValue = "";
  for (let index = 0; index < target.value.length; index++) {
    const character = target.value[index]!;
    const code = character.charCodeAt(0);
    targetValue += code < 0x20 ? `%${code.toString(16).padStart(2, "0").toUpperCase()}` : character;
  }
  if (targetValue.includes("|")) return { kind: "dde", target: targetValue, sheetNames };
  if (/^(ole:|package:)/i.test(targetValue)) {
    return {
      kind: "ole",
      target: targetValue.replace(/^(ole:|package:)/i, ""),
      sheetNames,
    };
  }
  return { kind: "external", target: targetValue, sheetNames };
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
  const decodedName = isBiff8
    ? readNoCchBiff8String(body, 14, body[3]!, encoding)
    : readNoCchBiff5String(body, 14, body[3]!, encoding);
  const builtInNames: Record<number, string> = {
    0x00: "_xlnm.Consolidate_Area",
    0x01: "_xlnm.Auto_Open",
    0x02: "_xlnm.Auto_Close",
    0x03: "_xlnm.Extract",
    0x04: "_xlnm.Database",
    0x05: "_xlnm.Criteria",
    0x06: "_xlnm.Print_Area",
    0x07: "_xlnm.Print_Titles",
    0x08: "_xlnm.Recorder",
    0x09: "_xlnm.Data_Form",
    0x0a: "_xlnm.Auto_Activate",
    0x0b: "_xlnm.Auto_Deactivate",
    0x0c: "_xlnm.Sheet_Title",
    0x0d: "_xlnm._FilterDatabase",
  };
  const builtIn = (flags & 0x0020) !== 0;
  const name =
    builtIn && decodedName.value.length === 1 && builtInNames[decodedName.value.charCodeAt(0)]
      ? builtInNames[decodedName.value.charCodeAt(0)]!
      : decodedName.value;
  let value = name;
  const tokens = body.subarray(decodedName.offset, decodedName.offset + formulaLength);
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
    name,
    value,
    localSheetId,
    hidden: (flags & 0x0001) !== 0,
    function: (flags & 0x0002) !== 0,
    vbProcedure: (flags & 0x0004) !== 0,
  };
}

function mergeStyleXf(target: CellStyleXfOptions, source: CellStyleXfOptions): void {
  if (source.font) target.font = { ...target.font, ...source.font };
  if (source.fill) target.fill = { ...target.fill, ...source.fill };
  if (source.border) target.border = { ...target.border, ...source.border };
  if (source.alignment) target.alignment = { ...target.alignment, ...source.alignment };
  if (source.protection) target.protection = { ...target.protection, ...source.protection };
  if (source.numFmt !== undefined) target.numFmt = source.numFmt;
  if (source.numFmtId !== undefined) target.numFmtId = source.numFmtId;
  if (source.quotePrefix !== undefined) target.quotePrefix = source.quotePrefix;
}

function invalidStyleExtension(extension: LegacyStyleExtension): XlsParseError {
  return new XlsParseError("Invalid legacy XLS file: StyleExt does not match a Style record", {
    path: "Workbook/globals/StyleExt",
    recordName: "StyleExt",
    recordType: RecordCode.StyleExt,
    reason: "style-record-mismatch",
    offset: extension.record.start,
    length: extension.record.body.byteLength,
    ...(extension.styleRecordIndex === undefined
      ? {}
      : { styleRecordIndex: extension.styleRecordIndex }),
  });
}

function styleRecordMatchesExtension(
  styleRecord: LegacyStyleRecord,
  extension: LegacyStyleExtension,
): boolean {
  if (styleRecord.builtIn) return styleRecord.builtInData === extension.builtInData;
  return styleRecord.name === extension.name;
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
  const raw = nulTerminatedRaw(body.subarray(offset, offset + byteLength - 2));
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
  const raw = nulTerminatedRaw(body.subarray(offset, offset + byteLength - 2));
  return { value: new TextDecoder("utf-16le").decode(raw), offset: offset + byteLength };
}

function nulTerminatedRaw(raw: Uint8Array): Uint8Array {
  for (let offset = 0; offset + 1 < raw.byteLength; offset += 2) {
    if (raw[offset] === 0 && raw[offset + 1] === 0) return raw.subarray(0, offset);
  }
  return raw;
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
  fonts: readonly FontOptions[] = [],
): (string | RichTextOptions)[] {
  const first = parts[0]!;
  const firstView = new DataView(first.buffer, first.byteOffset, first.byteLength);
  const stringCount = firstView.getUint32(4, true);
  const strings: (string | RichTextOptions)[] = [];
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
      throw new XlsParseError("Invalid legacy XLS file: truncated shared string table");
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
    const value = new TextDecoder(highByte ? "utf-16le" : encoding).decode(
      new Uint8Array(characterBytes),
    );
    const richRuns = Array.from({ length: richRunCount }, () => ({
      characterIndex: takeUint16(),
      fontIndex: takeUint16(),
    }));
    skipBytes(extensionLength);
    strings.push(richTextValue(value, richRuns, fonts));
  }
  return strings;
}
