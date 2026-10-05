import {
  CompoundFileReader,
  decryptLegacyRc4,
  decryptRc4CryptoApi,
  parseLegacyRc4Verifier,
  parseRc4CryptoApiHeader,
  verifyLegacyRc4Password,
  verifyRc4CryptoApiPassword,
} from "@office-open/core";
import type { WorkbookOptions, WorksheetOptions } from "@office-open/xlsx";

import { readRequiredStream } from "../cfb/container";
import { LegacyExcelError } from "../errors";
import { readSummaryInformation } from "../parts/metadata";
import { RecordCode, readRecord, recordHeaderSize, type BiffRecord } from "../records";
import { PLAINTEXT_ENCRYPTION_CODES, detectBiffVersion } from "./biff";
import { concatBytes } from "./cells";
import type { WorkbookState } from "./models";
import { readWorkbookGlobals } from "./workbook-globals";
import { parseWorksheetStream } from "./worksheet";

export interface LegacyParseOptions {
  /** Password used to verify and decrypt legacy Office RC4 containers. */
  password?: string;
}

function readWorkbookStream(data: Uint8Array): WorkbookState {
  if (data.byteLength > 8 && data[0] === 0x09) {
    const globals = readWorkbookGlobals(data);
    return { ...globals, stream: data };
  }
  const reader = new CompoundFileReader(data);
  const stream = readRequiredStream(reader, reader.entry("Workbook") ? "Workbook" : "Book");
  if (stream.byteLength < 8) {
    throw new LegacyExcelError("Invalid legacy XLS file: workbook stream is truncated", {
      part: "stream",
      path: "Workbook",
      offset: 0,
      length: 8,
      reason: "truncated-header",
    });
  }
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
  return new LegacyExcelError(error instanceof Error ? error.message : String(error), {
    part: "container",
    path: "/",
    reason: "invalid-container",
  });
}

function decryptWorkbookContainer(data: Uint8Array, password?: string): Uint8Array | undefined {
  let stream: Uint8Array;
  if (data.byteLength > 8 && data[0] === 0x09) {
    stream = data;
  } else {
    const reader = new CompoundFileReader(data);
    const entry = reader.entry("Workbook") ?? reader.entry("Book");
    if (!entry || entry.type !== "stream") {
      throw new LegacyExcelError("Invalid legacy XLS file: missing Workbook or Book stream", {
        part: "stream",
        path: "Workbook",
        reason: "missing-required-stream",
      });
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
    throw new LegacyExcelError("Encrypted legacy XLS files are not supported", {
      part: "stream",
      path: "Workbook",
      recordType: "0x002f",
      recordName: "FilePass",
      offset: filePass.offset,
      reason: "encrypted-unsupported",
    });
  }
  const body = filePass.body;
  if (body.byteLength < 2 || body[0] !== 1 || body[1] !== 0) {
    throw new LegacyExcelError("Encrypted legacy XLS files are not supported", {
      part: "stream",
      path: "Workbook",
      recordType: "0x002f",
      recordName: "FilePass",
      offset: filePass.offset,
      length: body.byteLength,
      reason: "invalid-encryption-version",
    });
  }
  const version = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const decrypt = (() => {
    const majorVersion = body.byteLength >= 4 ? version.getUint16(2, true) : 0;
    const minorVersion = body.byteLength >= 6 ? version.getUint16(4, true) : 0;
    if (majorVersion === 1 && minorVersion !== 1) {
      throw new LegacyExcelError("Encrypted legacy XLS files are not supported", {
        part: "stream",
        path: "Workbook",
        recordName: "FilePass",
        offset: filePass.offset,
        reason: "invalid-encryption-version",
      });
    }
    if (majorVersion !== 1 && (minorVersion !== 0x0002 || majorVersion < 2 || majorVersion > 4)) {
      throw new LegacyExcelError("Encrypted legacy XLS files are not supported", {
        part: "stream",
        path: "Workbook",
        recordName: "FilePass",
        offset: filePass.offset,
        reason: "invalid-encryption-version",
      });
    }
    if (majorVersion === 1) {
      const verifier = parseLegacyRc4Verifier(body, 6);
      if (!verifyLegacyRc4Password(password, verifier)) {
        throw new LegacyExcelError("Invalid legacy XLS password", {
          part: "stream",
          path: "Workbook",
          recordName: "FilePass",
          offset: filePass.offset,
          reason: "invalid-password",
        });
      }
      return (data: Uint8Array): Uint8Array =>
        decryptLegacyRc4(data, password, verifier.salt, 1024);
    }
    const { keySizeBits, verifier } = parseRc4CryptoApiHeader(body, 6);
    if (!verifyRc4CryptoApiPassword(password, verifier, keySizeBits)) {
      throw new LegacyExcelError("Invalid legacy XLS password", {
        part: "stream",
        path: "Workbook",
        recordName: "FilePass",
        offset: filePass.offset,
        reason: "invalid-password",
      });
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
