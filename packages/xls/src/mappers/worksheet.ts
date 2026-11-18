import type {
  ConditionalFormatRule,
  DataValidationOptions,
  NoteAnchorOptions,
  WorksheetOptions,
} from "@office-open/xlsx";

import { XlsParseError } from "../errors";
import type { BiffRecord } from "../records";
import { RecordCode, isClearedFilePass, recordsFrom } from "../records";
import { assertSupportedBiffRecord, biffRecordContext } from "../records/biff-records";
import { escherAnchors, escherDrawings } from "../records/escher";
import { decodeFormula } from "../records/formula";
import {
  BEGIN_OF_FILE_CODES,
  encodingForCodepage,
  readBiff2String,
  readBiff35String,
  readBiff5String,
  readBiff8String,
  readShortBiff8String,
  richTextValue,
} from "./biff";
import {
  cachedFormulaResult,
  concatBytes,
  decodeParts,
  decodeRk,
  ensureRow,
  errorLiteral,
  formulaOptions,
  makeCell,
  reference,
} from "./cells";
import type {
  FormulaLocation,
  PendingConditional,
  SheetState,
  TxoState,
  TxoText,
  WorkbookState,
} from "./models";
import { parseHyperlink } from "./workbook-globals";

type ConditionalOperator = NonNullable<ConditionalFormatRule["operator"]>;

function dataView(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

function parsePageMargins(
  record: BiffRecord,
  margins: WorksheetOptions["pageMargins"],
): NonNullable<WorksheetOptions["pageMargins"]> {
  if (record.body.byteLength < 8)
    throw new XlsParseError("Invalid legacy XLS file: truncated page margin");
  const value = dataView(record.body).getFloat64(0, true);
  if (record.code === RecordCode.LeftMargin) return { ...margins, left: value };
  if (record.code === RecordCode.RightMargin) return { ...margins, right: value };
  if (record.code === RecordCode.TopMargin) return { ...margins, top: value };
  return { ...margins, bottom: value };
}

function parsePageSetup(record: BiffRecord): {
  pageSetup: NonNullable<WorksheetOptions["pageSetup"]>;
  headerFooterMargins: Pick<NonNullable<WorksheetOptions["pageMargins"]>, "header" | "footer">;
} {
  if (record.body.byteLength < 32)
    throw new XlsParseError("Invalid legacy XLS file: truncated PAGESETUP");
  const view = dataView(record.body);
  const flags = view.getUint16(10, true);
  return {
    pageSetup: {
      paperSize: view.getUint16(0, true),
      scale: view.getUint16(2, true),
      fitToWidth: view.getUint16(6, true),
      fitToHeight: view.getUint16(8, true),
      pageOrder: (flags & 1) === 0 ? "downThenOver" : "overThenDown",
      orientation: (flags & 2) === 0 ? "landscape" : "portrait",
    },
    headerFooterMargins: {
      header: view.getFloat64(16, true),
      footer: view.getFloat64(24, true),
    },
  };
}

const DATA_VALIDATION_TYPES: Readonly<Record<number, DataValidationOptions["type"]>> = {
  0: "none",
  1: "whole",
  2: "decimal",
  3: "list",
  4: "date",
  5: "time",
  6: "textLength",
  7: "custom",
};

const DATA_VALIDATION_OPERATORS: Readonly<Record<number, DataValidationOptions["operator"]>> = {
  0: "between",
  1: "notBetween",
  2: "equal",
  3: "notEqual",
  4: "greaterThan",
  5: "lessThan",
  6: "greaterThanOrEqual",
  7: "lessThanOrEqual",
};

const DATA_VALIDATION_ERROR_STYLES: Readonly<
  Record<number, NonNullable<DataValidationOptions["errorStyle"]>>
> = {
  0: "stop",
  1: "warning",
  2: "information",
};

const DATA_VALIDATION_IME_MODES: Readonly<
  Record<number, NonNullable<DataValidationOptions["imeMode"]>>
> = {
  0: "noControl",
  1: "on",
  2: "off",
  3: "disabled",
  4: "hiragana",
  5: "fullKatakana",
  6: "halfKatakana",
  7: "fullAlpha",
  8: "halfAlpha",
  9: "fullHangul",
  10: "halfHangul",
};

interface DataValidationRange {
  firstRow: number;
  lastRow: number;
  firstColumn: number;
  lastColumn: number;
}

function parseDataValidationText(
  body: Uint8Array,
  offset: number,
  state: WorkbookState,
  description: string,
): { value: string | undefined; offset: number } {
  let decoded: ReturnType<typeof readBiff8String>;
  try {
    decoded = readBiff8String(body, offset, encodingForCodepage(state.codepage));
  } catch (error) {
    throw new XlsParseError(`Invalid legacy XLS file: malformed data validation ${description}`, {
      part: "stream",
      path: "worksheet/data-validations",
      reason: "malformed-data-validation-text",
      cause: error,
    });
  }
  return { value: decoded.value === "\0" ? undefined : decoded.value, offset: decoded.offset };
}

function parseDataValidationFormula(
  body: Uint8Array,
  offset: number,
  description: string,
): { tokens: Uint8Array | undefined; offset: number; length: number } {
  if (offset + 4 > body.byteLength)
    throw new XlsParseError(`Invalid legacy XLS file: truncated data validation ${description}`, {
      part: "stream",
      path: "worksheet/data-validations",
      reason: "truncated-data-validation-formula",
    });
  const view = dataView(body);
  const length = view.getUint16(offset, true);
  const formulaStart = offset + 4;
  const formulaEnd = formulaStart + length;
  if (formulaEnd > body.byteLength)
    throw new XlsParseError(`Invalid legacy XLS file: truncated data validation ${description}`, {
      part: "stream",
      path: "worksheet/data-validations",
      reason: "truncated-data-validation-formula",
    });
  return {
    tokens: length === 0 ? undefined : body.subarray(formulaStart, formulaEnd),
    offset: formulaEnd,
    length,
  };
}

function decodeDataValidationFormula(
  tokens: Uint8Array | undefined,
  range: DataValidationRange,
  description: string,
): string | undefined {
  if (!tokens) return undefined;
  try {
    return decodeFormula(tokens, {
      baseRow: range.firstRow,
      baseColumn: range.firstColumn,
    }).formula;
  } catch (error) {
    throw new XlsParseError(`Invalid legacy XLS file: malformed data validation ${description}`, {
      part: "stream",
      path: "worksheet/data-validations",
      reason: "malformed-data-validation-formula",
      cause: error,
    });
  }
}

function parseDataValidationRanges(
  body: Uint8Array,
  offset: number,
): { ranges: DataValidationRange[]; reference: string; offset: number } {
  if (offset + 2 > body.byteLength)
    throw new XlsParseError("Invalid legacy XLS file: truncated data validation ranges", {
      part: "stream",
      path: "worksheet/data-validations",
      reason: "truncated-data-validation-ranges",
    });
  const view = dataView(body);
  const count = view.getUint16(offset, true);
  offset += 2;
  if (count === 0 || count > 432 || offset + count * 8 !== body.byteLength)
    throw new XlsParseError("Invalid legacy XLS file: invalid data validation ranges", {
      part: "stream",
      path: "worksheet/data-validations",
      reason: "invalid-data-validation-ranges",
    });
  const ranges: DataValidationRange[] = [];
  for (let index = 0; index < count; index++) {
    const firstRow = view.getUint16(offset, true);
    const lastRow = view.getUint16(offset + 2, true);
    const firstColumn = view.getUint16(offset + 4, true);
    const lastColumn = view.getUint16(offset + 6, true);
    if (firstRow > lastRow || firstColumn > lastColumn || lastColumn > 255)
      throw new XlsParseError("Invalid legacy XLS file: invalid data validation range", {
        part: "stream",
        path: "worksheet/data-validations",
        reason: "invalid-data-validation-range",
      });
    ranges.push({ firstRow, lastRow, firstColumn, lastColumn });
    offset += 8;
  }
  return {
    ranges,
    reference: ranges
      .map(
        ({ firstRow, lastRow, firstColumn, lastColumn }) =>
          `${reference(firstRow, firstColumn)}:${reference(lastRow, lastColumn)}`,
      )
      .join(" "),
    offset,
  };
}

function parseDataValidation(record: BiffRecord, state: WorkbookState): DataValidationOptions {
  if (state.version !== 8 || record.body.byteLength < 4)
    throw new XlsParseError("Invalid legacy XLS file: invalid DV record", {
      ...biffRecordContext(record, "invalid-data-validation-record"),
    });
  const view = dataView(record.body);
  const flags = view.getUint32(0, true);
  const type = DATA_VALIDATION_TYPES[flags & 0x0f];
  const errorStyle = DATA_VALIDATION_ERROR_STYLES[(flags >> 4) & 0x07];
  const imeMode = DATA_VALIDATION_IME_MODES[(flags >> 10) & 0xff];
  const operatorValue = DATA_VALIDATION_OPERATORS[(flags >> 20) & 0x0f];
  const operatorIsDefined = type !== "none" && type !== "list" && type !== "custom";
  if (!type || !errorStyle || !imeMode || (operatorIsDefined && !operatorValue))
    throw new XlsParseError("Invalid legacy XLS file: invalid DV flags", {
      ...biffRecordContext(record, "invalid-data-validation-flags"),
    });
  let offset = 4;
  const promptTitle = parseDataValidationText(record.body, offset, state, "prompt title");
  offset = promptTitle.offset;
  const errorTitle = parseDataValidationText(record.body, offset, state, "error title");
  offset = errorTitle.offset;
  const prompt = parseDataValidationText(record.body, offset, state, "prompt");
  offset = prompt.offset;
  const error = parseDataValidationText(record.body, offset, state, "error");
  offset = error.offset;
  if (
    (promptTitle.value?.length ?? 0) > 32 ||
    (errorTitle.value?.length ?? 0) > 32 ||
    (prompt.value?.length ?? 0) > 255 ||
    (error.value?.length ?? 0) > 225
  )
    throw new XlsParseError("Invalid legacy XLS file: data validation text exceeds its limit", {
      ...biffRecordContext(record, "invalid-data-validation-text-length"),
    });
  if (offset + 4 > record.body.byteLength)
    throw new XlsParseError("Invalid legacy XLS file: truncated DV formula", {
      ...biffRecordContext(record, "truncated-data-validation-formula"),
    });
  const formula1 = parseDataValidationFormula(record.body, offset, "formula1");
  const formula2 = parseDataValidationFormula(record.body, formula1.offset, "formula2");
  const ranges = parseDataValidationRanges(record.body, formula2.offset);
  if (ranges.offset !== record.body.byteLength)
    throw new XlsParseError("Invalid legacy XLS file: malformed DV record", {
      ...biffRecordContext(record, "malformed-data-validation-record"),
    });
  if ((type === "none" || type === "list" || type === "custom") && formula2.length !== 0)
    throw new XlsParseError("Invalid legacy XLS file: unexpected DV formula2", {
      ...biffRecordContext(record, "unexpected-data-validation-formula"),
    });
  if (type === "none" && formula1.length !== 0)
    throw new XlsParseError("Invalid legacy XLS file: unexpected DV formula1", {
      ...biffRecordContext(record, "unexpected-data-validation-formula"),
    });
  if (type !== "none" && formula1.length === 0)
    throw new XlsParseError("Invalid legacy XLS file: missing DV formula1", {
      ...biffRecordContext(record, "missing-data-validation-formula"),
    });
  if (
    type !== "none" &&
    type !== "list" &&
    type !== "custom" &&
    (operatorValue === "between" || operatorValue === "notBetween") !== formula2.length > 0
  )
    throw new XlsParseError("Invalid legacy XLS file: invalid DV formula2 for operator", {
      ...biffRecordContext(record, "invalid-data-validation-formula-pair"),
    });
  const decodedFormula1 = decodeDataValidationFormula(
    formula1.tokens,
    ranges.ranges[0]!,
    "formula1",
  );
  const decodedFormula2 = decodeDataValidationFormula(
    formula2.tokens,
    ranges.ranges[0]!,
    "formula2",
  );
  return {
    sqref: ranges.reference,
    type,
    ...(type === "none" || decodedFormula1 === undefined ? {} : { formula1: decodedFormula1 }),
    ...(decodedFormula2 === undefined ? {} : { formula2: decodedFormula2 }),
    ...(operatorIsDefined ? { operator: operatorValue } : {}),
    errorStyle,
    imeMode,
    allowBlank: (flags & 0x00000100) !== 0,
    ...(type === "list" ? { showDropDown: (flags & 0x00000200) !== 0 } : {}),
    showInputMessage: (flags & 0x00040000) !== 0,
    showErrorMessage: (flags & 0x00080000) !== 0,
    ...(promptTitle.value === undefined ? {} : { promptTitle: promptTitle.value }),
    ...(errorTitle.value === undefined ? {} : { errorTitle: errorTitle.value }),
    ...(prompt.value === undefined ? {} : { prompt: prompt.value }),
    ...(error.value === undefined ? {} : { error: error.value }),
  };
}

function parseHeaderText(record: BiffRecord, state: WorkbookState): string {
  const encoding = encodingForCodepage(state.codepage);
  if (record.body.byteLength === 0) return "";
  if (state.version === 2) return readBiff2String(record.body, 0, encoding).value;
  if (state.version <= 4) return readBiff35String(record.body, 0, encoding).value;
  if (state.version === 5) return readBiff5String(record.body, 0, encoding).value;
  return readBiff8String(record.body).value;
}

function parseConditionalFormatRange(record: BiffRecord): { reference: string; remaining: number } {
  if (record.body.byteLength < 14)
    throw new XlsParseError("Invalid legacy XLS file: truncated CONDFMT");
  const view = dataView(record.body);
  const rangeCount = view.getUint16(12, true);
  if (record.body.byteLength < 14 + rangeCount * 8)
    throw new XlsParseError("Invalid legacy XLS file: truncated CONDFMT ranges");
  return {
    reference: `${reference(view.getUint16(4, true), view.getUint16(8, true))}:${reference(
      view.getUint16(6, true),
      view.getUint16(10, true),
    )}`,
    remaining: view.getUint16(0, true),
  };
}

function parseConditionalRule(record: BiffRecord, priority: number): ConditionalFormatRule {
  if (record.body.byteLength < 12)
    throw new XlsParseError("Invalid legacy XLS file: conditional format rule outside CONDFMT");
  const view = dataView(record.body);
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
    if (offset + length > record.body.byteLength)
      throw new XlsParseError("Invalid legacy XLS file: truncated conditional formula", {
        part: "stream",
        recordType: record.code,
        recordName: "CF",
        offset: record.start,
        reason: "malformed-formula",
      });
    const bytes = record.body.subarray(offset, offset + length);
    offset += length;
    return decodeFormula(bytes, {
      baseRow: 0,
      baseColumn: 0,
      arrayData: record.body.subarray(offset),
    }).formula;
  };
  const formulas = [decodeRuleFormula(firstLength), decodeRuleFormula(secondLength)].filter(
    (formula): formula is string => formula !== undefined,
  );
  return {
    type: type === 1 ? "expression" : "cellIs",
    operator: CONDITIONAL_OPERATORS[operatorCode] ?? "between",
    formulas,
    priority,
  };
}

const CONDITIONAL_OPERATORS: Partial<Record<number, ConditionalOperator>> = {
  3: "equal",
  4: "notEqual",
  5: "greaterThan",
  6: "lessThan",
  7: "greaterThanOrEqual",
  8: "lessThanOrEqual",
};

function formulaDecodeError(error: unknown, record: BiffRecord): XlsParseError {
  const source = error instanceof XlsParseError ? error : undefined;
  return new XlsParseError(
    source?.message ?? "Invalid legacy XLS file: malformed formula token stream",
    {
      part: "stream",
      recordType: record.code,
      recordName: "FORMULA",
      offset: record.start,
      reason: "malformed-formula",
      cause: error,
    },
  );
}

function setCellFormula(
  sheet: SheetState,
  row: number,
  column: number,
  state: WorkbookState,
  formula: NonNullable<ReturnType<typeof formulaOptions>>,
): void {
  const rowOptions = ensureRow(sheet, row);
  const cell = rowOptions.cells.find((candidate) => candidate.reference === reference(row, column));
  if (cell) {
    cell.formula = formula;
    return;
  }
  rowOptions.cells.push({
    ...makeCell(row, column, 0, cachedFormulaResult(new Uint8Array(8)), state),
    formula,
  });
}
export function parseWorksheetStream(
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
  const noteAnchorsByObjectId = new Map<number, NoteAnchorOptions>();
  let expectedDataValidationCount: number | undefined;
  let dataValidationCount = 0;

  const addDrawings = (bytes: Uint8Array) => {
    const drawings = escherDrawings(bytes, state.drawingImages);
    if (drawings.images.length) worksheetExtras.images = drawings.images;
    if (drawings.shapes.length) worksheetExtras.shapes = drawings.shapes;
    if (drawings.connectors.length) worksheetExtras.connectors = drawings.connectors;
    for (const { anchor, shapeId } of escherAnchors(bytes)) {
      if (shapeId !== undefined) noteAnchorsByObjectId.set(shapeId, anchor);
    }
  };

  for (const record of recordsFrom(stream, position, state.version, false)) {
    if (BEGIN_OF_FILE_CODES.has(record.code)) {
      depth++;
      continue;
    }
    if (depth === 0) continue;
    if (record.code === RecordCode.FilePass && !isClearedFilePass(record))
      throw new XlsParseError("Encrypted legacy XLS files are not supported", {
        part: "stream",
        recordType: record.code,
        recordName: "FilePass",
        offset: record.start,
        reason: "encrypted-unsupported",
      });
    if (record.code === RecordCode.EndOfFile) {
      if (depth === 1 && currentDrawing) addDrawings(concatBytes(currentDrawing));
      depth--;
      if (depth === 0) {
        if (
          (expectedDataValidationCount !== undefined &&
            expectedDataValidationCount !== dataValidationCount) ||
          (expectedDataValidationCount === undefined && dataValidationCount !== 0)
        )
          throw new XlsParseError(
            "Invalid legacy XLS file: DVAL record count does not match DV records",
            {
              part: "stream",
              path: "worksheet/data-validations",
              reason: "data-validation-count-mismatch",
            },
          );
        return { ...sheet, extras: worksheetExtras };
      }
      continue;
    }
    if (depth !== 1) continue;
    if (
      currentDrawing &&
      record.code !== RecordCode.Continue &&
      record.code !== RecordCode.MsoDrawing
    ) {
      addDrawings(concatBytes(currentDrawing));
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
        throw new XlsParseError("Invalid legacy XLS file: truncated formula string");
      }
      ensureRow(sheet, pendingFormula.row).cells.push({
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
          throw new XlsParseError("Invalid legacy XLS file: truncated dimensions");
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
          throw new XlsParseError("Invalid legacy XLS file: truncated row");
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
          throw new XlsParseError("Invalid legacy XLS file: truncated blank cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        ensureRow(sheet, row).cells.push({
          reference: reference(row, column),
          value: null,
          style: view.getUint16(4, true),
        });
        break;
      }
      case RecordCode.BoolErr:
      case RecordCode.Biff2BoolErr: {
        if (record.body.byteLength < 8)
          throw new XlsParseError("Invalid legacy XLS file: truncated boolean/error cell");
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
        ensureRow(sheet, row).cells.push(
          makeCell(row, column, style, isError ? errorLiteral(value) : value !== 0, state),
        );
        break;
      }
      case RecordCode.MulBlank: {
        if (record.body.byteLength < 10)
          throw new XlsParseError("Invalid legacy XLS file: truncated multiple blank cells");
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
          throw new XlsParseError(
            "Invalid legacy XLS file: multiple blank cell range is truncated",
          );
        }
        const rowOptions = ensureRow(sheet, row);
        for (let index = 0; index < count; index++) {
          const column = firstColumn + index;
          const style = view.getUint16(4 + index * 2, true);
          rowOptions.cells.push({ reference: reference(row, column), value: null, style });
        }
        break;
      }
      case RecordCode.Biff2Number:
      case RecordCode.Number: {
        if (state.version === 2) {
          if (record.body.byteLength < 15)
            throw new XlsParseError("Invalid legacy XLS file: truncated BIFF2 number cell");
          const view = new DataView(
            record.body.buffer,
            record.body.byteOffset,
            record.body.byteLength,
          );
          ensureRow(sheet, view.getUint16(0, true)).cells.push(
            makeCell(view.getUint16(0, true), view.getUint16(2, true), 0, view.getFloat64(7, true)),
          );
          break;
        }
        if (record.body.byteLength < 14)
          throw new XlsParseError("Invalid legacy XLS file: truncated number cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        ensureRow(sheet, row).cells.push(
          makeCell(row, column, view.getUint16(4, true), view.getFloat64(6, true), state),
        );
        break;
      }
      case RecordCode.Biff2Integer: {
        if (record.body.byteLength < 9)
          throw new XlsParseError("Invalid legacy XLS file: truncated BIFF2 integer cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        ensureRow(sheet, view.getUint16(0, true)).cells.push(
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
          throw new XlsParseError("Invalid legacy XLS file: truncated label cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        const style = view.getUint16(4, true);
        let text: string;
        let richRuns: ReturnType<typeof readBiff8String>["richRuns"];
        try {
          if (state.version === 8) {
            const decoded = readBiff8String(record.body, 6);
            text = decoded.value;
            richRuns = decoded.richRuns;
          } else {
            text = readBiff35String(record.body, 6).value;
          }
        } catch {
          throw new XlsParseError("Invalid legacy XLS file: truncated label cell");
        }
        ensureRow(sheet, row).cells.push(
          makeCell(row, column, style, richTextValue(text, richRuns, state.fonts), state),
        );
        break;
      }
      case RecordCode.Biff2Label: {
        if (record.body.byteLength < 8)
          throw new XlsParseError("Invalid legacy XLS file: truncated BIFF2 label cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        let text: { value: string; offset: number };
        try {
          text = readBiff2String(record.body, 7);
        } catch {
          throw new XlsParseError("Invalid legacy XLS file: truncated BIFF2 label cell");
        }
        ensureRow(sheet, view.getUint16(0, true)).cells.push(
          makeCell(view.getUint16(0, true), view.getUint16(2, true), 0, text.value, state),
        );
        break;
      }
      case RecordCode.RichLabel: {
        if (record.body.byteLength < 7)
          throw new XlsParseError("Invalid legacy XLS file: truncated rich label cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        const style = view.getUint16(4, true);
        let text: ReturnType<typeof readBiff8String>;
        try {
          text = readBiff8String(record.body, 6);
        } catch {
          throw new XlsParseError("Invalid legacy XLS file: truncated rich label cell");
        }
        ensureRow(sheet, row).cells.push(
          makeCell(
            row,
            column,
            style,
            richTextValue(text.value, text.richRuns, state.fonts),
            state,
          ),
        );
        break;
      }
      case RecordCode.LabelSst: {
        if (record.body.byteLength < 10)
          throw new XlsParseError("Invalid legacy XLS file: truncated shared string cell");
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
          throw new XlsParseError(`Invalid legacy XLS file: missing shared string ${index}`);
        ensureRow(sheet, row).cells.push(makeCell(row, column, style, value, state));
        break;
      }
      case RecordCode.Rk: {
        if (record.body.byteLength < 10)
          throw new XlsParseError("Invalid legacy XLS file: truncated RK cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        ensureRow(sheet, row).cells.push(
          makeCell(row, column, view.getUint16(4, true), decodeRk(view.getInt32(6, true)), state),
        );
        break;
      }
      case RecordCode.MulRk: {
        if (record.body.byteLength < 10)
          throw new XlsParseError("Invalid legacy XLS file: truncated multiple RK cells");
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
          throw new XlsParseError("Invalid legacy XLS file: multiple RK cell range is truncated");
        }
        const rowOptions = ensureRow(sheet, row);
        for (let index = 0; index < count; index++) {
          const offset = 4 + index * 6;
          const column = firstColumn + index;
          const style = view.getUint16(offset, true);
          const value = decodeRk(view.getInt32(offset + 2, true));
          rowOptions.cells.push(makeCell(row, column, style, value, state));
        }
        break;
      }
      case RecordCode.Formula: {
        if (record.body.byteLength < 22)
          throw new XlsParseError("Invalid legacy XLS file: truncated formula cell");
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
          const arrayData = record.body.subarray(22 + tokenLength);
          try {
            expression =
              decodeFormula(tokens, {
                baseRow: row,
                baseColumn: column,
                formulaVersion: state.version,
                sheetNameByExternIndex: (index) => state.externSheetNames.get(index),
                arrayData,
              }).formula || undefined;
          } catch (error) {
            throw formulaDecodeError(error, record);
          }
        }
        if (result === undefined) {
          formula = { row, column, style: view.getUint16(4, true), expression };
          pendingFormula = formula;

          continue;
        }
        ensureRow(sheet, row).cells.push({
          ...makeCell(row, column, view.getUint16(4, true), result, state),
          formula: formulaOptions(expression),
        });
        break;
      }
      case RecordCode.SharedFormula: {
        if (record.body.byteLength < 10) {
          throw new XlsParseError("Invalid legacy XLS file: truncated shared formula", {
            part: "stream",
            recordType: record.code,
            recordName: "SHRFMLA",
            offset: record.start,
            length: 10,
            reason: "invalid-record-length",
          });
        }
        const sharedView = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const sharedRow = sharedView.getUint16(0, true);
        const sharedColumn = sharedView.getUint16(4, true);
        const tokens = record.body.subarray(10);
        try {
          const decoded = decodeFormula(tokens, {
            baseRow: sharedRow,
            baseColumn: sharedColumn,
            formulaVersion: state.version,
            sheetNameByExternIndex: (index) => state.externSheetNames.get(index),
          });
          setCellFormula(sheet, sharedRow, sharedColumn, state, {
            formula: decoded.formula,
            type: "shared",
            sharedIndex: sharedFormulaIndex++,
          });
        } catch (error) {
          throw formulaDecodeError(error, record);
        }
        break;
      }
      case RecordCode.ArrayFormula: {
        if (record.body.byteLength < 14) {
          throw new XlsParseError("Invalid legacy XLS file: truncated array formula", {
            part: "stream",
            recordType: record.code,
            recordName: "ARRAY",
            offset: record.start,
            length: 14,
            reason: "invalid-record-length",
          });
        }
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
            baseRow: firstRow,
            baseColumn: firstColumn,
            formulaVersion: state.version,
            sheetNameByExternIndex: (index) => state.externSheetNames.get(index),
          });
          setCellFormula(sheet, firstRow, firstColumn, state, {
            formula: decoded.formula,
            type: "array",
            reference: referenceText,
          });
        } catch (error) {
          throw formulaDecodeError(error, record);
        }
        break;
      }
      case RecordCode.ColInfo: {
        if (record.body.byteLength < 10)
          throw new XlsParseError("Invalid legacy XLS file: truncated COLINFO");
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
      case RecordCode.DefaultColWidth: {
        if (record.body.byteLength < 2)
          throw new XlsParseError("Invalid legacy XLS file: truncated default column width");
        worksheetExtras.sheetFormat = {
          ...(worksheetExtras.sheetFormat === false ? {} : worksheetExtras.sheetFormat),
          baseColWidth: dataView(record.body).getUint16(0, true),
        };
        break;
      }
      case RecordCode.PrintHeaders:
      case RecordCode.PrintGridlines:
      case RecordCode.Gridset:
      case RecordCode.HCenter:
      case RecordCode.VCenter: {
        if (record.body.byteLength < 2)
          throw new XlsParseError("Invalid legacy XLS file: truncated print setting");
        const value = dataView(record.body).getUint16(0, true);
        if (value !== 0 && value !== 1)
          throw new XlsParseError("Invalid legacy XLS file: invalid print setting", {
            recordType: record.code,
            recordName: "Print Setting",
            offset: record.start,
            length: record.body.byteLength,
            byteRange: [record.start + 4, record.start + 6],
            reason: "invalid-boolean-value",
          });
        const enabled = value === 1;
        const printOptions = { ...worksheetExtras.printOptions };
        if (record.code === RecordCode.PrintHeaders) printOptions.headings = enabled;
        if (record.code === RecordCode.PrintGridlines) printOptions.gridLines = enabled;
        if (record.code === RecordCode.Gridset) printOptions.gridLinesSet = enabled;
        if (record.code === RecordCode.HCenter) printOptions.horizontalCentered = enabled;
        if (record.code === RecordCode.VCenter) printOptions.verticalCentered = enabled;
        worksheetExtras.printOptions = printOptions;
        break;
      }
      case RecordCode.MergedCells: {
        if (record.body.byteLength < 2) {
          throw new XlsParseError("Invalid legacy XLS file: truncated MERGEDCELLS count", {
            part: "stream",
            path: "worksheet",
            recordType: record.code,
            offset: record.start,
            reason: "truncated-merged-cell-count",
          });
        }
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const count = view.getUint16(0, true);
        const available = Math.floor((record.body.byteLength - 2) / 8);
        if (count !== available) {
          throw new XlsParseError("Invalid legacy XLS file: truncated MERGEDCELLS ranges", {
            part: "stream",
            path: "worksheet",
            recordType: record.code,
            offset: record.start,
            reason: "truncated-merged-cell-ranges",
          });
        }
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
        worksheetExtras.hyperlinks = [
          ...(worksheetExtras.hyperlinks ?? []),
          parseHyperlink(record.body),
        ];
        break;
      }
      case RecordCode.Obj: {
        if (record.body.byteLength < 8) {
          throw new XlsParseError("Invalid legacy XLS file: truncated OBJ", {
            part: "stream",
            recordType: record.code,
            recordName: "OBJ",
            offset: record.start,
            length: 8,
            reason: "invalid-record-length",
          });
        }
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        if (view.getUint16(4, true) === 0x0019) noteObjectId = view.getUint16(6, true);
        break;
      }
      case RecordCode.Txo: {
        if (record.body.byteLength < 18) {
          throw new XlsParseError("Invalid legacy XLS file: truncated TXO", {
            part: "stream",
            recordType: record.code,
            recordName: "TXO",
            offset: record.start,
            length: 18,
            reason: "invalid-record-length",
          });
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
          throw new XlsParseError("Invalid legacy XLS file: truncated NOTE");
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
        const anchor = noteAnchorsByObjectId.get(objectId);
        worksheetExtras.comments = [
          ...(worksheetExtras.comments ?? []),
          {
            cell: reference(row, column),
            author,
            text: noteText,
            visible: (view.getUint16(4, true) & 0x0002) !== 0,
            ...(anchor
              ? { anchor }
              : {
                  anchor: {
                    from: { col: column, row },
                    to: { col: column + 1, row: row + 1 },
                  },
                }),
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
        conditional = { ...parseConditionalFormatRange(record), rules: [] };
        break;
      }
      case RecordCode.ConditionalFormat: {
        if (!conditional)
          throw new XlsParseError(
            "Invalid legacy XLS file: conditional format rule outside CONDFMT",
          );
        conditional.rules.push(parseConditionalRule(record, conditional.rules.length + 1));
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
      case RecordCode.Dv: {
        if (expectedDataValidationCount === undefined)
          throw new XlsParseError("Invalid legacy XLS file: DV record precedes DVAL", {
            ...biffRecordContext(record, "data-validation-before-list"),
          });
        worksheetExtras.dataValidations = [
          ...(worksheetExtras.dataValidations ?? []),
          parseDataValidation(record, state),
        ];
        dataValidationCount += 1;
        break;
      }
      case RecordCode.Dval: {
        if (record.body.byteLength !== 18 || expectedDataValidationCount !== undefined)
          throw new XlsParseError("Invalid legacy XLS file: invalid or duplicate DVAL record", {
            ...biffRecordContext(record, "invalid-data-validation-list"),
          });
        const view = dataView(record.body);
        const xWindow = view.getUint32(2, true);
        const yWindow = view.getUint32(6, true);
        const objectId = view.getInt32(10, true);
        const count = view.getUint32(14, true);
        if (
          xWindow > 65535 ||
          yWindow > 65535 ||
          (objectId !== -1 && (objectId < 1 || objectId > 32767)) ||
          count > 65534
        )
          throw new XlsParseError("Invalid legacy XLS file: invalid DVAL values", {
            ...biffRecordContext(record, "invalid-data-validation-list"),
          });
        expectedDataValidationCount = count;
        worksheetExtras.dataValidationsXWindow = xWindow;
        worksheetExtras.dataValidationsYWindow = yWindow;
        break;
      }
      case RecordCode.Header: {
        worksheetExtras.headerFooter = {
          ...worksheetExtras.headerFooter,
          oddHeader: parseHeaderText(record, state),
        };
        break;
      }
      case RecordCode.Footer: {
        worksheetExtras.headerFooter = {
          ...worksheetExtras.headerFooter,
          oddFooter: parseHeaderText(record, state),
        };
        break;
      }
      case RecordCode.LeftMargin:
      case RecordCode.RightMargin:
      case RecordCode.TopMargin:
      case RecordCode.BottomMargin: {
        worksheetExtras.pageMargins = parsePageMargins(record, worksheetExtras.pageMargins);
        break;
      }
      case RecordCode.PageSetup: {
        const { pageSetup, headerFooterMargins } = parsePageSetup(record);
        worksheetExtras.pageSetup = pageSetup;
        worksheetExtras.pageMargins = { ...worksheetExtras.pageMargins, ...headerFooterMargins };
        break;
      }
      case RecordCode.Scl: {
        if (record.body.byteLength !== 4)
          throw new XlsParseError("Invalid legacy XLS file: invalid SCL length", {
            part: "stream",
            recordType: record.code,
            recordName: "SCL",
            offset: record.start,
            reason: "invalid-record-length",
          });
        const view = dataView(record.body);
        const numerator = view.getInt16(0, true);
        const denominator = view.getInt16(2, true);
        const zoomScale = Math.round((numerator * 100) / denominator);
        if (
          numerator === 0 ||
          denominator === 0 ||
          numerator < 1 ||
          denominator < 1 ||
          numerator * 10 < denominator ||
          numerator > denominator * 4
        )
          throw new XlsParseError("Invalid legacy XLS file: invalid SCL scale", {
            part: "stream",
            recordType: record.code,
            recordName: "SCL",
            offset: record.start,
            reason: "invalid-scale",
          });
        if (worksheetExtras.sheetView?.zoomScale !== undefined)
          throw new XlsParseError("Invalid legacy XLS file: duplicate SCL record", {
            part: "stream",
            recordType: record.code,
            recordName: "SCL",
            offset: record.start,
            reason: "duplicate-record",
          });
        worksheetExtras.sheetView = {
          ...worksheetExtras.sheetView,
          zoomScale,
        };
        break;
      }
      case RecordCode.CodeName: {
        if (state.version !== 8 || worksheetExtras.properties?.codeName !== undefined)
          throw new XlsParseError("Invalid legacy XLS file: misplaced or duplicate CODENAME", {
            part: "stream",
            recordType: record.code,
            recordName: "CODENAME",
            offset: record.start,
            reason: "misplaced-or-duplicate-codename",
          });
        let decoded: { value: string; offset: number };
        try {
          decoded = readShortBiff8String(record.body, 0, encodingForCodepage(state.codepage));
        } catch (error) {
          throw new XlsParseError("Invalid legacy XLS file: malformed worksheet CODENAME", {
            part: "stream",
            recordType: record.code,
            recordName: "CODENAME",
            offset: record.start,
            reason: "malformed-codename",
            cause: error,
          });
        }
        if (decoded.offset !== record.body.byteLength || decoded.value.length > 31)
          throw new XlsParseError("Invalid legacy XLS file: invalid worksheet CODENAME length", {
            part: "stream",
            recordType: record.code,
            recordName: "CODENAME",
            offset: record.start,
            reason: "invalid-codename-length",
          });
        worksheetExtras.properties = {
          ...worksheetExtras.properties,
          codeName: decoded.value,
        };
        break;
      }
      default:
        assertSupportedBiffRecord(record);
        break;
    }
  }

  if (depth !== 0)
    throw new XlsParseError("Invalid legacy XLS file: sheet stream has no Begin Of File");
  return { ...sheet, extras: worksheetExtras };
}
