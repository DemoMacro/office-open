import type { ConditionalFormatRule, NoteAnchorOptions, WorksheetOptions } from "@office-open/xlsx";

import { XlsParseError } from "../errors";
import type { BiffRecord } from "../records";
import { RecordCode, isClearedFilePass, recordsFrom } from "../records";
import { assertSupportedBiffRecord } from "../records/biff-records";
import { escherAnchors, escherDrawings } from "../records/escher";
import { decodeFormula } from "../records/formula";
import {
  BEGIN_OF_FILE_CODES,
  encodingForCodepage,
  readBiff2String,
  readBiff35String,
  readBiff8String,
  readShortBiff8String,
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
import { readCString } from "./workbook-globals";
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
    const bytes = record.body.subarray(offset, offset + length);
    offset += length;
    try {
      return decodeFormula(bytes, {
        baseRow: 0,
        baseColumn: 0,
        arrayData: record.body.subarray(offset),
      }).formula;
    } catch {
      return `#FORMULA_${length}`;
    }
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
      throw new XlsParseError("Encrypted legacy XLS files are not supported");
    if (record.code === RecordCode.EndOfFile) {
      if (depth === 1 && currentDrawing) addDrawings(concatBytes(currentDrawing));
      depth--;
      if (depth === 0) return { ...sheet, extras: worksheetExtras };
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
        let text: { value: string; offset: number };
        try {
          text =
            state.version === 8
              ? readBiff8String(record.body, 6)
              : readBiff35String(record.body, 6);
        } catch {
          throw new XlsParseError("Invalid legacy XLS file: truncated label cell");
        }
        ensureRow(sheet, row).cells.push(makeCell(row, column, style, text.value, state));
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
        let text: { value: string; offset: number };
        try {
          text = readBiff8String(record.body, 6);
        } catch {
          throw new XlsParseError("Invalid legacy XLS file: truncated rich label cell");
        }
        ensureRow(sheet, row).cells.push(makeCell(row, column, style, text.value, state));
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
        worksheetExtras.pageMargins = parsePageMargins(record, worksheetExtras.pageMargins);
        break;
      }
      case RecordCode.PageSetup: {
        const { pageSetup, headerFooterMargins } = parsePageSetup(record);
        worksheetExtras.pageSetup = pageSetup;
        worksheetExtras.pageMargins = { ...worksheetExtras.pageMargins, ...headerFooterMargins };
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
