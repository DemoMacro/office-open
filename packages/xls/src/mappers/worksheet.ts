import type { WorksheetOptions } from "@office-open/xlsx";

import { LegacyExcelError } from "../errors";
import { RecordCode, isClearedFilePass, recordsFrom } from "../records";
import { escherPictures } from "../records/escher";
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
