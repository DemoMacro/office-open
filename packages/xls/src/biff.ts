import { CompoundFileReader } from "@office-open/core";
import type { CellOptions, RowOptions, WorkbookOptions, WorksheetOptions } from "@office-open/xlsx";

const RECORD_LENGTH_SIZE = 4;
const enum RecordCode {
  Formula = 0x0006,
  String = 0x0007,
  FilePass = 0x002f,
  Continue = 0x003c,
  EndOfFile = 0x000a,
  Blank = 0x0201,
  BoolErr = 0x0205,
  Dimensions = 0x0200,
  Row = 0x0208,
  Number = 0x0203,
  Label = 0x0204,
  BoundSheet = 0x0085,
  MulBlank = 0x00be,
  MulRk = 0x00bd,
  Rk = 0x027e,
  LabelSst = 0x00fd,
  SharedStringTable = 0x00fc,
  Xf = 0x00e0,
  RichLabel = 0x00d6,
}

interface BiffRecord {
  readonly code: number;
  readonly start: number;
  readonly body: Uint8Array;
}

interface BoundSheet {
  readonly position: number;
  readonly state: WorksheetOptions["state"];
  readonly name: string;
}

interface SheetState {
  rows: Map<number, RowOptions>;
  dimension?: string;
}

interface WorkbookState {
  version: 5 | 8;
  stream: Uint8Array;
  sharedStrings: string[];
  sheets: readonly BoundSheet[];
}

function requireBytes(data: Uint8Array, length: number, description: string): void {
  if (length < 0 || length > data.byteLength) {
    throw new Error(`Invalid legacy XLS file: truncated ${description}`);
  }
}

function readRecord(stream: Uint8Array, position: number): BiffRecord {
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  if (position < 0 || position + RECORD_LENGTH_SIZE > stream.byteLength) {
    throw new Error("Invalid legacy XLS file: truncated BIFF record header");
  }
  const code = view.getUint16(position, true);
  const length = view.getUint16(position + 2, true);
  const bodyStart = position + RECORD_LENGTH_SIZE;
  if (bodyStart + length > stream.byteLength) {
    throw new Error(
      `Invalid legacy XLS file: truncated BIFF record 0x${code.toString(16).padStart(4, "0")}`,
    );
  }
  return {
    code,
    start: position,
    body: stream.slice(bodyStart, bodyStart + length),
  };
}

function* recordsFrom(
  data: Uint8Array,
  cursor: number,
  stopAtFirstEndOfFile = true,
): Generator<BiffRecord> {
  while (cursor < data.byteLength) {
    const record = readRecord(data, cursor);
    cursor += RECORD_LENGTH_SIZE + record.body.byteLength;
    yield record;
    if (record.code === RecordCode.EndOfFile && stopAtFirstEndOfFile) return;
  }
  throw new Error("Invalid legacy XLS file: workbook stream ends before End Of File");
}

function decodeBiffString(
  body: Uint8Array,
  view: DataView,
  offset: number,
  characterCount: number,
  flags: number,
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
    : new TextDecoder("windows-1252").decode(raw);
  cursor += byteLength + richRunCount * 4;
  if (hasExtension) {
    if (cursor + 4 > body.byteLength)
      throw new Error("Invalid legacy XLS file: truncated rich string");
    cursor += view.getUint32(cursor, true);
  }
  return { value, offset: cursor };
}

function readBiff8String(body: Uint8Array, offset = 0): { value: string; offset: number } {
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  requireBytes(body, offset + 3, "string");
  const characterCount = view.getUint16(offset, true);
  const flags = body[offset + 2]!;
  return decodeBiffString(body, view, offset + 3, characterCount, flags);
}

function readBiff5String(body: Uint8Array, offset = 0): { value: string; offset: number } {
  requireBytes(body, offset + 1, "string");
  const characterCount = body[offset]!;
  requireBytes(body, offset + 1 + characterCount, "string");
  const raw = body.subarray(offset + 1, offset + 1 + characterCount);
  return {
    value: new TextDecoder("windows-1252").decode(raw),
    offset: offset + 1 + characterCount,
  };
}

function readShortBiff8String(body: Uint8Array, offset = 0): { value: string; offset: number } {
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  requireBytes(body, offset + 2, "string");
  return decodeBiffString(body, view, offset + 2, body[offset]!, body[offset + 1]!);
}

function isVersionSupported(version: number): version is 5 | 8 {
  return version === 0x0500 || version === 0x0600;
}

function supportedVersion(version: number): 5 | 8 {
  return version === 0x0600 ? 8 : 5;
}

function readWorkbookGlobals(
  stream: Uint8Array,
): Pick<WorkbookState, "version" | "sharedStrings" | "sheets"> {
  let workbookVersion: 5 | 8 | undefined;
  const sharedStrings: string[] = [];
  const sheets: BoundSheet[] = [];
  let sharedStringParts: Uint8Array[] | undefined;
  let inGlobals = false;

  for (const record of recordsFrom(stream, 0)) {
    if (record.code === RecordCode.FilePass)
      throw new Error("Encrypted legacy XLS files are not supported");
    if (record.code === RecordCode.Formula || record.code === RecordCode.String) {
      throw new Error("Invalid legacy XLS file: worksheet record appears in workbook globals");
    }
    if (record.code === 0x0809) {
      if (inGlobals) throw new Error("Invalid legacy XLS file: nested workbook globals");
      if (record.body.byteLength < 4)
        throw new Error("Invalid legacy XLS file: truncated Begin Of File");
      const recordVersion = new DataView(
        record.body.buffer,
        record.body.byteOffset,
        record.body.byteLength,
      ).getUint16(0, true);
      if (!isVersionSupported(recordVersion)) {
        throw new Error(`Unsupported legacy XLS BIFF version: ${recordVersion}`);
      }
      workbookVersion = supportedVersion(recordVersion);
      inGlobals = true;
      continue;
    }
    if (!inGlobals) continue;

    if (sharedStringParts && record.code !== RecordCode.Continue) {
      sharedStrings.push(...parseSharedStrings(sharedStringParts, workbookVersion === 8));
      sharedStringParts = undefined;
    }
    switch (record.code) {
      case RecordCode.SharedStringTable: {
        if (record.body.byteLength < 8)
          throw new Error("Invalid legacy XLS file: truncated shared string table");
        sharedStringParts = [record.body];
        break;
      }
      case RecordCode.Continue: {
        sharedStringParts?.push(record.body);
        break;
      }
      case RecordCode.BoundSheet: {
        if (record.body.byteLength < 6)
          throw new Error("Invalid legacy XLS file: truncated sheet definition");
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
              : readBiff5String(record.body, 6).value;
        } catch {
          throw new Error("Invalid legacy XLS file: truncated sheet name");
        }
        sheets.push({ position, state, name });
        break;
      }
      default:
        break;
    }
  }

  if (sharedStringParts)
    sharedStrings.push(...parseSharedStrings(sharedStringParts, workbookVersion === 8));

  if (workbookVersion === undefined)
    throw new Error("Invalid legacy XLS file: missing Begin Of File");
  if (sheets.length === 0)
    throw new Error("Invalid legacy XLS file: workbook has no sheet definitions");
  return { version: workbookVersion, sharedStrings, sheets };
}

function parseSharedStrings(parts: readonly Uint8Array[], isBiff8: boolean): string[] {
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
      throw new Error("Invalid legacy XLS file: truncated shared string table");
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
    let value = "";
    for (let characterIndex = 0; characterIndex < characterCount; characterIndex++) {
      beginCharacter();
      const firstByte = takeByte();
      const secondByte = highByte ? takeByte() : 0;
      value += String.fromCharCode(highByte ? firstByte | (secondByte << 8) : firstByte);
    }
    skipBytes(richRunCount * 4 + extensionLength);
    strings.push(value);
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

function makeCell(
  row: number,
  column: number,
  style: number,
  value: CellOptions["value"],
): CellOptions {
  return { reference: reference(row, column), value, style };
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
    throw new Error(`Invalid legacy XLS file: unknown formula error code 0x${code.toString(16)}`);
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

interface FormulaState {
  row: number;
  column: number;
  style: number;
}

function parseWorksheetStream(
  stream: Uint8Array,
  position: number,
  state: WorkbookState,
): SheetState {
  const sheet: SheetState = { rows: new Map() };
  let formula: FormulaState | undefined;
  let depth = 0;

  for (const record of recordsFrom(stream, position, false)) {
    if (record.code === 0x0809) {
      depth++;
      continue;
    }
    if (depth === 0) continue;
    if (record.code === RecordCode.EndOfFile) {
      depth--;
      if (depth === 0) return sheet;
      continue;
    }
    if (depth !== 1) continue;

    if (record.code === RecordCode.FilePass)
      throw new Error("Encrypted legacy XLS files are not supported");
    if (record.code === RecordCode.String && formula) {
      let text: string;
      try {
        text =
          state.version === 8
            ? readBiff8String(record.body).value
            : readBiff5String(record.body).value;
      } catch {
        throw new Error("Invalid legacy XLS file: truncated formula string");
      }
      ensureRow(sheet, formula.row).cells!.push({
        ...makeCell(formula.row, formula.column, formula.style, undefined),
        value: text,
      });
      formula = undefined;
      continue;
    }
    formula = undefined;

    switch (record.code) {
      case RecordCode.Dimensions: {
        if (state.version === 8 ? record.body.byteLength < 14 : record.body.byteLength < 10) {
          throw new Error("Invalid legacy XLS file: truncated dimensions");
        }
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const firstRow = state.version === 8 ? view.getUint32(0, true) : view.getUint16(0, true);
        const lastRow = state.version === 8 ? view.getUint32(4, true) : view.getUint16(2, true);
        const firstColumn = state.version === 8 ? view.getUint16(8, true) : view.getUint16(4, true);
        const lastColumn = state.version === 8 ? view.getUint16(10, true) : view.getUint16(6, true);
        sheet.dimension = `${reference(firstRow, firstColumn)}:${reference(lastRow, lastColumn)}`;
        break;
      }
      case RecordCode.Row: {
        if (record.body.byteLength < 16) throw new Error("Invalid legacy XLS file: truncated row");
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
      case RecordCode.Blank: {
        if (record.body.byteLength < 6)
          throw new Error("Invalid legacy XLS file: truncated blank cell");
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
      case RecordCode.BoolErr: {
        if (record.body.byteLength < 8)
          throw new Error("Invalid legacy XLS file: truncated boolean/error cell");
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
          makeCell(row, column, style, isError ? errorLiteral(value) : value !== 0),
        );
        break;
      }
      case RecordCode.MulBlank: {
        if (record.body.byteLength < 10)
          throw new Error("Invalid legacy XLS file: truncated multiple blank cells");
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
          throw new Error("Invalid legacy XLS file: multiple blank cell range is truncated");
        }
        const rowOptions = ensureRow(sheet, row);
        for (let index = 0; index < count; index++) {
          const column = firstColumn + index;
          const style = view.getUint16(4 + index * 2, true);
          rowOptions.cells!.push({ reference: reference(row, column), value: null, style });
        }
        break;
      }
      case RecordCode.Number: {
        if (record.body.byteLength < 14)
          throw new Error("Invalid legacy XLS file: truncated number cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        ensureRow(sheet, row).cells!.push(
          makeCell(row, column, view.getUint16(4, true), view.getFloat64(6, true)),
        );
        break;
      }
      case RecordCode.Label: {
        if (record.body.byteLength < 7)
          throw new Error("Invalid legacy XLS file: truncated label cell");
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
            state.version === 8 ? readBiff8String(record.body, 6) : readBiff5String(record.body, 6);
        } catch {
          throw new Error("Invalid legacy XLS file: truncated label cell");
        }
        ensureRow(sheet, row).cells!.push(makeCell(row, column, style, text.value));
        break;
      }
      case RecordCode.RichLabel: {
        if (record.body.byteLength < 7)
          throw new Error("Invalid legacy XLS file: truncated rich label cell");
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
          throw new Error("Invalid legacy XLS file: truncated rich label cell");
        }
        ensureRow(sheet, row).cells!.push(makeCell(row, column, style, text.value));
        break;
      }
      case RecordCode.LabelSst: {
        if (record.body.byteLength < 10)
          throw new Error("Invalid legacy XLS file: truncated shared string cell");
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
          throw new Error(`Invalid legacy XLS file: missing shared string ${index}`);
        ensureRow(sheet, row).cells!.push(makeCell(row, column, style, value));
        break;
      }
      case RecordCode.Rk: {
        if (record.body.byteLength < 10)
          throw new Error("Invalid legacy XLS file: truncated RK cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        ensureRow(sheet, row).cells!.push(
          makeCell(row, column, view.getUint16(4, true), decodeRk(view.getInt32(6, true))),
        );
        break;
      }
      case RecordCode.MulRk: {
        if (record.body.byteLength < 10)
          throw new Error("Invalid legacy XLS file: truncated multiple RK cells");
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
          throw new Error("Invalid legacy XLS file: multiple RK cell range is truncated");
        }
        const rowOptions = ensureRow(sheet, row);
        for (let index = 0; index < count; index++) {
          const offset = 4 + index * 6;
          const column = firstColumn + index;
          const style = view.getUint16(offset, true);
          const value = decodeRk(view.getInt32(offset + 2, true));
          rowOptions.cells!.push(makeCell(row, column, style, value));
        }
        break;
      }
      case RecordCode.Formula: {
        if (record.body.byteLength < 22)
          throw new Error("Invalid legacy XLS file: truncated formula cell");
        const view = new DataView(
          record.body.buffer,
          record.body.byteOffset,
          record.body.byteLength,
        );
        const row = view.getUint16(0, true);
        const column = view.getUint16(2, true);
        const cached = record.body.subarray(6, 14);
        const result = cachedFormulaResult(cached);
        if (result === undefined) {
          formula = { row, column, style: view.getUint16(4, true) };
          continue;
        }
        ensureRow(sheet, row).cells!.push(makeCell(row, column, view.getUint16(4, true), result));
        break;
      }
      default:
        break;
    }
  }

  if (depth !== 0) throw new Error("Invalid legacy XLS file: sheet stream has no Begin Of File");
  return sheet;
}

function readWorkbookStream(data: Uint8Array): WorkbookState {
  const reader = new CompoundFileReader(data);
  const entry = reader.entry("Workbook") ?? reader.entry("Book");
  if (!entry || entry.type !== "stream") {
    throw new Error("Invalid legacy XLS file: missing Workbook or Book stream");
  }
  const stream = reader.read(entry.path);
  if (stream.byteLength < 8)
    throw new Error("Invalid legacy XLS file: workbook stream is truncated");
  const globals = readWorkbookGlobals(stream);
  return { ...globals, stream };
}

export function parseWorkbook(data: Uint8Array): WorkbookOptions {
  if (!(data instanceof Uint8Array)) throw new TypeError("XLS data must be a Uint8Array");
  const state = readWorkbookStream(data);
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
    };
  });
  return { worksheets };
}
