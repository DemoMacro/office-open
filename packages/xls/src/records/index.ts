import { LegacyExcelError } from "../errors";

export const BIFF2_RECORD_HEADER_SIZE = 4;
export const RECORD_LENGTH_SIZE = 4;

export const enum RecordCode {
  Biff2Blank = 0x0001,
  Biff2Integer = 0x0002,
  Biff2Number = 0x0003,
  Biff2Label = 0x0004,
  Biff2BoolErr = 0x0005,
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
  Sheethdr = 0x008f,
  Sheetsoffset = 0x008e,
  MulBlank = 0x00be,
  MulRk = 0x00bd,
  Rk = 0x027e,
  LabelSst = 0x00fd,
  SharedStringTable = 0x00fc,
  Xf = 0x00e0,
  RichLabel = 0x00d6,
  Name = 0x0018,
  ExternSheet = 0x0017,
  Supbook = 0x01ae,
  Font = 0x0031,
  Format = 0x041e,
  Codepage = 0x0042,
  DateMode = 0x0022,
  Palette = 0x0092,
  ColInfo = 0x007d,
  MergedCells = 0x00e5,
  Hyperlink = 0x01b8,
  Note = 0x001c,
  Txo = 0x01b6,
  MsoDrawingGroup = 0x00eb,
  MsoDrawing = 0x00ec,
  Obj = 0x005d,
  CondFmt = 0x01b0,
  ConditionalFormat = 0x01b1,
  Header = 0x0014,
  Footer = 0x0015,
  LeftMargin = 0x0026,
  RightMargin = 0x0027,
  TopMargin = 0x0028,
  BottomMargin = 0x0029,
  PageSetup = 0x00a1,
  SharedFormula = 0x04bc,
  ArrayFormula = 0x0221,
}

export interface BiffRecord {
  readonly code: number;
  readonly start: number;
  readonly body: Uint8Array;
}

export type BiffVersion = 2 | 3 | 4 | 5 | 8;

export function requireBytes(data: Uint8Array, length: number, description: string): void {
  if (length < 0 || length > data.byteLength) {
    throw new LegacyExcelError(`Invalid legacy XLS file: truncated ${description}`);
  }
}

export function recordHeaderSize(version: BiffVersion): number {
  return version === 2 ? BIFF2_RECORD_HEADER_SIZE : RECORD_LENGTH_SIZE;
}

export function readRecord(stream: Uint8Array, position: number, version: BiffVersion): BiffRecord {
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  const headerSize = recordHeaderSize(version);
  if (position < 0 || position + headerSize > stream.byteLength) {
    throw new LegacyExcelError("Invalid legacy XLS file: truncated BIFF record header");
  }
  const code = view.getUint16(position, true);
  const length = view.getUint16(position + 2, true);
  const bodyStart = position + headerSize;
  if (bodyStart + length > stream.byteLength) {
    throw new LegacyExcelError(
      `Invalid legacy XLS file: truncated BIFF record 0x${code.toString(16).padStart(4, "0")}`,
    );
  }
  return { code, start: position, body: stream.slice(bodyStart, bodyStart + length) };
}

export function isClearedFilePass(record: BiffRecord): boolean {
  return record.code === RecordCode.FilePass && record.body.every((byte) => byte === 0);
}

export function* recordsFrom(
  data: Uint8Array,
  cursor: number,
  version: BiffVersion,
  stopAtFirstEndOfFile = true,
): Generator<BiffRecord> {
  while (cursor < data.byteLength) {
    const record = readRecord(data, cursor, version);
    cursor += recordHeaderSize(version) + record.body.byteLength;
    yield record;
    if (record.code === RecordCode.EndOfFile && stopAtFirstEndOfFile) return;
  }
  throw new LegacyExcelError("Invalid legacy XLS file: workbook stream ends before End Of File");
}
