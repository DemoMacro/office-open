import type { CellOptions, WorksheetOptions } from "@office-open/xlsx";
import { describe, expect, it } from "vitest";

import { buildSmallContainer } from "./cfb/test-container";
import { XlsParseError, parseWorkbook } from "./index";
import { parseHyperlink } from "./mappers/workbook-globals";

const BIFF5_TOKEN = 0x0500;
const BIFF8_TOKEN = 0x0600;

function record(code: number, body: Uint8Array = new Uint8Array()): Uint8Array {
  const bytes = new Uint8Array(4 + body.byteLength);
  const view = new DataView(bytes.buffer);
  view.setUint16(0, code, true);
  view.setUint16(2, body.byteLength, true);
  bytes.set(body, 4);
  return bytes;
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const bytes = new Uint8Array(parts.reduce((total, part) => total + part.byteLength, 0));
  let offset = 0;
  for (const part of parts) {
    bytes.set(part, offset);
    offset += part.byteLength;
  }
  return bytes;
}

function uint16Body(values: readonly number[]): Uint8Array {
  const bytes = new Uint8Array(values.length * 2);
  const view = new DataView(bytes.buffer);
  values.forEach((value, index) => view.setUint16(index * 2, value, true));
  return bytes;
}

function biff8String(value: string): Uint8Array {
  const characters = Array.from(value);
  const bytes = new Uint8Array(3 + characters.length * 2);
  const view = new DataView(bytes.buffer);
  view.setUint16(0, characters.length, true);
  bytes[2] = 1;
  characters.forEach((character, index) =>
    view.setUint16(3 + index * 2, character.charCodeAt(0), true),
  );
  return bytes;
}

function biff5String(value: string): Uint8Array {
  const characters = Array.from(value);
  const bytes = new Uint8Array(1 + characters.length);
  bytes[0] = characters.length;
  characters.forEach((character, index) => {
    bytes[index + 1] = character.charCodeAt(0);
  });
  return bytes;
}

function biff5LabelString(value: string): Uint8Array {
  const characters = Array.from(value);
  const bytes = new Uint8Array(2 + characters.length);
  const view = new DataView(bytes.buffer);
  view.setUint16(0, characters.length, true);
  characters.forEach((character, index) => {
    bytes[index + 2] = character.charCodeAt(0);
  });
  return bytes;
}

function beginOfFile(versionToken: number, type = 0x0010): Uint8Array {
  return record(
    0x0809,
    new Uint8Array([versionToken & 0xff, versionToken >> 8, type & 0xff, type >> 8]),
  );
}

function endOfFile(): Uint8Array {
  return record(0x000a);
}

function sheet(
  versionToken: number,
  cellRecords: readonly Uint8Array[],
  nestedSubstream = false,
): Uint8Array {
  return concat([
    beginOfFile(versionToken),
    ...cellRecords,
    ...(nestedSubstream ? [beginOfFile(versionToken), endOfFile()] : []),
    endOfFile(),
  ]);
}

function dimensions(version: 5 | 8, lastRow: number, lastColumn: number): Uint8Array {
  const bytes = new Uint8Array(version === 8 ? 14 : 10);
  const view = new DataView(bytes.buffer);
  if (version === 8) {
    view.setUint32(4, lastRow, true);
    view.setUint16(10, lastColumn, true);
  } else {
    view.setUint16(2, lastRow, true);
    view.setUint16(6, lastColumn, true);
  }
  return record(0x0200, bytes);
}

function rowRecord(
  row: number,
  firstColumn: number,
  lastColumn: number,
  height: number,
  flags: number,
  style = 0,
): Uint8Array {
  const bytes = new Uint8Array(16);
  const view = new DataView(bytes.buffer);
  view.setUint16(0, row, true);
  view.setUint16(2, firstColumn, true);
  view.setUint16(4, lastColumn, true);
  view.setUint16(6, height, true);
  bytes[13] = 1;
  view.setUint16(12, flags, true);
  view.setUint16(14, style, true);
  return record(0x0208, bytes);
}

function numberCell(row: number, column: number, value: number, style = 0): Uint8Array {
  const body = new Uint8Array(14);
  const view = new DataView(body.buffer);
  view.setUint16(0, row, true);
  view.setUint16(2, column, true);
  view.setUint16(4, style, true);
  view.setFloat64(6, value, true);
  return record(0x0203, body);
}

function labelCell(version: 5 | 8, row: number, column: number, value: string): Uint8Array {
  return record(
    0x0204,
    concat([
      uint16Body([row, column, 1]),
      version === 8 ? biff8String(value) : biff5LabelString(value),
    ]),
  );
}

function labelSstCell(row: number, column: number, index: number): Uint8Array {
  const body = new Uint8Array(10);
  const view = new DataView(body.buffer);
  view.setUint16(0, row, true);
  view.setUint16(2, column, true);
  view.setUint16(4, 1, true);
  view.setUint32(6, index, true);
  return record(0x00fd, body);
}

function richLabelCell(row: number, column: number, value: string): Uint8Array {
  return record(0x00d6, concat([uint16Body([row, column, 3]), biff8String(value)]));
}

function rkCell(row: number, column: number, value: number, style = 2): Uint8Array {
  const body = new Uint8Array(10);
  const view = new DataView(body.buffer);
  view.setUint16(0, row, true);
  view.setUint16(2, column, true);
  view.setUint16(4, style, true);
  view.setInt32(6, value, true);
  return record(0x027e, body);
}

function mulRkCell(row: number, firstColumn: number, values: readonly number[]): Uint8Array {
  const body = new Uint8Array(6 + values.length * 6);
  const view = new DataView(body.buffer);
  view.setUint16(0, row, true);
  view.setUint16(2, firstColumn, true);
  values.forEach((value, index) => {
    view.setUint16(4 + index * 6, 3, true);
    view.setInt32(6 + index * 6, value, true);
  });
  view.setUint16(body.byteLength - 2, firstColumn + values.length - 1, true);
  return record(0x00bd, body);
}

function blankCell(row: number, column: number, style = 4): Uint8Array {
  return record(0x0201, uint16Body([row, column, style]));
}

function boolErrCell(row: number, column: number, value: number, isError: number): Uint8Array {
  return record(0x0205, new Uint8Array([row, 0, column, 0, 5, 0, value, isError]));
}

function biff2Record(code: number, body: Uint8Array): Uint8Array {
  const bytes = new Uint8Array(4 + body.byteLength);
  const view = new DataView(bytes.buffer);
  view.setUint16(0, code, true);
  view.setUint16(2, body.byteLength, true);
  bytes.set(body, 4);
  return bytes;
}

function biff2BeginOfFile(): Uint8Array {
  return biff2Record(0x09, uint16Body([0x0200, 0x0010]));
}

function biff2NumberCell(row: number, column: number, value: number): Uint8Array {
  const body = new Uint8Array(15);
  const view = new DataView(body.buffer);
  view.setUint16(0, row, true);
  view.setUint16(2, column, true);
  view.setFloat64(7, value, true);
  return biff2Record(0x03, body);
}

function biff2IntegerCell(row: number, column: number, value: number): Uint8Array {
  const body = new Uint8Array(9);
  const view = new DataView(body.buffer);
  view.setUint16(0, row, true);
  view.setUint16(2, column, true);
  view.setUint16(7, value, true);
  return biff2Record(0x02, body);
}

function biff2LabelCell(row: number, column: number, value: string): Uint8Array {
  return biff2Record(
    0x04,
    concat([uint16Body([row, column]), new Uint8Array(3), biff5LegacyString(value)]),
  );
}

function biff5LegacyString(value: string): Uint8Array {
  const bytes = new Uint8Array(1 + value.length);
  bytes[0] = value.length;
  for (let index = 0; index < value.length; index++) bytes[index + 1] = value.charCodeAt(index);
  return bytes;
}

function biff4BeginOfFile(streamType = 0x0010): Uint8Array {
  const body = new Uint8Array(6);
  const view = new DataView(body.buffer);
  view.setUint16(0, 0x0400, true);
  view.setUint16(2, streamType, true);
  return record(0x0409, body);
}

function biff4NumberCell(row: number, column: number, value: number): Uint8Array {
  return record(0x0203, concat([uint16Body([row, column, 0]), numericResult(value)]));
}

function biff4LabelCell(row: number, column: number, value: string): Uint8Array {
  const length = new Uint8Array(2);
  new DataView(length.buffer).setUint16(0, value.length, true);
  return record(
    0x0204,
    concat([uint16Body([row, column, 0]), length, new TextEncoder().encode(value)]),
  );
}

function mulBlankCell(row: number, firstColumn: number, styles: readonly number[]): Uint8Array {
  const body = new Uint8Array(6 + styles.length * 2);
  const view = new DataView(body.buffer);
  view.setUint16(0, row, true);
  view.setUint16(2, firstColumn, true);
  styles.forEach((style, index) => view.setUint16(4 + index * 2, style, true));
  view.setUint16(body.byteLength - 2, firstColumn + styles.length - 1, true);
  return record(0x00be, body);
}

function formulaCell(row: number, column: number, cached: Uint8Array): Uint8Array {
  return formulaCellTokens(row, column, cached);
}

function formulaCellTokens(
  row: number,
  column: number,
  cached: Uint8Array,
  tokens: Uint8Array = new Uint8Array(),
): Uint8Array {
  const body = new Uint8Array(22 + tokens.byteLength);
  const view = new DataView(body.buffer);
  view.setUint16(0, row, true);
  view.setUint16(2, column, true);
  view.setUint16(4, 5, true);
  body.set(cached, 6);
  view.setUint16(20, tokens.byteLength, true);
  body.set(tokens, 22);
  return record(0x0006, body);
}

function integerToken(value: number): Uint8Array {
  const bytes = new Uint8Array(3);
  bytes[0] = 0x1e;
  new DataView(bytes.buffer).setUint16(1, value, true);
  return bytes;
}

function sharedExpressionToken(row: number, column: number): Uint8Array {
  return concat([new Uint8Array([0x01]), uint16Body([row, column])]);
}

function referenceToken(row: number, column: number): Uint8Array {
  const bytes = new Uint8Array(5);
  bytes[0] = 0x24;
  const view = new DataView(bytes.buffer);
  view.setUint16(1, row, true);
  view.setUint16(3, column | 0x4000, true);
  return bytes;
}

function relativeReferenceToken(rowDelta: number, columnDelta: number): Uint8Array {
  const bytes = new Uint8Array(5);
  bytes[0] = 0x4c;
  const view = new DataView(bytes.buffer);
  view.setUint16(1, rowDelta & 0xffff, true);
  view.setUint16(3, (columnDelta & 0x3fff) | 0xc000, true);
  return bytes;
}

function biff3LabelCell(row: number, column: number, value: string): Uint8Array {
  const text = new TextEncoder().encode(value);
  const length = new Uint8Array(2);
  new DataView(length.buffer).setUint16(0, text.byteLength, true);
  return record(0x0204, concat([uint16Body([row, column, 0]), length, text]));
}

function biff3NumberCell(row: number, column: number, value: number): Uint8Array {
  const body = new Uint8Array(14);
  const view = new DataView(body.buffer);
  view.setUint16(0, row, true);
  view.setUint16(2, column, true);
  view.setUint16(4, 0, true);
  view.setFloat64(6, value, true);
  return record(0x0203, body);
}

function escherHeader(type: number, instance: number, length: number): Uint8Array {
  const bytes = new Uint8Array(8);
  const view = new DataView(bytes.buffer);
  view.setUint16(0, (instance << 4) | 0x000f, true);
  view.setUint16(2, type, true);
  view.setUint32(4, length, true);
  return bytes;
}

function clientAnchor(): Uint8Array {
  const bytes = new Uint8Array(18);
  const view = new DataView(bytes.buffer);
  view.setUint16(0, 0, true);
  view.setInt16(2, 1, true);
  view.setInt16(4, 2, true);
  view.setInt32(6, 0, true);
  view.setInt32(10, 0, true);
  view.setInt16(14, 3, true);
  view.setInt16(16, 4, true);
  return bytes;
}

function msodrawingAnchor(): Uint8Array {
  const anchor = clientAnchor();
  const sp = new Uint8Array(8);
  const spView = new DataView(sp.buffer);
  spView.setInt32(0, 1025, true);
  spView.setUint32(4, 0x0004, true);
  const opt = uint16Body([0x0104, 0, 0]);
  const spContainerBody = concat([
    escherHeader(0xf00a, 2, sp.byteLength),
    sp,
    escherHeader(0xf00b, 3, opt.byteLength),
    opt,
    escherHeader(0xf010, 1, anchor.byteLength),
    anchor,
  ]);
  const spgrBody = concat([
    escherHeader(0xf004, 0x000f, spContainerBody.byteLength),
    spContainerBody,
  ]);
  const dgBody = concat([escherHeader(0xf003, 0x000f, spgrBody.byteLength), spgrBody]);
  const dggBody = concat([escherHeader(0xf002, 0x000f, dgBody.byteLength), dgBody]);
  return record(0x00ec, concat([escherHeader(0xf000, 0x000f, dggBody.byteLength), dggBody]));
}

function msodrawingOrphanProperties(): Uint8Array {
  const opt = uint16Body([0x0104, 0, 0]);
  const dggBody = concat([escherHeader(0xf000, 0x000f, opt.byteLength), opt]);
  return record(0x00ec, dggBody);
}

function msodrawingShape(
  instance: number,
  flags: number,
  shapeId: number,
  properties?: Uint8Array,
): Uint8Array {
  const anchor = clientAnchor();
  const sp = new Uint8Array(8);
  const spView = new DataView(sp.buffer);
  spView.setInt32(0, shapeId, true);
  spView.setUint32(4, flags | 0x0004, true);
  const spContainerBody = concat([
    escherHeader(0xf00a, instance, sp.byteLength),
    sp,
    ...(properties ? [escherHeader(0xf00b, 1, properties.byteLength), properties] : []),
    escherHeader(0xf010, 1, anchor.byteLength),
    anchor,
  ]);
  const spgrBody = concat([
    escherHeader(0xf004, 0x000f, spContainerBody.byteLength),
    spContainerBody,
  ]);
  const dgBody = concat([escherHeader(0xf003, 0x000f, spgrBody.byteLength), spgrBody]);
  const dggBody = concat([escherHeader(0xf002, 0x000f, dgBody.byteLength), dgBody]);
  return record(0x00ec, concat([escherHeader(0xf000, 0x000f, dggBody.byteLength), dggBody]));
}

function msodrawingNote(shapeId: number): Uint8Array {
  const anchor = clientAnchor();
  const sp = new Uint8Array(8);
  const spView = new DataView(sp.buffer);
  spView.setInt32(0, shapeId, true);
  spView.setUint32(4, 0x0004, true);
  const spContainerBody = concat([
    escherHeader(0xf00a, 1, sp.byteLength),
    sp,
    escherHeader(0xf011, 1, 0),
    escherHeader(0xf010, 1, anchor.byteLength),
    anchor,
  ]);
  const spgrBody = concat([
    escherHeader(0xf004, 0x000f, spContainerBody.byteLength),
    spContainerBody,
  ]);
  const dgBody = concat([escherHeader(0xf003, 0x000f, spgrBody.byteLength), spgrBody]);
  const dggBody = concat([escherHeader(0xf002, 0x000f, dgBody.byteLength), dgBody]);
  return record(0x00ec, concat([escherHeader(0xf000, 0x000f, dggBody.byteLength), dggBody]));
}

function msodrawingGroup(): Uint8Array {
  const bseMetadata = new Uint8Array(36);
  const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const bseBody = concat([bseMetadata, escherHeader(0xf01e, 0x046a, png.byteLength), png]);
  const bstoreBody = concat([escherHeader(0xf007, 0x0001, bseBody.byteLength), bseBody]);
  const dggBody = concat([escherHeader(0xf001, 0x000f, bstoreBody.byteLength), bstoreBody]);
  return record(0x00eb, concat([escherHeader(0xf000, 0x000f, dggBody.byteLength), dggBody]));
}

function workbookGlobals(version: 5 | 8, records: readonly Uint8Array[]): readonly Uint8Array[] {
  return version === 8 ? records : [];
}

function formatRecord(id: number, code: string): Uint8Array {
  const prefix = uint16Body([id]);
  return record(0x041e, concat([prefix, biff8String(code)]));
}

function fontRecord(name: string): Uint8Array {
  const body = new Uint8Array(14);
  const view = new DataView(body.buffer);
  view.setUint16(0, 240, true);
  view.setUint16(2, 0x0002, true);
  view.setUint16(4, 10, true);
  view.setUint16(6, 700, true);
  return record(0x0031, concat([body, shortBiff8String(name)]));
}

function xfRecord(fontIndex: number, numberFormatId: number): Uint8Array {
  const body = new Uint8Array(20);
  const view = new DataView(body.buffer);
  view.setUint16(0, fontIndex, true);
  view.setUint16(2, numberFormatId, true);
  view.setUint16(4, 0x0001, true);
  body[9] = 0xfc;
  return record(0x00e0, body);
}

function styleXfRecord(fontIndex: number): Uint8Array {
  const body = new Uint8Array(20);
  const view = new DataView(body.buffer);
  view.setUint16(0, fontIndex, true);
  view.setUint16(4, 0x0005, true);
  body[9] = 0xfc;
  return record(0x00e0, body);
}

function nameRecord(name: string, tokens: Uint8Array): Uint8Array {
  const encodedName = shortBiff8String(name).slice(1);
  const body = new Uint8Array(14);
  const view = new DataView(body.buffer);
  view.setUint16(0, 0, true);
  body[3] = name.length;
  view.setUint16(4, tokens.byteLength, true);
  view.setUint16(8, 0, true);
  return record(0x0018, concat([body, encodedName, tokens]));
}

function supbookRecord(target: string, sheets: readonly string[]): Uint8Array {
  return record(
    0x01ae,
    concat([
      uint16Body([sheets.length, target.length]),
      biff8String(target).slice(2),
      ...sheets.map(shortBiff8String),
    ]),
  );
}

function partialSupbookRecord(target: string, sheet: string): Uint8Array {
  return record(
    0x01ae,
    concat([uint16Body([2, target.length]), biff8String(target).slice(2), shortBiff8String(sheet)]),
  );
}

function externSheetRecord(): Uint8Array {
  return record(0x0017, uint16Body([1, 0, 0, 0]));
}

function externalReferenceToken(externIndex: number, row = 0, column = 0): Uint8Array {
  return concat([new Uint8Array([0x3a]), uint16Body([externIndex, row, column | 0x4000])]);
}

function mergedCellsRecord(): Uint8Array {
  return record(0x00e5, uint16Body([1, 0, 0, 0, 1, 0]));
}

function colInfoRecord(): Uint8Array {
  return record(0x007d, uint16Body([0, 0, 512, 0, 0x0001, 0]));
}

function hyperlinkRecord(): Uint8Array {
  const body = new Uint8Array(32);
  const view = new DataView(body.buffer);
  view.setInt32(28, 0x00000001, true);
  const urlText = "https://example.com/";
  const urlBytes = new Uint8Array((urlText.length + 1) * 2);
  const urlView = new DataView(urlBytes.buffer);
  Array.from(urlText).forEach((character, index) =>
    urlView.setUint16(index * 2, character.charCodeAt(0), true),
  );
  const urlLength = new Uint8Array(4);
  new DataView(urlLength.buffer).setUint32(0, urlBytes.byteLength, true);
  const moniker = new Uint8Array([
    0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0, 0xaa, 0, 0x4b, 0xa9, 0x0b,
  ]);
  return record(0x01b8, concat([body, moniker, urlLength, urlBytes]));
}

function txoRecord(characterCount: number, runByteCount: number): Uint8Array {
  const body = new Uint8Array(18);
  const view = new DataView(body.buffer);
  view.setUint16(10, characterCount, true);
  view.setUint16(12, runByteCount, true);
  return record(0x01b6, body);
}

function noteRecord(row: number, column: number, objectId: number, author: string): Uint8Array {
  return record(0x001c, concat([uint16Body([row, column, 0, objectId]), shortBiff8String(author)]));
}

function objRecord(objectId: number): Uint8Array {
  return record(0x005d, uint16Body([0x0015, 18, 0x0019, objectId, 0]));
}

function sharedFormulaRecord(tokens: Uint8Array, row = 1): Uint8Array {
  const header = new Uint8Array(10);
  const view = new DataView(header.buffer);
  view.setUint16(0, row, true);
  view.setUint16(2, 0, true);
  view.setUint16(8, 1, true);
  return record(0x04bc, concat([header, tokens]));
}

function arrayFormulaRecord(tokens: Uint8Array): Uint8Array {
  return record(0x0221, concat([uint16Body([2, 3, 0, 0]), new Uint8Array(6), tokens]));
}

function condfmtRecord(): Uint8Array {
  return record(0x01b0, uint16Body([1, 0, 0, 0, 0, 1, 1, 1, 0, 1, 0]));
}

function conditionalFormatRecord(tokens: Uint8Array = integerToken(10)): Uint8Array {
  const body = new Uint8Array(12 + tokens.byteLength);
  const view = new DataView(body.buffer);
  body[0] = 2;
  body[1] = 6;
  view.setUint16(2, 3, true);
  view.setUint32(6, 0, true);
  body.set(tokens, 12);
  return record(0x01b1, body);
}

function pageSetupRecords(): readonly Uint8Array[] {
  const setup = new Uint8Array(32);
  const view = new DataView(setup.buffer);
  view.setUint16(0, 9, true);
  view.setUint16(2, 80, true);
  view.setUint16(6, 1, true);
  view.setUint16(8, 1, true);
  view.setUint16(10, 0x0002, true);
  view.setFloat64(16, 0.5, true);
  view.setFloat64(24, 0.5, true);
  return [
    record(0x0014, new Uint8Array([1, 0, 0x48])),
    record(0x0015, new Uint8Array([1, 0, 0x46])),
    record(0x0026, float64Bytes(0.75)),
    record(0x0027, float64Bytes(0.75)),
    record(0x0028, float64Bytes(0.75)),
    record(0x0029, float64Bytes(0.75)),
    record(0x00a1, setup),
  ];
}

function float64Bytes(value: number): Uint8Array {
  const bytes = new Uint8Array(8);
  new DataView(bytes.buffer).setFloat64(0, value, true);
  return bytes;
}

function numericResult(value: number): Uint8Array {
  const bytes = new Uint8Array(8);
  new DataView(bytes.buffer).setFloat64(0, value, true);
  return bytes;
}

function typedResult(type: number, detail = 0): Uint8Array {
  return new Uint8Array([0, 0, 0, 0, 0, 0, type, detail]);
}

function sharedStringTable(values: readonly string[]): Uint8Array {
  const header = new Uint8Array(8);
  const view = new DataView(header.buffer);
  view.setUint32(0, values.length, true);
  view.setUint32(4, values.length, true);
  return record(0x00fc, concat([header, ...values.map(biff8String)]));
}

function shortBiff8String(value: string): Uint8Array {
  const characters = Array.from(value);
  const bytes = new Uint8Array(2 + characters.length * 2);
  bytes[0] = characters.length;
  bytes[1] = 1;
  const view = new DataView(bytes.buffer);
  characters.forEach((character, index) =>
    view.setUint16(2 + index * 2, character.charCodeAt(0), true),
  );
  return bytes;
}

function boundSheet(version: 5 | 8, position: number, name: string, state = 0): Uint8Array {
  const body = new Uint8Array(6);
  const view = new DataView(body.buffer);
  view.setUint32(0, position, true);
  view.setUint16(4, state, true);
  return record(0x0085, concat([body, version === 8 ? shortBiff8String(name) : biff5String(name)]));
}

interface SyntheticSheet {
  name: string;
  state?: number;
  nestedSubstream?: boolean;
  cells: readonly Uint8Array[];
}

function workbook(
  version: 5 | 8,
  sheets: readonly SyntheticSheet[],
  sharedStrings: readonly string[] = [],
  sharedStringParts: readonly Uint8Array[] = [],
  globals: readonly Uint8Array[] = [],
): Uint8Array {
  const versionToken = version === 8 ? BIFF8_TOKEN : BIFF5_TOKEN;
  const globalBegin = beginOfFile(versionToken, 0x0005);
  const sstRecords = [sharedStringTable(sharedStrings), ...sharedStringParts];
  const positions: number[] = [];
  const sheetRecords = sheets.map((item) => boundSheet(version, 0, item.name, item.state));
  const globalsLength = workbookGlobals(version, globals).reduce(
    (total, recordBytes) => total + recordBytes.byteLength,
    0,
  );
  let position =
    globalBegin.byteLength +
    globalsLength +
    sstRecords.reduce((total, recordBytes) => total + recordBytes.byteLength, 0) +
    endOfFile().byteLength +
    sheetRecords.reduce((total, recordBytes) => total + recordBytes.byteLength, 0);
  for (const item of sheets) {
    positions.push(position);
    position += sheet(versionToken, item.cells, item.nestedSubstream).byteLength;
  }
  const positionedSheetRecords = sheets.map((item, index) =>
    boundSheet(version, positions[index]!, item.name, item.state),
  );
  return concat([
    globalBegin,
    ...workbookGlobals(version, globals),
    ...sstRecords,
    ...positionedSheetRecords,
    endOfFile(),
    ...sheets.map((item) => sheet(versionToken, item.cells, item.nestedSubstream)),
  ]);
}

function xls(stream: Uint8Array, path = "Workbook"): Uint8Array {
  return buildSmallContainer({ path, data: stream });
}

function cell(worksheet: WorksheetOptions, reference: string): CellOptions {
  const value = worksheet.rows
    ?.flatMap((row) => row.cells ?? [])
    .find((candidate) => candidate.reference === reference);
  if (!value) throw new Error(`Missing test cell ${reference}`);
  return value;
}

describe("parseWorkbook", () => {
  it("reports structured low-level parse failures", () => {
    const malformedRecord = (code: number, declaredLength: number): Uint8Array => {
      const bytes = new Uint8Array(4);
      const view = new DataView(bytes.buffer);
      view.setUint16(0, code, true);
      view.setUint16(2, declaredLength, true);
      return bytes;
    };
    const impossibleStreamContainer = xls(beginOfFile(BIFF8_TOKEN));
    new DataView(impossibleStreamContainer.buffer).setUint32(512 + 128 + 116, 0xfffffff0, true);
    const actions: readonly (readonly [string, () => unknown])[] = [
      ["truncated-header", () => parseWorkbook(new Uint8Array(256))],
      [
        "bad-record-length",
        () =>
          parseWorkbook(xls(concat([beginOfFile(BIFF8_TOKEN), malformedRecord(0x0085, 0xffff)]))),
      ],
      ["invalid-container-traversal", () => parseWorkbook(impossibleStreamContainer)],
      [
        "unsupported-required-structure",
        () => parseWorkbook(xls(record(0x00ff, new Uint8Array(4)))),
      ],
      [
        "encrypted-unsupported",
        () =>
          parseWorkbook(
            xls(
              concat([
                record(0x002f, new Uint8Array(6)),
                beginOfFile(BIFF8_TOKEN, 0x0005),
                boundSheet(8, 0, "Sheet"),
                endOfFile(),
              ]),
            ),
          ),
      ],
    ];

    for (const [category, action] of actions) {
      let error: XlsParseError | undefined;
      try {
        action();
      } catch (thrown) {
        error = thrown as XlsParseError;
      }
      expect(error, category).toBeInstanceOf(XlsParseError);
      expect(error!.context.format, category).toBe("xls");
      expect(error!.context.reason, category).toMatch(
        /invalid-container|invalid-record-length|missing-required-record|truncated-header|encrypted-unsupported/,
      );
      expect(error!.context.path, category).toBeDefined();
    }

    try {
      parseWorkbook(new Uint8Array(256));
    } catch (error) {
      expect((error as XlsParseError).context).toMatchObject({
        format: "xls",
        part: "container",
        path: "/",
      });
    }
  });

  it("parses a raw BIFF2 worksheet stream", () => {
    const data = concat([
      biff2BeginOfFile(),
      biff2NumberCell(0, 0, 1.5),
      biff2IntegerCell(0, 1, 7),
      biff2LabelCell(1, 0, "BIFF2"),
      biff2Record(0x0a, new Uint8Array()),
    ]);
    const worksheets = parseWorkbook(data).worksheets ?? [];
    expect(worksheets).toHaveLength(1);
    expect(cell(worksheets[0]!, "A1").value).toBe(1.5);
    expect(cell(worksheets[0]!, "B1").value).toBe(7);
    expect(cell(worksheets[0]!, "A2").value).toBe("BIFF2");
  });

  it("parses a raw BIFF4 worksheet stream", () => {
    const data = concat([
      biff4BeginOfFile(),
      biff4NumberCell(0, 0, 2.25),
      biff4LabelCell(0, 1, "BIFF4"),
      endOfFile(),
    ]);
    const worksheets = parseWorkbook(data).worksheets ?? [];
    expect(worksheets).toHaveLength(1);
    expect(cell(worksheets[0]!, "A1").value).toBe(2.25);
    expect(cell(worksheets[0]!, "B1").value).toBe("BIFF4");
  });

  it("parses a raw BIFF3 worksheet stream", () => {
    const data = concat([
      record(0x0209, uint16Body([0x0200, 0x0010])),
      biff3LabelCell(0, 0, "BIFF3"),
      biff3NumberCell(0, 1, 2.5),
      endOfFile(),
    ]);
    const worksheets = parseWorkbook(data).worksheets ?? [];
    expect(worksheets).toHaveLength(1);
    expect(cell(worksheets[0]!, "A1").value).toBe("BIFF3");
    expect(cell(worksheets[0]!, "B1").value).toBe(2.5);
  });

  it("reads BIFF4W workbook streams through SHEETHDR substreams", () => {
    const sheetStream = (value: string) =>
      concat([biff4BeginOfFile(), biff4LabelCell(0, 0, value), endOfFile()]);
    const first = sheetStream("First");
    const second = sheetStream("Second");
    const sheetHeader = (name: string, stream: Uint8Array) =>
      record(
        0x008f,
        concat([
          (() => {
            const bytes = new Uint8Array(4);
            new DataView(bytes.buffer).setInt32(0, stream.byteLength, true);
            return bytes;
          })(),
          biff5LegacyString(name),
        ]),
      );
    const data = concat([
      biff4BeginOfFile(0x0100),
      sheetHeader("One", first),
      first,
      sheetHeader("Two", second),
      second,
      endOfFile(),
    ]);
    const worksheets = parseWorkbook(data).worksheets ?? [];
    expect(worksheets.map((worksheet) => worksheet.name)).toEqual(["One", "Two"]);
    expect(cell(worksheets[0]!, "A1").value).toBe("First");
    expect(cell(worksheets[1]!, "A1").value).toBe("Second");
  });

  it("parses BIFF8 cells, dimensions, and cached formula results", () => {
    const data = workbook(
      8,
      [
        {
          name: "Sheet",
          cells: [
            dimensions(8, 3, 3),
            formulaCell(0, 0, numericResult(123.5)),
            formulaCell(0, 1, typedResult(1, 1)),
            formulaCell(0, 2, typedResult(2, 0x07)),
            formulaCell(0, 3, typedResult(0xff)),
            record(0x0007, biff8String("from formula")),
            numberCell(1, 0, 3.25, 7),
            labelCell(8, 1, 1, "label"),
            labelSstCell(1, 2, 0),
            rkCell(2, 0, 0xaa),
            rkCell(3, 1, 0xffffffe6),
            mulRkCell(2, 1, [0x2a, 0x2e]),
            blankCell(3, 0),
            mulBlankCell(3, 1, [5, 6]),
          ],
        },
      ],
      ["shared"],
    );
    const parsed = parseWorkbook(xls(data));
    const worksheets = parsed.worksheets ?? [];
    const worksheet = worksheets[0]!;
    expect(worksheets).toHaveLength(1);
    expect(worksheet.name).toBe("Sheet");
    expect(worksheet.sheetId).toBe(1);
    expect(worksheet.state).toBe("visible");
    expect(worksheet.dimension).toBe("A1:D4");
    expect(cell(worksheet, "A1").value).toBe(123.5);
    expect(cell(worksheet, "B1").value).toBe(true);
    expect(cell(worksheet, "C1").value).toBe("#DIV/0!");
    expect(cell(worksheet, "D1").value).toBe("from formula");
    expect(cell(worksheet, "A2").value).toBe(3.25);
    expect(cell(worksheet, "B2").value).toBe("label");
    expect(cell(worksheet, "C2").value).toBe("shared");
    expect(cell(worksheet, "A3").value).toBe(42);
    expect(cell(worksheet, "B4").value).toBe(-7);
    expect(cell(worksheet, "B3").value).toBe(10);
    expect(cell(worksheet, "C3").value).toBe(11);
    expect(cell(worksheet, "A4").value).toBeNull();
    expect(cell(worksheet, "C4").value).toBeNull();
    expect(cell(worksheet, "A4").style).toBe(4);
  });

  it("parses BIFF5 records and 16-bit dimensions", () => {
    const data = workbook(5, [
      {
        name: "Legacy",
        cells: [
          dimensions(5, 2, 2),
          numberCell(0, 0, 1.5),
          labelCell(5, 1, 0, "old"),
          rkCell(2, 2, 0xaa),
        ],
      },
    ]);
    const parsed = parseWorkbook(xls(data));
    const worksheets = parsed.worksheets ?? [];
    const worksheet = worksheets[0]!;
    expect(worksheet.name).toBe("Legacy");
    expect(worksheet.dimension).toBe("A1:C3");
    expect(cell(worksheet, "A1").value).toBe(1.5);
    expect(cell(worksheet, "A2").value).toBe("old");
    expect(cell(worksheet, "C3").value).toBe(42);
  });

  it("parses boolean and error cells", () => {
    const data = workbook(8, [
      {
        name: "Sheet",
        cells: [boolErrCell(0, 0, 1, 0), boolErrCell(0, 1, 0x0f, 1), boolErrCell(0, 2, 0x2b, 1)],
      },
    ]);
    const worksheet = parseWorkbook(xls(data)).worksheets?.[0];
    if (!worksheet) throw new TypeError("Expected a worksheet");
    expect(cell(worksheet, "A1").value).toBe(true);
    expect(cell(worksheet, "B1").value).toBe("#VALUE!");
    expect(cell(worksheet, "C1").value).toBe("#GETTING_DATA");
  });

  it("reads BIFF8 rich label cells", () => {
    const data = workbook(8, [{ name: "Sheet", cells: [richLabelCell(0, 0, "rich")] }]);
    const options = parseWorkbook(xls(data));
    const worksheet = options.worksheets?.[0];
    if (!worksheet) throw new TypeError("Expected a worksheet");
    expect(options.cellXfs).toHaveLength(4);
    const richCell = cell(worksheet, "A1");
    expect(richCell.value).toBe("rich");
    expect(richCell.style).toBe(3);
  });

  it("reads row layout semantics", () => {
    const flags = 0x0003 | 0x0010 | 0x0020 | 0x0080 | 0x0100 | 0x0200 | 0x0400;
    const row = rowRecord(4, 2, 5, 450, flags, 123);
    const data = workbook(8, [{ name: "Sheet", cells: [row, numberCell(4, 2, 7)] }]);
    const worksheet = parseWorkbook(xls(data)).worksheets?.[0];
    if (!worksheet) throw new TypeError("Expected a worksheet");
    const parsedRow = worksheet.rows?.[0];
    expect(parsedRow).toMatchObject({
      rowNumber: 5,
      spans: "3:5",
      height: 22.5,
      outlineLevel: 3,
      collapsed: true,
      hidden: true,
      customFormat: true,
      thickTop: true,
      thickBot: true,
      phonetic: true,
      style: 123,
    });
    expect(cell(worksheet, "C5").value).toBe(7);
  });

  it("reads BIFF8 shared strings across continuation records", () => {
    const first = new Uint8Array(17);
    const firstView = new DataView(first.buffer);
    firstView.setUint32(0, 1, true);
    firstView.setUint32(4, 1, true);
    firstView.setUint16(8, 6, true);
    first[10] = 1;
    "ABCDEF"
      .slice(0, 3)
      .split("")
      .forEach((character, index) => {
        firstView.setUint16(11 + index * 2, character.charCodeAt(0), true);
      });
    const continuation = new Uint8Array(7);
    const continuationView = new DataView(continuation.buffer);
    continuation[0] = 1;
    "ABCDEF"
      .slice(3)
      .split("")
      .forEach((character, index) => {
        continuationView.setUint16(1 + index * 2, character.charCodeAt(0), true);
      });
    const data = workbook(
      8,
      [{ name: "Sheet", cells: [labelSstCell(0, 0, 0)] }],
      [],
      [record(0x00fc, first), record(0x003c, continuation)],
    );
    const worksheets = parseWorkbook(xls(data)).worksheets ?? [];
    expect(cell(worksheets[0]!, "A1").value).toBe("ABCDEF");
  });

  it("maps sheet visibility and sorts rows", () => {
    const data = workbook(8, [
      { name: "Hidden", state: 1, cells: [numberCell(1, 0, 2), numberCell(0, 0, 1)] },
      { name: "Very hidden", state: 2, cells: [numberCell(0, 0, 3)] },
      { name: "Visible", cells: [numberCell(0, 0, 4)] },
    ]);
    const parsed = parseWorkbook(xls(data));
    const worksheets = parsed.worksheets ?? [];
    expect(worksheets.map((worksheet) => worksheet.state)).toEqual([
      "hidden",
      "veryHidden",
      "visible",
    ]);
    expect(worksheets[0]!.rows?.map((row) => row.rowNumber)).toEqual([1, 2]);
  });

  it("stops a stale shared-string count at the physical table end", () => {
    const header = new Uint8Array(8);
    const headerView = new DataView(header.buffer);
    headerView.setUint32(0, 2, true);
    headerView.setUint32(4, 2, true);
    const staleSst = record(0x00fc, concat([header, biff8String("only")]));
    const data = workbook(8, [{ name: "Sheet", cells: [labelSstCell(0, 0, 0)] }], [], [staleSst]);
    const worksheets = parseWorkbook(xls(data)).worksheets ?? [];
    expect(cell(worksheets[0]!, "A1").value).toBe("only");
  });

  it("ignores nested substreams before the worksheet EOF", () => {
    const data = workbook(8, [
      { name: "Nested", nestedSubstream: true, cells: [numberCell(0, 0, 7)] },
    ]);
    const worksheets = parseWorkbook(xls(data)).worksheets ?? [];
    expect(worksheets).toHaveLength(1);
    expect(cell(worksheets[0]!, "A1").value).toBe(7);
  });

  it("accepts Book as the stream name", () => {
    const data = workbook(8, [{ name: "Sheet", cells: [numberCell(0, 0, 1)] }]);
    const worksheets = parseWorkbook(xls(data, "Book")).worksheets ?? [];
    expect(worksheets[0]!.name).toBe("Sheet");
  });

  it("decodes CODEPAGE strings and DATEMODE", () => {
    const data = workbook(
      8,
      [{ name: "Sheet", cells: [labelSstCell(0, 0, 0)] }],
      ["café"],
      [],
      [record(0x0042, uint16Body([65001])), record(0x0022, uint16Body([1]))],
    );
    const parsed = parseWorkbook(xls(data));
    expect(cell((parsed.worksheets ?? [])[0]!, "A1").value).toBe("café");
    expect(parsed.properties?.date1904).toBe(true);
  });

  it("maps FONT, FORMAT and XF to cell formatting", () => {
    const data = workbook(
      8,
      [{ name: "Sheet", cells: [numberCell(0, 0, 1, 0)] }],
      [],
      [],
      [fontRecord("Arial"), formatRecord(164, "0.0%"), xfRecord(0, 164)],
    );
    const parsed = parseWorkbook(xls(data));
    expect(parsed.fonts?.[0]).toMatchObject({ font: "Arial", size: 12, bold: true, italic: true });
    expect(parsed.numFmts).toEqual([{ numFmtId: 164, formatCode: "0.0%" }]);
    expect(cell((parsed.worksheets ?? [])[0]!, "A1").style).toEqual({
      font: parsed.fonts?.[0],
      numFmt: "0.0%",
    });
  });

  it("rejects malformed style metadata records structurally", () => {
    expect(() =>
      parseWorkbook(
        xls(
          workbook(8, [{ name: "Sheet", cells: [] }], [], [], [record(0x0031, new Uint8Array(14))]),
        ),
      ),
    ).toThrow(XlsParseError);
    expect(() =>
      parseWorkbook(
        xls(
          workbook(8, [{ name: "Sheet", cells: [] }], [], [], [record(0x041e, uint16Body([164]))]),
        ),
      ),
    ).toThrow(XlsParseError);
  });

  it("uses the implicit default font for built-in style formats", () => {
    const data = workbook(
      8,
      [{ name: "Sheet", cells: [numberCell(0, 0, 1, 1)] }],
      [],
      [],
      [fontRecord("Test"), xfRecord(0, 0), styleXfRecord(1)],
    );
    const parsed = parseWorkbook(xls(data));
    expect(cell((parsed.worksheets ?? [])[0]!, "A1").style).toEqual({
      font: {
        size: 10,
        bold: false,
        italic: false,
        color: { indexed: 0 },
        underline: false,
        font: "Arial",
      },
    });

    const invalid = workbook(
      8,
      [{ name: "Sheet", cells: [numberCell(0, 0, 1, 1)] }],
      [],
      [],
      [fontRecord("Test"), xfRecord(0, 0), styleXfRecord(2)],
    );
    expect(() => parseWorkbook(xls(invalid))).toThrow(/missing font/);
  });

  it("projects the complete XF format and inherits parent style attributes", () => {
    const parentXf = new Uint8Array(20);
    const cellXf = new Uint8Array(20);
    const parentView = new DataView(parentXf.buffer);
    const cellView = new DataView(cellXf.buffer);
    parentView.setUint16(4, 0x0004, true);
    parentView.setUint16(6, 0x0012, true);
    parentXf[9] = 0x8c;
    parentView.setUint32(10, 0x00000001, true);
    parentView.setUint32(12, 0x00000008, true);
    parentView.setUint32(14, 0x04000000, true);
    parentView.setUint16(18, 0x2088, true);
    cellView.setUint16(2, 164, true);
    cellXf[9] = 0x0c;
    const data = workbook(
      8,
      [{ name: "Sheet", cells: [numberCell(0, 0, 1, 1)] }],
      [],
      [],
      [
        fontRecord("Arial"),
        formatRecord(164, "0.0%"),
        record(0x0092, concat([uint16Body([1]), new Uint8Array([255, 0, 0, 255])])),
        record(0x00e0, parentXf),
        record(0x00e0, cellXf),
      ],
    );
    const parsed = parseWorkbook(xls(data));
    expect(cell((parsed.worksheets ?? [])[0]!, "A1").style).toMatchObject({
      numFmt: "0.0%",
      alignment: { horizontal: "center", vertical: "center" },
      border: { left: { style: "thin", color: { rgb: "FF0000ff" } } },
      fill: {
        type: "pattern",
        patternType: "solid",
        foregroundColor: { rgb: "FF0000ff" },
        backgroundColor: { indexed: 65 },
      },
    });
  });

  it("rejects truncated XF records", () => {
    const data = workbook(
      8,
      [{ name: "Sheet", cells: [numberCell(0, 0, 1)] }],
      [],
      [],
      [record(0x00e0, uint16Body([0, 0]))],
    );
    expect(() => parseWorkbook(xls(data))).toThrow(XlsParseError);
  });

  it("parses PALETTE colors", () => {
    const colors = new Uint8Array([255, 0, 0, 255]);
    const data = workbook(
      8,
      [{ name: "Sheet", cells: [] }],
      [],
      [],
      [record(0x0092, concat([uint16Body([1]), colors]))],
    );
    expect(parseWorkbook(xls(data)).colors?.indexedColors?.[0]?.rgb).toBe("FF0000ff");
  });

  it("rejects malformed auxiliary formatting records", () => {
    const data = workbook(
      8,
      [{ name: "Sheet", cells: [numberCell(0, 0, 1)] }],
      [],
      [],
      [
        record(0x0031, new Uint8Array(3)),
        record(0x041e, new Uint8Array(1)),
        record(0x0092, new Uint8Array(1)),
      ],
    );
    expect(() => parseWorkbook(xls(data))).toThrow(XlsParseError);
  });

  it("parses MERGEDCELLS and COLINFO", () => {
    const data = workbook(
      8,
      [{ name: "Sheet", cells: [mergedCellsRecord(), colInfoRecord()] }],
      [],
      [],
      [],
    );
    const worksheet = (parseWorkbook(xls(data)).worksheets ?? [])[0]!;
    expect(worksheet.mergeCells).toEqual([{ ref: "A1:B1" }]);
    expect(worksheet.columns).toEqual([
      {
        min: 1,
        max: 1,
        width: 2,
        customWidth: true,
        hidden: true,
        outlineLevel: 0,
        collapsed: false,
      },
    ]);
  });

  it("parses NAME defined names", () => {
    const data = workbook(
      8,
      [{ name: "Sheet", cells: [] }],
      [],
      [],
      [nameRecord("Total", referenceToken(0, 0))],
    );
    expect(parseWorkbook(xls(data)).definedNames).toEqual([
      {
        name: "Total",
        value: "$A$1",
        localSheetId: undefined,
        hidden: false,
        function: false,
        vbProcedure: false,
      },
    ]);
  });

  it("parses internal and external SUPBOOK references", () => {
    const internalData = workbook(
      8,
      [{ name: "Sheet", cells: [] }],
      [],
      [],
      [
        record(0x01ae, uint16Body([1, 0x0401])),
        externSheetRecord(),
        nameRecord("Link", externalReferenceToken(1)),
      ],
    );
    expect(parseWorkbook(xls(internalData)).definedNames).toEqual([
      expect.objectContaining({ name: "Link", value: "Sheet!$A$1" }),
    ]);
    expect(parseWorkbook(xls(internalData)).externalLinks).toBeUndefined();

    const data = workbook(
      8,
      [{ name: "Sheet", cells: [] }],
      [],
      [],
      [supbookRecord("external.xlsx", ["Data"]), externSheetRecord()],
    );
    expect(parseWorkbook(xls(data)).externalLinks?.[0]?.externalBook).toEqual({
      target: "external.xlsx",
      sheetNames: ["Data"],
    });
  });

  it("parses DDE and OLE SUPBOOK references", () => {
    const dde = workbook(
      8,
      [{ name: "Sheet", cells: [] }],
      [],
      [],
      [supbookRecord("Excel|Topic", ["Item"]), externSheetRecord()],
    );
    expect(parseWorkbook(xls(dde)).externalLinks?.[0]?.ddeLink).toEqual({
      ddeService: "Excel",
      ddeTopic: "Topic",
      ddeItems: [{ name: "Item" }],
    });

    const ole = workbook(
      8,
      [{ name: "Sheet", cells: [] }],
      [],
      [],
      [supbookRecord("package:Word.Document.12", ["Item"]), externSheetRecord()],
    );
    expect(parseWorkbook(xls(ole)).externalLinks?.[0]?.oleLink).toEqual({
      progId: "Word.Document.12",
      oleItems: [{ name: "Item" }],
    });
  });

  it("tolerates truncated SUPBOOK sheet lists", () => {
    const warnings: string[] = [];
    const warn = console.warn;
    console.warn = (message: string) => warnings.push(message);
    try {
      const data = workbook(
        8,
        [{ name: "Sheet", cells: [] }],
        [],
        [],
        [partialSupbookRecord("external.xlsx", "Data"), externSheetRecord()],
      );
      const parsed = parseWorkbook(xls(data));
      expect(parsed.externalLinks?.[0]?.externalBook).toEqual({
        target: "external.xlsx",
        sheetNames: ["Data"],
      });
      expect(warnings.some((message) => message.includes("truncated-sheet-list"))).toBe(true);
    } finally {
      console.warn = warn;
    }

    const missingReference = workbook(
      8,
      [{ name: "Sheet", cells: [] }],
      [],
      [],
      [
        partialSupbookRecord("external.xlsx", "Data"),
        record(0x0017, uint16Body([1, 0, 1, 1])),
        nameRecord("Link", externalReferenceToken(1)),
      ],
    );
    expect(parseWorkbook(xls(missingReference)).definedNames).toEqual([
      expect.objectContaining({ name: "Link", value: "#REF!!$A$1" }),
    ]);
  });

  it("tolerates unknown records but rejects malformed auxiliary structures", () => {
    const warnings: string[] = [];
    const warn = console.warn;
    console.warn = (message: string) => warnings.push(message);
    try {
      const unknown = workbook(
        8,
        [{ name: "Sheet", cells: [] }],
        [],
        [],
        [record(0xffff, new Uint8Array(1))],
      );
      const parsed = parseWorkbook(xls(unknown));
      expect(parsed.worksheets?.[0]?.name).toBe("Sheet");
      expect(warnings.some((message) => message.includes("unknown-record"))).toBe(true);
      expect(warnings.some((message) => message.includes("0xffff"))).toBe(true);

      const malformedSupbook = workbook(
        8,
        [{ name: "Sheet", cells: [] }],
        [],
        [],
        [record(0x01ae, new Uint8Array(1))],
      );
      expect(() => parseWorkbook(xls(malformedSupbook))).toThrow(/truncated SUPBOOK/);
      const malformedName = workbook(
        8,
        [{ name: "Sheet", cells: [] }],
        [],
        [],
        [nameRecord("Broken", new Uint8Array([0xff]))],
      );
      expect(() => parseWorkbook(xls(malformedName))).toThrow(/unsupported formula token 0xff/);
    } finally {
      console.warn = warn;
    }
  });

  it("skips registered structural no-op records", () => {
    const data = workbook(
      8,
      [{ name: "Sheet", cells: [] }],
      [],
      [],
      [record(0x005c, new TextEncoder().encode("office-open"))],
    );
    expect(parseWorkbook(xls(data)).worksheets).toHaveLength(1);
  });

  it("parses HLINK hyperlinks", () => {
    const data = workbook(8, [{ name: "Sheet", cells: [hyperlinkRecord()] }]);
    expect((parseWorkbook(xls(data)).worksheets ?? [])[0]!.hyperlinks?.[0]?.url).toBe(
      "https://example.com/",
    );
  });

  it("stops HLINK text at its unicode terminator", () => {
    const url = "https://example.com/";
    const encoded = new Uint8Array((url.length + 3) * 2);
    const encodedView = new DataView(encoded.buffer);
    Array.from(url).forEach((character, index) =>
      encodedView.setUint16(index * 2, character.charCodeAt(0), true),
    );
    const length = new Uint8Array(4);
    new DataView(length.buffer).setUint32(0, encoded.byteLength, true);
    const moniker = new Uint8Array(16);
    moniker.set([0xe0, 0xc9]);
    const body = concat([new Uint8Array(32), moniker, length, encoded]);
    new DataView(body.buffer).setInt32(28, 0x00000001, true);
    expect(parseHyperlink(body).url).toBe(url);
  });

  it("parses OBJ/TXO/NOTE comments", () => {
    const text = new TextEncoder().encode("note");
    const runs = new Uint8Array(8);
    const data = workbook(8, [
      {
        name: "Sheet",
        cells: [
          objRecord(7),
          txoRecord(4, 8),
          record(0x003c, concat([new Uint8Array([0]), text])),
          record(0x003c, runs),
          noteRecord(2, 3, 7, "Ada"),
        ],
      },
    ]);
    expect((parseWorkbook(xls(data)).worksheets ?? [])[0]!.comments).toEqual([
      {
        cell: "D3",
        author: "Ada",
        text: "note",
        visible: false,
        anchor: { from: { col: 3, row: 2 }, to: { col: 4, row: 3 } },
      },
    ]);
  });

  it("projects NOTE placement from its Escher anchor", () => {
    const text = new TextEncoder().encode("note");
    const runs = new Uint8Array(8);
    const data = workbook(8, [
      {
        name: "Sheet",
        cells: [
          objRecord(7),
          txoRecord(4, 8),
          record(0x003c, concat([new Uint8Array([0]), text])),
          record(0x003c, runs),
          msodrawingNote(7),
          noteRecord(2, 3, 7, "Ada"),
        ],
      },
    ]);
    const worksheet = (parseWorkbook(xls(data)).worksheets ?? [])[0]!;
    expect(worksheet.shapes).toBeUndefined();
    expect(worksheet.comments?.[0]).toMatchObject({
      cell: "D3",
      anchor: {
        from: { col: 2, row: 1, colOff: 0, rowOff: 0 },
        to: { col: 4, row: 3, colOff: 0, rowOff: 0 },
      },
    });
  });

  it("associates multiple TXO notes and decodes UTF-16 continuations", () => {
    const latinText = new Uint8Array([0, ...new TextEncoder().encode("old")]);
    const unicodeText = new Uint8Array([1, 0xe9, 0]);
    const runs = new Uint8Array(8);
    const data = workbook(8, [
      {
        name: "Sheet",
        cells: [
          objRecord(8),
          txoRecord(1, 8),
          record(0x003c, unicodeText),
          record(0x003c, runs),
          objRecord(7),
          txoRecord(3, 8),
          record(0x003c, latinText),
          record(0x003c, runs),
          noteRecord(0, 0, 7, "Ada"),
          noteRecord(1, 1, 8, "Grace"),
        ],
      },
    ]);
    const comments = (parseWorkbook(xls(data)).worksheets ?? [])[0]!.comments ?? [];
    expect(comments.map((comment) => [comment.text, comment.cell])).toEqual([
      ["old", "A1"],
      ["é", "B2"],
    ]);
  });

  it("recursively parses MSODRAWING Escher anchors", () => {
    const data = workbook(
      8,
      [{ name: "Sheet", cells: [msodrawingAnchor(), numberCell(0, 0, 1)] }],
      [],
      [],
      [msodrawingGroup()],
    );
    const image = (parseWorkbook(xls(data)).worksheets ?? [])[0]!.images?.[0];
    expect(image).toMatchObject({
      col: 3,
      row: 2,
      toCol: 5,
      toRow: 4,
      anchorType: "twoCell",
      shapeId: 1025,
    });
    const imageData = image?.data;
    expect(imageData instanceof Uint8Array ? imageData[0] : undefined).toBe(0x89);
  });

  it("parses Escher shapes and connectors separately from pictures", () => {
    const data = workbook(8, [
      {
        name: "Sheet",
        cells: [msodrawingShape(1, 0, 1026), msodrawingShape(32, 0x0100, 1027)],
      },
    ]);
    const worksheet = (parseWorkbook(xls(data)).worksheets ?? [])[0]!;
    expect(worksheet.images).toBeUndefined();
    expect(worksheet.shapes?.[0]).toMatchObject({
      name: "Shape 1",
      shapeId: 1026,
      properties: { geometry: "rect" },
    });
    expect(worksheet.connectors?.[0]).toMatchObject({
      name: "Connector 2",
      shapeId: 1027,
      properties: { geometry: "straightConnector1" },
    });
  });

  it("maps equivalent Escher shape and connector instances", () => {
    const expectedShapes: readonly [number, string][] = [
      [1, "rect"],
      [12, "star5"],
      [13, "rightArrow"],
      [20, "line"],
      [41, "callout1"],
      [58, "star8"],
      [95, "blockArc"],
      [121, "flowChartPunchedCard"],
      [177, "flowChartOffpageConnector"],
      [187, "star4"],
      [190, "actionButtonHome"],
      [200, "actionButtonMovie"],
    ];
    const expectedConnectors: readonly [number, string][] = [
      [32, "straightConnector1"],
      [36, "bentConnector5"],
      [40, "curvedConnector5"],
    ];
    for (const [index, [instance, geometry]] of expectedShapes.entries()) {
      const data = workbook(8, [
        { name: "Sheet", cells: [msodrawingShape(instance, 0, 2000 + index)] },
      ]);
      const worksheet = (parseWorkbook(xls(data)).worksheets ?? [])[0]!;
      expect(worksheet.shapes?.[0]?.properties.geometry).toBe(geometry);
    }
    for (const [index, [instance, geometry]] of expectedConnectors.entries()) {
      const data = workbook(8, [
        { name: "Sheet", cells: [msodrawingShape(instance, 0x0100, 3000 + index)] },
      ]);
      const worksheet = (parseWorkbook(xls(data)).worksheets ?? [])[0]!;
      expect(worksheet.connectors?.[0]?.properties.geometry).toBe(geometry);
    }
  });

  it("rejects unsupported Escher visual instances", () => {
    const data = workbook(8, [{ name: "Sheet", cells: [msodrawingShape(24, 0, 1026)] }]);
    expect(() => parseWorkbook(xls(data))).toThrow(XlsParseError);
  });

  it("projects Escher text-box and picture-frame instances canonically", () => {
    const textBox = workbook(8, [{ name: "Sheet", cells: [msodrawingShape(202, 0, 1026)] }]);
    const textBoxSheet = (parseWorkbook(xls(textBox)).worksheets ?? [])[0]!;
    expect(textBoxSheet.shapes?.[0]).toMatchObject({
      name: "Text Box 1",
      shapeId: 1026,
      textBox: true,
      properties: { geometry: "rect" },
    });

    const picture = workbook(
      8,
      [
        {
          name: "Sheet",
          cells: [msodrawingShape(75, 0, 1027, uint16Body([0x4104, 0, 0]))],
        },
      ],
      [],
      [],
      [msodrawingGroup()],
    );
    const pictureSheet = (parseWorkbook(xls(picture)).worksheets ?? [])[0]!;
    expect(pictureSheet.images?.[0]).toMatchObject({ shapeId: 1027, type: "png" });
    expect(pictureSheet.shapes).toBeUndefined();

    const missingPicture = workbook(8, [{ name: "Sheet", cells: [msodrawingShape(75, 0, 1027)] }]);
    expect(() => parseWorkbook(xls(missingPicture))).toThrow(/missing Escher picture stream/);
  });

  it("accepts document-level Escher properties", () => {
    const data = workbook(8, [
      { name: "Sheet", cells: [msodrawingOrphanProperties(), numberCell(0, 0, 1)] },
    ]);
    const worksheet = (parseWorkbook(xls(data)).worksheets ?? [])[0]!;
    expect(worksheet.images).toBeUndefined();
    expect(worksheet.shapes).toBeUndefined();
    expect(worksheet.connectors).toBeUndefined();
    expect(cell(worksheet, "A1").value).toBe(1);
  });

  it("parses CONDFMT and CF rules", () => {
    const data = workbook(8, [
      { name: "Sheet", cells: [condfmtRecord(), conditionalFormatRecord()] },
    ]);
    expect((parseWorkbook(xls(data)).worksheets ?? [])[0]!.conditionalFormats).toEqual([
      {
        sqref: "A1:B1",
        rules: [{ type: "cellIs", operator: "lessThan", formulas: ["10"], priority: 1 }],
      },
    ]);
  });

  it("rejects malformed conditional format formulas structurally", () => {
    const data = workbook(8, [
      { name: "Sheet", cells: [condfmtRecord(), conditionalFormatRecord(new Uint8Array([0xff]))] },
    ]);
    expect(() => parseWorkbook(xls(data))).toThrow(XlsParseError);
  });

  it("parses HEADER, FOOTER, margins and PAGESETUP", () => {
    const data = workbook(8, [{ name: "Sheet", cells: [...pageSetupRecords()] }]);
    const worksheet = (parseWorkbook(xls(data)).worksheets ?? [])[0]!;
    expect(worksheet.headerFooter).toEqual({ oddHeader: "H", oddFooter: "F" });
    expect(worksheet.pageMargins).toEqual({
      left: 0.75,
      right: 0.75,
      top: 0.75,
      bottom: 0.75,
      header: 0.5,
      footer: 0.5,
    });
    expect(worksheet.pageSetup).toMatchObject({ paperSize: 9, scale: 80, orientation: "portrait" });
  });

  it("rejects truncated HEADER and FOOTER strings", () => {
    for (const code of [0x0014, 0x0015]) {
      const data = workbook(8, [
        { name: "Sheet", cells: [record(code, new Uint8Array([4, 0, 0x26]))] },
      ]);
      expect(() => parseWorkbook(xls(data))).toThrow(XlsParseError);
    }
  });

  it("parses BIFF5 HEADER and FOOTER strings without BIFF8 flags", () => {
    const data = workbook(5, [
      {
        name: "Sheet",
        cells: [
          record(0x0014, new Uint8Array([1, 0x48])),
          record(0x0015, new Uint8Array([1, 0x46])),
        ],
      },
    ]);
    const worksheet = (parseWorkbook(xls(data)).worksheets ?? [])[0]!;
    expect(worksheet.headerFooter).toEqual({ oddHeader: "H", oddFooter: "F" });
  });

  it("parses normal, shared and array formula tokens", () => {
    const data = workbook(8, [
      {
        name: "Sheet",
        cells: [
          formulaCellTokens(0, 0, numericResult(42), integerToken(42)),
          formulaCellTokens(1, 0, typedResult(0xff)),
          sharedFormulaRecord(integerToken(21), 1),
          formulaCellTokens(2, 0, typedResult(0xff)),
          arrayFormulaRecord(integerToken(7)),
        ],
      },
    ]);
    const worksheet = (parseWorkbook(xls(data)).worksheets ?? [])[0]!;
    expect(cell(worksheet, "A1").formula).toEqual({ formula: "42" });
    expect(cell(worksheet, "A2").formula).toMatchObject({ type: "shared", sharedIndex: 0 });
    expect(cell(worksheet, "A2").formula).toMatchObject({ formula: "21" });
    expect(cell(worksheet, "A3").formula).toEqual({
      formula: "7",
      type: "array",
      reference: "A3:A4",
    });
  });

  it("parses formula cells that use variable function tokens", () => {
    const data = workbook(8, [
      {
        name: "Sheet",
        cells: [
          formulaCellTokens(
            0,
            0,
            numericResult(42),
            Uint8Array.from([0x44, 0, 0, 0, 0x40, 0x42, 1, 4, 0]),
          ),
        ],
      },
    ]);
    expect(cell((parseWorkbook(xls(data)).worksheets ?? [])[0]!, "A1").formula).toEqual({
      formula: "SUM($A$1)",
    });
  });

  it("rejects malformed formula records structurally", () => {
    const malformed = workbook(8, [
      { name: "Sheet", cells: [formulaCellTokens(0, 0, numericResult(1), new Uint8Array([0xff]))] },
    ]);
    try {
      parseWorkbook(xls(malformed));
      throw new Error("expected malformed FORMULA to throw");
    } catch (error) {
      expect(error).toBeInstanceOf(XlsParseError);
      expect((error as XlsParseError).context).toMatchObject({
        part: "stream",
        recordType: 0x0006,
        recordName: "FORMULA",
        reason: "malformed-formula",
      });
    }
  });

  it("decodes shared expression and class-qualified list tokens", () => {
    const shared = workbook(8, [
      {
        name: "Sheet",
        cells: [
          formulaCellTokens(0, 0, typedResult(0xff), sharedExpressionToken(0, 0)),
          sharedFormulaRecord(integerToken(21), 0),
        ],
      },
    ]);
    expect(cell(parseWorkbook(xls(shared)).worksheets![0]!, "A1").formula).toMatchObject({
      formula: "21",
      type: "shared",
    });

    const list = workbook(8, [
      {
        name: "Sheet",
        cells: [
          formulaCellTokens(
            0,
            0,
            numericResult(0),
            concat([integerToken(1), integerToken(2), new Uint8Array([0x50])]),
          ),
        ],
      },
    ]);
    expect(cell(parseWorkbook(xls(list)).worksheets![0]!, "A1").formula).toEqual({
      formula: "1,2",
    });
  });

  it("applies base column deltas to relative formula references", () => {
    const data = workbook(8, [
      {
        name: "Sheet",
        cells: [
          formulaCellTokens(0, 1, typedResult(1, 1), relativeReferenceToken(-1, -1)),
          record(0x0007, biff8String("cached")),
        ],
      },
    ]);
    expect(cell((parseWorkbook(xls(data)).worksheets ?? [])[0]!, "B1").formula).toEqual({
      formula: "$A$1",
    });
  });

  it("throws clear errors for invalid inputs", () => {
    expect(() => parseWorkbook(new Uint8Array([1, 2, 3]))).toThrow(/Invalid CFB file/);
    expect(() => parseWorkbook(xls(new Uint8Array([1, 2, 3]), "Other"))).toThrow(
      /missing Workbook or Book/,
    );
    expect(() => parseWorkbook(xls(new Uint8Array(10)))).toThrow(/truncated/);
    expect(() => parseWorkbook(xls(concat([endOfFile(), endOfFile()])))).toThrow(
      /no sheet definitions/,
    );
    expect(() => parseWorkbook(xls(beginOfFile(0x0300)))).toThrow(
      /Unsupported legacy XLS BIFF version/,
    );
    expect(() =>
      parseWorkbook(xls(concat([beginOfFile(BIFF8_TOKEN, 0x0005), endOfFile()]))),
    ).toThrow(/no sheet definitions/);
    const encrypted = concat([
      record(0x002f, new Uint8Array(6)),
      beginOfFile(BIFF8_TOKEN, 0x0005),
      boundSheet(8, 0, "Sheet"),
      endOfFile(),
    ]);
    expect(() => parseWorkbook(xls(encrypted))).toThrow(/Encrypted legacy XLS/);

    const actions = [
      () => parseWorkbook(new Uint8Array([1, 2, 3])),
      () => parseWorkbook(xls(new Uint8Array([1, 2, 3]), "Other")),
      () => parseWorkbook(xls(new Uint8Array(10))),
      () => parseWorkbook(xls(encrypted)),
    ];
    const errors = actions.map((action) => {
      try {
        action();
      } catch (error) {
        return error;
      }
      throw new Error("Expected parseWorkbook to throw");
    });
    for (const error of errors) expect(error).toBeInstanceOf(XlsParseError);
  });
});
