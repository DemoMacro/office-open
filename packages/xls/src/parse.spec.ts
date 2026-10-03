import type { CellOptions, WorksheetOptions } from "@office-open/xlsx";
import { describe, expect, it } from "vitest";

import { buildSmallContainer } from "./cfb-builder";
import { parseWorkbook } from "./index";

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
    concat([uint16Body([row, column, 1]), version === 8 ? biff8String(value) : biff5String(value)]),
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
  const body = new Uint8Array(22);
  const view = new DataView(body.buffer);
  view.setUint16(0, row, true);
  view.setUint16(2, column, true);
  view.setUint16(4, 5, true);
  body.set(cached, 6);
  return record(0x0006, body);
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
): Uint8Array {
  const versionToken = version === 8 ? BIFF8_TOKEN : BIFF5_TOKEN;
  const globalBegin = beginOfFile(versionToken, 0x0005);
  const sstRecords = [sharedStringTable(sharedStrings), ...sharedStringParts];
  const positions: number[] = [];
  const sheetRecords = sheets.map((item) => boundSheet(version, 0, item.name, item.state));
  let position =
    globalBegin.byteLength +
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
    const worksheet = parseWorkbook(xls(data)).worksheets?.[0];
    if (!worksheet) throw new TypeError("Expected a worksheet");
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

  it("throws clear errors for invalid inputs", () => {
    expect(() => parseWorkbook(new Uint8Array([1, 2, 3]))).toThrow(/Invalid CFB file/);
    expect(() => parseWorkbook(xls(new Uint8Array([1, 2, 3]), "Other"))).toThrow(
      /missing Workbook or Book/,
    );
    expect(() => parseWorkbook(xls(new Uint8Array(10)))).toThrow(/truncated/);
    expect(() => parseWorkbook(xls(concat([endOfFile(), endOfFile()])))).toThrow(
      /missing Begin Of File/,
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
  });
});
