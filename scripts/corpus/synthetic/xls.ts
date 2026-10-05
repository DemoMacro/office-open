import { parseWorkbook as parseLegacyWorkbook } from "../../../packages/xls/dist/index.mjs";
import { generateWorkbook, parseWorkbook } from "../../../packages/xlsx/dist/index.mjs";
import { assertEqual } from "./support";

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

function supbookRecord(target: string, sheets: readonly string[]): Uint8Array {
  return record(
    0x01ae,
    concat([uint16Body([sheets.length]), biff8String(target), ...sheets.map(shortBiff8String)]),
  );
}

function legacyWorkbook(globals: readonly Uint8Array[]): Uint8Array {
  const begin = record(0x0809, new Uint8Array([0, 6, 5, 0]));
  const sheet = concat([record(0x0809, new Uint8Array([0, 6, 16, 0])), record(0x000a)]);
  const sheetRecord = record(0x0085, boundSheetBody(0));
  const sst = record(0x00fc, new Uint8Array(8));
  const position =
    begin.byteLength +
    globals.reduce((total, value) => total + value.byteLength, 0) +
    sst.byteLength +
    record(0x000a).byteLength +
    sheetRecord.byteLength;
  const positioned = record(0x0085, boundSheetBody(position));
  return concat([begin, ...globals, sst, positioned, record(0x000a), sheet]);
}

function boundSheetBody(position: number): Uint8Array {
  const body = new Uint8Array(6);
  new DataView(body.buffer).setUint32(0, position, true);
  return concat([body, shortBiff8String("Sheet")]);
}

function smallContainer(path: string, data: Uint8Array): Uint8Array {
  const sectorSize = 512;
  const miniSectorSize = 64;
  const miniSectorCount = Math.ceil(data.byteLength / miniSectorSize);
  const rootSectors = Math.ceil((miniSectorCount * miniSectorSize) / sectorSize);
  const bytes = new Uint8Array(sectorSize + (4 + rootSectors) * sectorSize);
  const view = new DataView(bytes.buffer);
  bytes.set(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]));
  view.setUint16(24, 0x3e, true);
  view.setUint16(26, 3, true);
  view.setUint16(28, 0xfffe, true);
  view.setUint16(30, 9, true);
  view.setUint16(32, 6, true);
  view.setUint32(44, 1, true);
  view.setUint32(48, 1, true);
  view.setUint32(56, 4096, true);
  view.setUint32(60, 2, true);
  view.setUint32(64, 1, true);
  view.setUint32(68, 0xfffffffe, true);
  view.setUint32(72, 0, true);
  for (let index = 0; index < 109; index++) {
    view.setUint32(76 + index * 4, index === 0 ? 0 : 0xffffffff, true);
  }
  const fat = [
    0xfffffffd,
    0xfffffffe,
    0xfffffffe,
    ...Array.from<number>({ length: rootSectors }).fill(0xfffffffe),
  ];
  fat.forEach((value, index) => view.setUint32(sectorSize + index * 4, value, true));
  for (let index = 0; index < miniSectorCount - 1; index++) {
    view.setUint32(sectorSize * 3 + index * 4, index + 1, true);
  }
  const entries = [
    { name: "Root Entry", type: 5, child: 1, start: 3, size: miniSectorCount * miniSectorSize },
    { name: path, type: 2, child: 0xffffffff, start: 0, size: data.byteLength },
  ];
  entries.forEach((entry, index) => {
    const offset = sectorSize * 2 + index * 128;
    for (let charIndex = 0; charIndex < entry.name.length; charIndex++) {
      view.setUint16(offset + charIndex * 2, entry.name.charCodeAt(charIndex), true);
    }
    view.setUint16(offset + entry.name.length * 2, 0, true);
    view.setUint16(offset + 64, (entry.name.length + 1) * 2, true);
    bytes[offset + 66] = entry.type;
    bytes[offset + 67] = 1;
    view.setUint32(offset + 68, 0xffffffff, true);
    view.setUint32(offset + 72, 0xffffffff, true);
    view.setUint32(offset + 76, entry.child, true);
    view.setUint32(offset + 116, entry.start, true);
    view.setUint32(offset + 120, entry.size, true);
  });
  bytes.set(data, sectorSize * 4);
  return bytes;
}

export async function xlsSupbookLinks(): Promise<void> {
  const part = "Workbook globals SUPBOOK";
  const externSheet = record(0x0017, uint16Body([1, 0, 0, 0]));
  const ddeData = smallContainer(
    "Workbook",
    legacyWorkbook([supbookRecord("Excel|Topic", ["Item"]), externSheet]),
  );
  const dde = parseLegacyWorkbook(ddeData);
  assertEqual(dde.externalLinks?.[0]?.ddeLink?.ddeService, "Excel", part, "DDE service");
  assertEqual(dde.externalLinks?.[0]?.ddeLink?.ddeItems?.[0]?.name, "Item", part, "DDE item");
  const ole = parseLegacyWorkbook(
    smallContainer(
      "Workbook",
      legacyWorkbook([supbookRecord("package:Word.Document.12", ["Item"]), externSheet]),
    ),
  );
  assertEqual(ole.externalLinks?.[0]?.oleLink?.progId, "Word.Document.12", part, "OLE progId");
  for (const options of [dde, ole]) {
    const output = await generateWorkbook(options, { type: "uint8array" });
    await parseWorkbook(output);
  }
}
