import { describe, expect, it } from "vitest";

import { parseDocument } from "./index";

const END_OF_CHAIN = 0xfffffffe;
const FAT_SECT = 0xfffffffd;
const FREE_SECT = 0xffffffff;
const NO_STREAM = 0xffffffff;
const SECTOR_SIZE = 512;
const STREAM_SIZE = 4096;

function fillName(bytes: Uint8Array, offset: number, name: string): void {
  const view = new DataView(bytes.buffer);
  for (const [index, character] of Array.from(name).entries()) {
    view.setUint16(offset + index * 2, character.charCodeAt(0), true);
  }
  view.setUint16(offset + name.length * 2, 0, true);
  view.setUint16(offset + 64, (name.length + 1) * 2, true);
}

function writeDirectoryEntry(
  bytes: Uint8Array,
  offset: number,
  entry: {
    name: string;
    objectType: number;
    rightSiblingId: number;
    childId: number;
    startSector: number;
    size: number;
  },
): void {
  const view = new DataView(bytes.buffer);
  fillName(bytes, offset, entry.name);
  bytes[offset + 66] = entry.objectType;
  bytes[offset + 67] = 1;
  view.setUint32(offset + 68, NO_STREAM, true);
  view.setUint32(offset + 72, entry.rightSiblingId, true);
  view.setUint32(offset + 76, entry.childId, true);
  view.setUint32(offset + 116, entry.startSector, true);
  view.setUint32(offset + 120, entry.size, true);
}

function buildContainer(streams: readonly { path: string; data: Uint8Array }[]): Uint8Array {
  const paths = ["", ...streams.map((stream) => stream.path)];
  const idByPath = new Map(paths.map((path, index) => [path, index]));
  const directorySectorCount = Math.ceil(paths.length / 4);
  const regularSectorCounts = streams.map((stream) =>
    Math.ceil(stream.data.byteLength / SECTOR_SIZE),
  );
  const sectorCount =
    1 + directorySectorCount + regularSectorCounts.reduce((sum, count) => sum + count, 0);
  const bytes = new Uint8Array(SECTOR_SIZE + sectorCount * SECTOR_SIZE);
  const view = new DataView(bytes.buffer);
  const fat = Array.from<number>({ length: sectorCount }).fill(FREE_SECT);
  const signature = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  signature.forEach((byte, index) => {
    bytes[index] = byte;
  });
  view.setUint16(24, 0x3e, true);
  view.setUint16(26, 3, true);
  view.setUint16(28, 0xfffe, true);
  view.setUint16(30, 9, true);
  view.setUint16(32, 6, true);
  view.setUint32(44, 1, true);
  view.setUint32(48, 1, true);
  view.setUint32(56, 4096, true);
  view.setUint32(60, END_OF_CHAIN, true);
  view.setUint32(64, 0, true);
  view.setUint32(68, END_OF_CHAIN, true);
  for (let index = 1; index < 109; index++) view.setUint32(76 + index * 4, FREE_SECT, true);

  let nextSector = 1 + directorySectorCount;
  const sectorByPath = new Map<string, number>();
  for (const [index, stream] of streams.entries()) {
    sectorByPath.set(stream.path, nextSector);
    const count = regularSectorCounts[index]!;
    for (let offset = 0; offset < count - 1; offset++)
      fat[nextSector + offset] = nextSector + offset + 1;
    fat[nextSector + count - 1] = END_OF_CHAIN;
    nextSector += count;
  }

  fat[0] = FAT_SECT;
  fat[1] = END_OF_CHAIN;
  const fatOffset = SECTOR_SIZE;
  fat.forEach((value, index) => view.setUint32(fatOffset + index * 4, value, true));

  const directoryOffset = SECTOR_SIZE * 2;
  for (const path of paths) {
    const stream = streams.find((candidate) => candidate.path === path);
    const id = idByPath.get(path)!;
    writeDirectoryEntry(bytes, directoryOffset + id * 128, {
      name: path || "Root Entry",
      objectType: stream ? 2 : path === "" ? 5 : 1,
      rightSiblingId: id + 1 < paths.length ? id + 1 : NO_STREAM,
      childId: path === "" ? 1 : NO_STREAM,
      startSector: stream ? sectorByPath.get(path)! : NO_STREAM,
      size: stream?.data.byteLength ?? 0,
    });
  }
  for (const stream of streams) {
    bytes.set(stream.data, SECTOR_SIZE * (1 + sectorByPath.get(stream.path)!));
  }
  return bytes;
}

function buildDocument(options: { characterBinTableLength?: number; secondCp?: number } = {}): {
  data: Uint8Array;
  word: Uint8Array;
  table: Uint8Array;
} {
  const word = new Uint8Array(STREAM_SIZE);
  const table = new Uint8Array(STREAM_SIZE);
  const wordView = new DataView(word.buffer);
  const tableView = new DataView(table.buffer);

  wordView.setUint16(0, 0xa5ec, true);
  wordView.setUint16(2, 0x00c1, true);
  wordView.setUint16(10, 0x0200, true);
  wordView.setUint16(32, 14, true);
  wordView.setUint16(62, 22, true);
  wordView.setUint32(76, 7, true);
  wordView.setUint16(152, 35, true);

  const setFibPair = (index: number, offset: number, length: number): void => {
    wordView.setUint32(154 + index * 8, offset, true);
    wordView.setUint32(154 + index * 8 + 4, length, true);
  };
  setFibPair(12, 0, options.characterBinTableLength ?? 12);
  setFibPair(13, 12, 12);
  setFibPair(34, 512, 33);

  word.set(new TextEncoder().encode("Hi\r"), 512);
  const unicode = new Uint8Array(6);
  const unicodeView = new DataView(unicode.buffer);
  unicodeView.setUint16(0, 0x4e16, true);
  unicodeView.setUint16(2, 0x754c, true);
  unicodeView.setUint16(4, 0x000d, true);
  word.set(unicode, 516);

  tableView.setUint8(512, 2);
  tableView.setUint32(513, 28, true);
  tableView.setUint32(517, 0, true);
  tableView.setUint32(521, options.secondCp ?? 4, true);
  tableView.setUint32(525, 7, true);
  tableView.setUint32(531, 0x40000400, true);
  tableView.setUint32(539, 516, true);

  tableView.setUint32(0, 512, true);
  tableView.setUint32(4, 528, true);
  tableView.setUint32(8, 2, true);
  tableView.setUint32(12, 512, true);
  tableView.setUint32(16, 528, true);
  tableView.setUint32(20, 3, true);

  const characterPage = 1024;
  tableView.setUint32(characterPage, 512, true);
  tableView.setUint32(characterPage + 4, 528, true);
  tableView.setUint8(characterPage + 8, 20);
  tableView.setUint8(characterPage + 511, 1);
  const grpprl = [
    0x35, 0x08, 0x01, 0x00, 0xc0, 0x02, 0xaa, 0xbb, 0x43, 0x4a, 0x18, 0x00, 0x00, 0x00, 0x00,
  ];
  table.set(Uint8Array.from(grpprl), characterPage + 41);
  tableView.setUint8(characterPage + 40, grpprl.length);

  const paragraphPage = 1536;
  tableView.setUint32(paragraphPage, 512, true);
  tableView.setUint32(paragraphPage + 4, 528, true);
  tableView.setUint8(paragraphPage + 8, 0);
  tableView.setUint8(paragraphPage + 511, 1);

  const data = buildContainer([
    { path: "WordDocument", data: word },
    { path: "1Table", data: table },
  ]);
  return { data, word, table };
}

describe("legacy DOC parser", () => {
  it("projects compressed and Unicode pieces with character properties", () => {
    const { data } = buildDocument();
    const document = parseDocument(data);
    const children = document.sections[0]!.children;

    expect(children).toHaveLength(2);
    expect(children[0]).toEqual({
      paragraph: { children: [{ text: "Hi", bold: true, size: 12 }] },
    });
    expect(children[1]).toEqual({
      paragraph: { children: [{ text: "世界", bold: true, size: 12 }] },
    });
  });

  it("validates PlcfBteChpx FC and page array boundaries", () => {
    const { data } = buildDocument({ characterBinTableLength: 13 });
    expect(() => parseDocument(data)).toThrow("character bin table");
  });
});
