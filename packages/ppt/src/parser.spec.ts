import { describe, expect, it } from "vitest";

import { parsePresentation } from "./parser";

const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;
const SECTOR_SIZE = 512;
const END_OF_CHAIN = 0xfffffffe;
const FAT_SECTOR = 0xfffffffd;
const FREE_SECTOR = 0xffffffff;
const NO_STREAM = 0xffffffff;

interface RecordOptions {
  version?: number;
  instance?: number;
}

function record(type: number, body: readonly number[], options: RecordOptions = {}): number[] {
  const packed = ((options.instance ?? 0) << 4) | (options.version ?? 0);
  return [
    packed & 0xff,
    (packed >>> 8) & 0xff,
    type & 0xff,
    (type >>> 8) & 0xff,
    body.length & 0xff,
    (body.length >>> 8) & 0xff,
    (body.length >>> 16) & 0xff,
    (body.length >>> 24) & 0xff,
    ...body,
  ];
}

function container(type: number, children: readonly number[][], instance = 0): number[] {
  const body = children.flat();
  return record(type, body, { version: 15, instance });
}

function int32(value: number): number[] {
  return [value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff];
}

function int16(value: number): number[] {
  return [value & 0xff, (value >>> 8) & 0xff];
}

function packedPersist(firstId: number, count: number): number[] {
  return int32(((count & 0xfff) << 20) | (firstId & 0xfffff));
}

function currentUserAtom(editOffset: number, encrypted = false): number[] {
  return record(4086, [
    ...int32(20),
    ...int32(encrypted ? 0xf3d1c4df : 0xe391c05f),
    ...int32(editOffset),
    ...int16(0),
    ...int16(0x03f4),
    3,
    0,
    ...int32(8),
  ]);
}

function userEditAtom(pointerOffset: number): number[] {
  return record(4085, [
    ...int32(1),
    ...int32(8),
    ...int32(0),
    ...int32(pointerOffset),
    ...int32(1),
    ...int32(2),
    ...int16(1),
    ...int16(0),
  ]);
}

function buildDocument(): { document: Uint8Array; currentUser: Uint8Array } {
  const slideList = container(4080, [
    record(1011, [...int32(2), ...int32(4), ...int32(1), ...int32(257), ...int32(0), ...int32(0)]),
    record(3999, int32(0)),
    record(4008, [...new TextEncoder().encode("First\rSecond")]),
  ]);
  const document = container(1000, [
    record(1001, [...int32(5760), ...int32(4320), ...Array.from<number>({ length: 40 }).fill(0)]),
    slideList,
  ]);

  const slide = container(1006, [
    container(1036, [
      container(0xf002, [
        container(0xf003, [
          container(0xf004, [
            record(0xf00a, [...int32(2050), ...int32(544)]),
            record(0xf010, [...int16(0), ...int16(576), ...int16(1728), ...int16(1152)]),
            container(0xf00d, [record(3998, int32(0))]),
          ]),
        ]),
      ]),
    ]),
  ]);

  const pointers = record(6002, [
    ...packedPersist(1, 1),
    ...int32(0),
    ...packedPersist(2, 1),
    ...int32(document.length),
  ]);
  const pointerOffset = document.length + slide.length;
  const editOffset = pointerOffset + pointers.length;
  const userEdit = userEditAtom(pointerOffset);
  const paddedDocument = [...document, ...slide, ...pointers, ...userEdit];
  paddedDocument.push(
    ...Array.from<number>({ length: Math.max(0, 4096 - paddedDocument.length) }).fill(0),
  );

  const currentUser = new Uint8Array(4096);
  currentUser.set(currentUserAtom(editOffset));
  return {
    document: new Uint8Array(paddedDocument),
    currentUser,
  };
}

function writeDirectoryEntry(
  bytes: Uint8Array,
  offset: number,
  name: string,
  objectType: number,
  rightSiblingId: number,
  childId: number,
  startSector: number,
  size: number,
): void {
  const view = new DataView(bytes.buffer);
  for (const [index, character] of Array.from(name).entries()) {
    view.setUint16(offset + index * 2, character.charCodeAt(0), true);
  }
  view.setUint16(offset + name.length * 2, 0, true);
  view.setUint16(offset + 64, (name.length + 1) * 2, true);
  bytes[offset + 66] = objectType;
  view.setUint32(offset + 68, NO_STREAM, true);
  view.setUint32(offset + 72, rightSiblingId, true);
  view.setUint32(offset + 76, childId, true);
  view.setUint32(offset + 116, startSector, true);
  view.setUint32(offset + 120, size, true);
}

function buildCfb(streams: readonly { name: string; data: Uint8Array }[]): Uint8Array {
  const streamSectorCounts = streams.map((stream) =>
    Math.ceil(stream.data.byteLength / SECTOR_SIZE),
  );
  const sectorCount = 2 + streamSectorCounts.reduce((total, count) => total + count, 0);
  const bytes = new Uint8Array(512 + sectorCount * SECTOR_SIZE);
  const view = new DataView(bytes.buffer);
  bytes.set(CFB_SIGNATURE);
  view.setUint16(24, 0x3e, true);
  view.setUint16(26, 3, true);
  view.setUint16(28, 0xfffe, true);
  view.setUint16(30, 9, true);
  view.setUint16(32, 6, true);
  view.setUint32(44, 1, true);
  view.setUint32(48, 1, true);
  view.setUint32(56, 4096, true);
  view.setUint32(60, END_OF_CHAIN, true);
  view.setUint32(68, END_OF_CHAIN, true);
  view.setUint32(76, 0, true);

  const fat = Array.from<number>({ length: sectorCount }).fill(FREE_SECTOR);
  fat[0] = FAT_SECTOR;
  fat[1] = END_OF_CHAIN;
  for (let index = 1; index < 109; index += 1) view.setUint32(76 + index * 4, FREE_SECTOR, true);

  let nextSector = 2;
  const startSectors: number[] = [];
  streams.forEach((stream, streamIndex) => {
    const start = nextSector;
    startSectors.push(start);
    const count = streamSectorCounts[streamIndex]!;
    for (let sector = start; sector < start + count - 1; sector += 1) fat[sector] = sector + 1;
    fat[start + count - 1] = END_OF_CHAIN;
    bytes.set(stream.data, 512 + start * SECTOR_SIZE);
    nextSector += count;
  });
  fat.forEach((value, sector) => view.setUint32(512 + sector * 4, value, true));

  writeDirectoryEntry(bytes, 512 + SECTOR_SIZE, "Root Entry", 5, NO_STREAM, 1, 2, 0);
  streams.forEach((stream, index) => {
    const offset = 512 + SECTOR_SIZE + (index + 1) * 128;
    writeDirectoryEntry(
      bytes,
      offset,
      stream.name,
      2,
      index + 2 < streams.length + 1 ? index + 2 : NO_STREAM,
      NO_STREAM,
      startSectors[index]!,
      stream.data.byteLength,
    );
  });
  return bytes;
}

function validFixture(): Uint8Array {
  const { document, currentUser } = buildDocument();
  return buildCfb([
    { name: "PowerPoint Document", data: document },
    { name: "Current User", data: currentUser },
  ]);
}

describe("parsePresentation", () => {
  it("projects persisted slides, anchored text shapes, and paragraph runs", () => {
    const options = parsePresentation(validFixture());
    expect(options.size).toEqual({ width: 9_144_000, height: 6_858_000 });
    expect(options.slides).toHaveLength(1);
    const child = options.slides![0]!.children![0]!;
    if (!("shape" in child)) throw new TypeError("Expected a shape child");
    expect(child.shape?.id).toBe(2050);
    expect(child.shape?.x).toBe(914_400);
    expect(child.shape?.y).toBe(0);
    expect(child.shape?.width).toBe(1_828_800);
    expect(child.shape?.height).toBe(1_828_800);
    expect(child.shape?.textBox).toBe(true);
    expect(child.shape?.textBody?.paragraphs).toEqual([
      { children: [{ text: "First" }] },
      { children: [{ text: "Second" }] },
    ]);
  });

  it("rejects input that is not CFB", () => {
    expect(() => parsePresentation(new Uint8Array(512))).toThrow(
      "Input is not a legacy PowerPoint CFB container",
    );
  });

  it("rejects encrypted Current User atoms", () => {
    const { document } = buildDocument();
    const currentUser = new Uint8Array(4096);
    currentUser.set(currentUserAtom(0, true));
    const encrypted = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
    ]);
    expect(() => parsePresentation(encrypted)).toThrow(
      "Encrypted legacy PowerPoint files are not supported",
    );
  });

  it("rejects an out-of-range Current User edit offset", () => {
    const fixture = validFixture();
    const view = new DataView(fixture.buffer);
    const currentSector = 2 + Math.ceil(buildDocument().document.byteLength / SECTOR_SIZE);
    const currentOffset = 512 + currentSector * SECTOR_SIZE;
    view.setUint32(currentOffset + 16, fixture.byteLength + 1, true);
    expect(() => parsePresentation(fixture)).toThrow(
      "Corrupt legacy PowerPoint file: invalid Current User edit offset",
    );
  });
});
