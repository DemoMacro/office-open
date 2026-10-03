import { describe, expect, it } from "vitest";

import { CompoundFileReader } from "./reader";

const END_OF_CHAIN = 0xfffffffe;
const DIF_SECT = 0xfffffffc;
const FAT_SECT = 0xfffffffd;
const FREE_SECT = 0xffffffff;
const NO_STREAM = 0xffffffff;
const SECTOR_SIZE = 512;

interface TestStream {
  path: string;
  data: Uint8Array;
}

interface BuiltContainer {
  bytes: Uint8Array;
  view: DataView;
  fatOffset: number;
  miniFatOffset: number | undefined;
  directoryOffset: number;
  sectorByPath: Map<string, number>;
  idByPath: Map<string, number>;
}

function fillName(bytes: Uint8Array, offset: number, name: string): void {
  const view = new DataView(bytes.buffer);
  for (const [index, char] of Array.from(name).entries()) {
    view.setUint16(offset + index * 2, char.charCodeAt(0), true);
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

function buildContainer(streams: readonly TestStream[]): BuiltContainer {
  const nodePaths = new Set<string>([""]);
  for (const stream of streams) {
    let parent = "";
    for (const segment of stream.path.split("/")) {
      parent = parent ? `${parent}/${segment}` : segment;
      nodePaths.add(parent);
    }
  }
  const paths = [...nodePaths];
  const idByPath = new Map(paths.map((path, index) => [path, index]));
  const childIds = new Map<string, number[]>();
  for (const path of paths.slice(1)) {
    const parent = path.includes("/") ? path.slice(0, path.lastIndexOf("/")) : "";
    const children = childIds.get(parent) ?? [];
    children.push(idByPath.get(path)!);
    childIds.set(parent, children);
  }

  const miniStreams = streams.filter((stream) => stream.data.byteLength < 4096);
  const regularStreams = streams.filter((stream) => stream.data.byteLength >= 4096);
  const miniSectorCounts = miniStreams.map((stream) => Math.ceil(stream.data.byteLength / 64));
  const miniCapacity = miniSectorCounts.reduce((sum, count) => sum + count, 0);
  const rootMiniSectorCount = Math.ceil((miniCapacity * 64) / SECTOR_SIZE);
  const directorySectorCount = Math.ceil(paths.length / 4);
  const miniFatSector = 1 + directorySectorCount;
  const regularSectorCounts = regularStreams.map((stream) =>
    Math.ceil(stream.data.byteLength / SECTOR_SIZE),
  );
  const miniFatSectorCount = miniCapacity > 1 ? 1 : 0;
  const sectorCount =
    1 +
    directorySectorCount +
    miniFatSectorCount +
    rootMiniSectorCount +
    regularSectorCounts.reduce((sum, count) => sum + count, 0);
  const bytes = new Uint8Array(HEADER_AND_SECTORS(sectorCount));
  const view = new DataView(bytes.buffer);
  const fat = Array.from<number>({ length: sectorCount }).fill(FREE_SECT);

  for (const [index, byte] of [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].entries()) {
    bytes[index] = byte;
  }
  view.setUint16(24, 0x3e, true);
  view.setUint16(26, 3, true);
  view.setUint16(28, 0xfffe, true);
  view.setUint16(30, 9, true);
  view.setUint16(32, 6, true);
  view.setUint32(44, 1, true);
  view.setUint32(48, 1, true);
  view.setUint32(56, 4096, true);
  view.setUint32(60, miniFatSectorCount ? miniFatSector : END_OF_CHAIN, true);
  view.setUint32(64, miniFatSectorCount, true);
  view.setUint32(68, END_OF_CHAIN, true);
  view.setUint32(72, 0, true);
  view.setUint32(76, 0, true);
  for (let index = 1; index < 109; index++) view.setUint32(76 + index * 4, FREE_SECT, true);

  const markChain = (start: number, count: number): void => {
    for (let index = 0; index < count - 1; index++) fat[start + index] = start + index + 1;
    fat[start + count - 1] = END_OF_CHAIN;
  };
  const rootMiniSector = miniFatSector + miniFatSectorCount;
  fat[0] = FAT_SECT;
  fat[1] = END_OF_CHAIN;
  for (let index = 0; index < directorySectorCount - 1; index++) fat[1 + index] = 2 + index;
  fat[directorySectorCount] = END_OF_CHAIN;
  if (miniFatSectorCount) fat[miniFatSector] = END_OF_CHAIN;
  if (rootMiniSectorCount) markChain(rootMiniSector, rootMiniSectorCount);

  let nextSector = rootMiniSector + rootMiniSectorCount;
  let nextMiniSector = 0;
  const sectorByPath = new Map<string, number>();
  const miniFatValues = new Map<number, number>();
  for (const [index, stream] of miniStreams.entries()) {
    const count = miniSectorCounts[index]!;
    sectorByPath.set(stream.path, nextMiniSector);
    for (let offset = 0; offset < count; offset++) {
      miniFatValues.set(
        nextMiniSector + offset,
        offset + 1 === count ? END_OF_CHAIN : nextMiniSector + offset + 1,
      );
    }
    nextMiniSector += count;
  }

  for (const [index, stream] of regularStreams.entries()) {
    const count = regularSectorCounts[index]!;
    sectorByPath.set(stream.path, nextSector);
    markChain(nextSector, count);
    nextSector += count;
  }

  const fatOffset = SECTOR_SIZE;
  fat.forEach((value, index) => view.setUint32(fatOffset + index * 4, value, true));
  const miniFatOffset = miniFatSectorCount ? HEADER_AND_SECTORS(miniFatSector) : undefined;
  if (miniFatOffset) {
    for (const [sector, value] of miniFatValues)
      view.setUint32(miniFatOffset + sector * 4, value, true);
  }

  const directoryOffset = HEADER_AND_SECTORS(1);
  for (const path of paths) {
    const stream = streams.find((candidate) => candidate.path === path);
    const id = idByPath.get(path)!;
    const parent = path.includes("/") ? path.slice(0, path.lastIndexOf("/")) : "";
    const siblings = childIds.get(parent) ?? [];
    const position = siblings.indexOf(id);
    const children = childIds.get(path) ?? [];
    writeDirectoryEntry(bytes, directoryOffset + idByPath.get(path)! * 128, {
      name: path.includes("/") ? path.slice(path.lastIndexOf("/") + 1) : path || "Root Entry",
      objectType: stream ? 2 : path === "" ? 5 : 1,
      rightSiblingId: siblings[position + 1] ?? NO_STREAM,
      childId: children[0] ?? NO_STREAM,
      startSector: stream
        ? stream.data.byteLength < 4096
          ? sectorByPath.get(stream.path)!
          : sectorByPath.get(stream.path)!
        : path === ""
          ? rootMiniSector
          : NO_STREAM,
      size: stream ? stream.data.byteLength : path === "" ? miniCapacity * 64 : 0,
    });
  }

  for (const stream of miniStreams) {
    const start = sectorByPath.get(stream.path)!;
    const target = HEADER_AND_SECTORS(rootMiniSector) + start * 64;
    if (target + stream.data.byteLength > bytes.byteLength) {
      throw new Error(
        `mini write overflow: target=${target} length=${stream.data.byteLength} bytes=${bytes.byteLength}`,
      );
    }
    bytes.set(stream.data, target);
  }
  for (const stream of regularStreams) {
    bytes.set(stream.data, HEADER_AND_SECTORS(sectorByPath.get(stream.path)!));
  }

  return { bytes, view, fatOffset, miniFatOffset, directoryOffset, sectorByPath, idByPath };
}

function HEADER_AND_SECTORS(sector: number): number {
  return SECTOR_SIZE + sector * SECTOR_SIZE;
}

function buildDifatContainer(): Uint8Array {
  const fatSectorCount = 110;
  const directorySector = 110;
  const difatSector = 111;
  const bytes = new Uint8Array(SECTOR_SIZE + 112 * SECTOR_SIZE);
  const view = new DataView(bytes.buffer);
  for (const [index, byte] of [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].entries()) {
    bytes[index] = byte;
  }
  view.setUint16(24, 0x3e, true);
  view.setUint16(26, 3, true);
  view.setUint16(28, 0xfffe, true);
  view.setUint16(30, 9, true);
  view.setUint16(32, 6, true);
  view.setUint32(44, fatSectorCount, true);
  view.setUint32(48, directorySector, true);
  view.setUint32(56, 4096, true);
  view.setUint32(60, END_OF_CHAIN, true);
  view.setUint32(64, 0, true);
  view.setUint32(68, difatSector, true);
  view.setUint32(72, 1, true);
  for (let index = 0; index < 109; index++) view.setUint32(76 + index * 4, index, true);

  const difatOffset = HEADER_AND_SECTORS(difatSector);
  view.setUint32(difatOffset, 109, true);
  for (let index = 1; index < 127; index++) {
    view.setUint32(difatOffset + index * 4, FREE_SECT, true);
  }
  view.setUint32(difatOffset + 127 * 4, END_OF_CHAIN, true);

  for (let sector = 0; sector < fatSectorCount; sector++) {
    const fatOffset = SECTOR_SIZE + sector * SECTOR_SIZE;
    for (let index = 0; index < 128; index++)
      view.setUint32(fatOffset + index * 4, FREE_SECT, true);
  }
  const firstFatOffset = SECTOR_SIZE;
  for (let sector = 0; sector < fatSectorCount; sector++) {
    view.setUint32(firstFatOffset + sector * 4, FAT_SECT, true);
  }
  view.setUint32(firstFatOffset + directorySector * 4, END_OF_CHAIN, true);
  view.setUint32(firstFatOffset + difatSector * 4, DIF_SECT, true);
  writeDirectoryEntry(bytes, HEADER_AND_SECTORS(directorySector), {
    name: "Root Entry",
    objectType: 5,
    rightSiblingId: NO_STREAM,
    childId: NO_STREAM,
    startSector: NO_STREAM,
    size: 0,
  });
  return bytes;
}

describe("CompoundFileReader", () => {
  it("reads nested mini and regular streams", () => {
    const small = Uint8Array.from({ length: 16 }, (_, index) => index + 1);
    const nested = Uint8Array.from({ length: 80 }, (_, index) => (index * 3) % 251);
    const large = Uint8Array.from({ length: 5000 }, (_, index) => (index + 7) % 256);
    const built = buildContainer([
      { path: "small.bin", data: small },
      { path: "folder/data.bin", data: nested },
      { path: "large.bin", data: large },
    ]);
    expect(built.view.getUint32(built.directoryOffset + 128 + 72, true)).toBe(2);
    expect(built.view.getUint32(built.directoryOffset + 256 + 76, true)).toBe(3);
    expect(built.view.getUint32(built.directoryOffset + 256 + 72, true)).toBe(4);
    const reader = new CompoundFileReader(built.bytes);

    expect(reader.entries.map((entry) => entry.path)).toEqual([
      "small.bin",
      "folder",
      "folder/data.bin",
      "large.bin",
    ]);
    expect(reader.entry("folder")?.type).toBe("storage");
    expect(reader.entry("large.bin")?.size).toBe(5000);
    expect(reader.read("small.bin")).toEqual(small);
    expect(reader.read("folder/data.bin")).toEqual(nested);
    expect(reader.read("large.bin")).toEqual(large);
    expect(reader.entry("small.bin")).toEqual({
      name: "small.bin",
      path: "small.bin",
      type: "stream",
      size: 16,
    });
  });

  it("follows DIFAT when FAT sectors exceed the header array", () => {
    const reader = new CompoundFileReader(buildDifatContainer());
    expect(reader.entries).toEqual([]);
  });

  it("reads left and right directory siblings", () => {
    const built = buildContainer([
      { path: "a.bin", data: new Uint8Array(16) },
      { path: "b.bin", data: new Uint8Array(16) },
    ]);
    const aId = built.idByPath.get("a.bin")!;
    const bId = built.idByPath.get("b.bin")!;
    built.view.setUint32(built.directoryOffset + 76, bId, true);
    built.view.setUint32(built.directoryOffset + aId * 128 + 72, NO_STREAM, true);
    built.view.setUint32(built.directoryOffset + bId * 128 + 68, aId, true);

    const reader = new CompoundFileReader(built.bytes);
    expect(reader.entries.map((entry) => entry.path)).toEqual(["b.bin", "a.bin"]);
  });

  it("tolerates malformed legacy root-name metadata", () => {
    const built = buildContainer([{ path: "small.bin", data: new Uint8Array(80) }]);
    built.view.setUint16(built.directoryOffset + 64, 2, true);

    const reader = new CompoundFileReader(built.bytes);
    expect(reader.read("small.bin")).toHaveLength(80);
  });

  it("rejects a bad signature", () => {
    const built = buildContainer([{ path: "small.bin", data: new Uint8Array(16) }]);
    built.bytes[0] = 0xd1;
    expect(() => new CompoundFileReader(built.bytes)).toThrow("bad signature");
  });

  it("rejects a FAT chain cycle", () => {
    const built = buildContainer([{ path: "large.bin", data: new Uint8Array(5000) }]);
    const start = built.sectorByPath.get("large.bin")!;
    built.view.setUint32(built.fatOffset + start * 4, start, true);
    expect(() => new CompoundFileReader(built.bytes).read("large.bin")).toThrow(
      "stream chain contains a cycle",
    );
  });

  it("rejects a Mini FAT chain cycle", () => {
    const built = buildContainer([{ path: "folder/data.bin", data: new Uint8Array(80) }]);
    const start = built.sectorByPath.get("folder/data.bin")!;
    built.view.setUint32(built.miniFatOffset! + start * 4, start, true);
    expect(() => new CompoundFileReader(built.bytes).read("folder/data.bin")).toThrow(
      "mini-stream chain contains a cycle",
    );
  });

  it("rejects a directory tree cycle", () => {
    const built = buildContainer([{ path: "folder/data.bin", data: new Uint8Array(80) }]);
    const folderId = built.idByPath.get("folder")!;
    built.view.setUint32(built.directoryOffset + folderId * 128 + 76, folderId, true);
    expect(() => new CompoundFileReader(built.bytes)).toThrow("directory tree contains a cycle");
  });

  it("rejects out-of-bounds sector references", () => {
    const built = buildContainer([{ path: "small.bin", data: new Uint8Array(16) }]);
    built.view.setUint32(48, 999_999, true);
    expect(() => new CompoundFileReader(built.bytes)).toThrow("directory start sector");
  });
});
