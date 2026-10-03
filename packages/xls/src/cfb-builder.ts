const END_OF_CHAIN = 0xfffffffe;
const FAT_SECT = 0xfffffffd;
const FREE_SECT = 0xffffffff;
const NO_STREAM = 0xffffffff;
const SECTOR_SIZE = 512;
const MINI_SECTOR_SIZE = 64;

export interface TestStream {
  path: string;
  data: Uint8Array;
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
  for (const [index, char] of Array.from(entry.name).entries()) {
    view.setUint16(offset + index * 2, char.charCodeAt(0), true);
  }
  view.setUint16(offset + entry.name.length * 2, 0, true);
  view.setUint16(offset + 64, (entry.name.length + 1) * 2, true);
  bytes[offset + 66] = entry.objectType;
  bytes[offset + 67] = 1;
  view.setUint32(offset + 68, NO_STREAM, true);
  view.setUint32(offset + 72, entry.rightSiblingId, true);
  view.setUint32(offset + 76, entry.childId, true);
  view.setUint32(offset + 116, entry.startSector, true);
  view.setUint32(offset + 120, entry.size, true);
}

export function buildSmallContainer(stream: TestStream): Uint8Array {
  if (stream.path.includes("/") || stream.data.byteLength >= 4096 || stream.data.byteLength === 0) {
    throw new Error("Test CFB containers only support one non-empty root stream below 4096 bytes");
  }

  const miniSectorCount = Math.ceil(stream.data.byteLength / MINI_SECTOR_SIZE);
  const rootMiniSectorCount = Math.ceil((miniSectorCount * MINI_SECTOR_SIZE) / SECTOR_SIZE);
  const sectorCount = 4 + rootMiniSectorCount;
  const bytes = new Uint8Array(SECTOR_SIZE + sectorCount * SECTOR_SIZE);
  const view = new DataView(bytes.buffer);
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
  view.setUint32(60, 2, true);
  view.setUint32(64, 1, true);
  view.setUint32(68, END_OF_CHAIN, true);
  view.setUint32(72, 0, true);
  for (let index = 0; index < 109; index++)
    view.setUint32(76 + index * 4, index === 0 ? 0 : FREE_SECT, true);

  const fat = [
    FAT_SECT,
    END_OF_CHAIN,
    END_OF_CHAIN,
    ...Array.from<number>({ length: rootMiniSectorCount }).fill(END_OF_CHAIN),
  ];
  fat.forEach((value, index) => view.setUint32(SECTOR_SIZE + index * 4, value, true));
  for (let index = 0; index < miniSectorCount - 1; index++) {
    view.setUint32(SECTOR_SIZE * 3 + index * 4, index + 1, true);
  }

  writeDirectoryEntry(bytes, SECTOR_SIZE * 2, {
    name: "Root Entry",
    objectType: 5,
    rightSiblingId: NO_STREAM,
    childId: 1,
    startSector: 3,
    size: miniSectorCount * MINI_SECTOR_SIZE,
  });
  writeDirectoryEntry(bytes, SECTOR_SIZE * 2 + 128, {
    name: stream.path,
    objectType: 2,
    rightSiblingId: NO_STREAM,
    childId: NO_STREAM,
    startSector: 0,
    size: stream.data.byteLength,
  });
  bytes.set(stream.data, SECTOR_SIZE * 4);
  return bytes;
}
