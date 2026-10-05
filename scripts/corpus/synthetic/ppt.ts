import { parsePresentation as parseLegacyPresentation } from "../../../packages/ppt/dist/index.mjs";
import { generatePresentation, parsePresentation } from "../../../packages/pptx/dist/index.mjs";
import { assert, assertEqual } from "./support";

function record(type: number, body: readonly number[], instance = 0, version = 0): number[] {
  const packed = (instance << 4) | version;
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
  return record(type, children.flat(), instance, 15);
}

function int32(value: number): number[] {
  return [value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff];
}

function int16(value: number): number[] {
  return [value & 0xff, (value >>> 8) & 0xff];
}

function buildCfb(streams: readonly { name: string; data: Uint8Array }[]): Uint8Array {
  const sectorSize = 512;
  const counts = streams.map((stream) => Math.ceil(stream.data.byteLength / sectorSize));
  const sectorCount = 2 + counts.reduce((total, count) => total + count, 0);
  const bytes = new Uint8Array(512 + sectorCount * sectorSize);
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
  view.setUint32(60, 0xfffffffe, true);
  view.setUint32(68, 0xfffffffe, true);
  view.setUint32(76, 0, true);
  const fat = Array.from<number>({ length: sectorCount }).fill(0xffffffff);
  fat[0] = 0xfffffffd;
  fat[1] = 0xfffffffe;
  for (let index = 1; index < 109; index++) view.setUint32(76 + index * 4, 0xffffffff, true);
  let nextSector = 2;
  const starts: number[] = [];
  streams.forEach((stream, streamIndex) => {
    const start = nextSector;
    starts.push(start);
    const count = counts[streamIndex] as number;
    for (let sector = start; sector < start + count - 1; sector++) fat[sector] = sector + 1;
    fat[start + count - 1] = 0xfffffffe;
    bytes.set(stream.data, 512 + start * sectorSize);
    nextSector += count;
  });
  fat.forEach((value, sector) => view.setUint32(512 + sector * 4, value, true));
  const directory = [
    { name: "Root Entry", type: 5, sibling: 0xffffffff, child: 1, start: 2, size: 0 },
    ...streams.map((stream, index) => ({
      name: stream.name,
      type: 2,
      sibling: index + 2 < streams.length + 1 ? index + 2 : 0xffffffff,
      child: 0xffffffff,
      start: starts[index] as number,
      size: stream.data.byteLength,
    })),
  ];
  directory.forEach((entry, index) => {
    const offset = 512 + sectorSize + index * 128;
    for (let charIndex = 0; charIndex < entry.name.length; charIndex++) {
      view.setUint16(offset + charIndex * 2, entry.name.charCodeAt(charIndex), true);
    }
    view.setUint16(offset + entry.name.length * 2, 0, true);
    view.setUint16(offset + 64, (entry.name.length + 1) * 2, true);
    bytes[offset + 66] = entry.type;
    bytes[offset + 67] = 0;
    view.setUint32(offset + 68, 0xffffffff, true);
    view.setUint32(offset + 72, entry.sibling, true);
    view.setUint32(offset + 76, entry.child, true);
    view.setUint32(offset + 116, entry.start, true);
    view.setUint32(offset + 120, entry.size, true);
  });
  return bytes;
}

function noTextShapePackage(): Uint8Array {
  const shape = container(0xf004, [
    record(0xf00a, [...int32(2060), ...int32(0x220)], 1),
    record(0xf010, [...int16(40), ...int16(100), ...int16(300), ...int16(180)]),
  ]);
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
    record(1007, [
      ...int32(5760),
      ...int32(4320),
      ...int32(0),
      ...int32(0),
      ...int32(0),
      ...int32(0),
      ...int16(0),
      ...int16(0),
    ]),
    container(1036, [container(0xf002, [container(0xf003, [shape])])]),
  ]);
  const pointers = record(6002, [
    ...int32(((1 & 0xfff) << 20) | (1 & 0xfffff)),
    ...int32(0),
    ...int32(((1 & 0xfff) << 20) | (2 & 0xfffff)),
    ...int32(document.length),
  ]);
  const pointerOffset = document.length + slide.length;
  const editOffset = pointerOffset + pointers.length;
  const userEdit = record(4085, [
    ...int32(1),
    ...int32(8),
    ...int32(0),
    ...int32(pointerOffset),
    ...int32(1),
    ...int32(2),
    ...int16(1),
    ...int16(0),
  ]);
  const rawStream = [...document, ...slide, ...pointers, ...userEdit];
  rawStream.push(...Array.from<number>({ length: Math.max(0, 4096 - rawStream.length) }).fill(0));
  const stream = new Uint8Array(rawStream);
  const currentUser = new Uint8Array(4096);
  currentUser.set(
    new Uint8Array(
      record(4086, [
        ...int32(20),
        ...int32(0xe391c05f),
        ...int32(editOffset),
        ...int16(0),
        ...int16(0x03f4),
        3,
        0,
        ...int32(8),
      ]),
    ),
  );
  return buildCfb([
    { name: "PowerPoint Document", data: stream },
    { name: "Current User", data: currentUser },
  ]);
}

export async function pptNoTextShape(): Promise<void> {
  const part = "PowerPoint Document shape";
  const options = parseLegacyPresentation(noTextShapePackage());
  const child = options.slides?.[0]?.children?.[0];
  assert(child && "shape" in child, part, "no-text shape projection");
  assertEqual(child.shape?.id, 2060, part, "shape id");
  assertEqual(child.shape?.textBox, undefined, part, "empty text box");
  const output = await generatePresentation(options, { type: "uint8array" });
  const generated = await parsePresentation(output);
  assert(
    "shape" in (generated.slides?.[0]?.children?.[0] ?? {}),
    part,
    "generated shape projection",
  );
}
