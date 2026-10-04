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

function utf16(value: string): number[] {
  const bytes: number[] = [];
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    bytes.push(code & 0xff, (code >>> 8) & 0xff);
  }
  return bytes;
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

function anchoredTextShapeBody(): number[][] {
  return [
    record(0xf00a, [...int32(2050), ...int32(544)]),
    record(0xf010, [...int16(0), ...int16(576), ...int16(1728), ...int16(1152)]),
    container(0xf00d, [record(3998, int32(0))]),
  ];
}

function anchoredTextShape(): number[] {
  return container(0xf004, anchoredTextShapeBody());
}

function childTextShape(
  id: number,
  textIndex: number,
  left: number,
  top: number,
  right: number,
  bottom: number,
): number[] {
  return container(0xf004, [
    record(0xf00a, [...int32(id), ...int32(0)]),
    record(0xf00f, [...int32(left), ...int32(top), ...int32(right), ...int32(bottom)]),
    container(0xf00d, [record(3998, int32(textIndex))]),
  ]);
}

function groupShape(
  children: readonly number[][],
  compactBounds = false,
  height = 150,
): number[][] {
  const bounds = compactBounds
    ? [...int32(0), ...int32(0), ...int32(300), ...int32(150)]
    : [
        ...int32(0),
        ...int32(0),
        ...int32(300),
        ...int32(height),
        ...int32(0),
        ...int32(0),
        ...int32(300),
        ...int32(height),
      ];
  return [
    container(0xf004, [
      record(0xf00a, [...int32(4000), ...int32(1)], { instance: 202 }),
      record(0xf009, bounds),
    ]),
    ...children,
  ];
}

function nestedGroupDrawing(children: readonly number[][], height = 150): number[][] {
  return [container(0xf003, groupShape(children, false, height))];
}

function compactGroupDrawing(children: readonly number[][]): number[][] {
  return [container(0xf003, groupShape(children, true))];
}

function wrappedCompactGroupDrawing(children: readonly number[][]): number[][] {
  return [container(0xf003, [...groupShape([], true), ...compactGroupDrawing(children)])];
}

function groupedTableDrawing(): number[][] {
  const cells = [
    [0, 0, 100, 50],
    [100, 0, 200, 50],
    [200, 0, 300, 50],
    [0, 50, 100, 100],
    [100, 50, 200, 100],
    [200, 50, 300, 100],
  ];
  return nestedGroupDrawing(
    cells.map((bounds, index) =>
      childTextShape(4100 + index, index, bounds[0]!, bounds[1]!, bounds[2]!, bounds[3]!),
    ),
  );
}

function compactGroupedTableDrawing(): number[][] {
  const cells = [
    [0, 0, 100, 50],
    [100, 0, 200, 50],
    [200, 0, 300, 50],
    [0, 50, 100, 100],
    [100, 50, 200, 100],
    [200, 50, 300, 100],
  ];
  return compactGroupDrawing(
    cells.map((bounds, index) =>
      childTextShape(4200 + index, index, bounds[0]!, bounds[1]!, bounds[2]!, bounds[3]!),
    ),
  );
}

function wrappedCompactGroupedTableDrawing(): number[][] {
  const cells = [
    [0, 0, 100, 50],
    [100, 0, 200, 50],
    [200, 0, 300, 50],
    [0, 50, 100, 100],
    [100, 50, 200, 100],
    [200, 50, 300, 100],
  ];
  const shapes = cells.map((bounds, index) =>
    childTextShape(4300 + index, index, bounds[0]!, bounds[1]!, bounds[2]!, bounds[3]!),
  );
  return wrappedCompactGroupDrawing(shapes);
}

function pictureShape(id: number): number[] {
  return container(0xf004, [
    record(0xf00a, [...int32(id), ...int32(544)], { instance: 75 }),
    record(0xf010, [...int16(100), ...int16(200), ...int16(500), ...int16(400)]),
    record(0xf00b, [0x04, 0xc1, 0x01, 0x00, 0x00, 0x00], { version: 3, instance: 1 }),
  ]);
}

function pictureShapeWithReference(id: number, reference: number): number[] {
  return container(0xf004, [
    record(0xf00a, [...int32(id), ...int32(544)], { instance: 75 }),
    record(0xf010, [...int16(100), ...int16(200), ...int16(500), ...int16(400)]),
    record(0xf00b, [...int16(0xc104), ...int32(reference)], { version: 3, instance: 1 }),
  ]);
}

function pictureTextShape(id: number): number[] {
  return container(0xf004, [
    record(0xf00a, [...int32(id), ...int32(544)], { instance: 75 }),
    record(0xf010, [...int16(100), ...int16(200), ...int16(500), ...int16(400)]),
    record(0xf00b, [0x04, 0xc1, 0x01, 0x00, 0x00, 0x00], { version: 3, instance: 1 }),
    container(0xf00d, [record(3998, int32(0))]),
  ]);
}

function fieldShape(id: number, atomType: number): number[] {
  return container(0xf004, [
    record(0xf00a, [...int32(id), ...int32(544)]),
    record(0xf010, [...int16(100), ...int16(200), ...int16(500), ...int16(100)]),
    container(0xf00d, [
      record(3999, int32(4)),
      record(4000, utf16("*")),
      record(atomType, int32(0)),
    ]),
  ]);
}

function embeddedTextShape(id: number, text: string): number[] {
  return container(0xf004, [
    record(0xf00a, [...int32(id), ...int32(544)]),
    record(0xf010, [...int16(0), ...int16(0), ...int16(1728), ...int16(576)]),
    container(0xf00d, [record(3999, int32(0)), record(4000, utf16(text))]),
  ]);
}

function animationAtom(effect: number, automatic = false): number[] {
  return record(4081, [
    ...int32(0),
    ...int16(automatic ? 0x0002 : 0x0000),
    ...int16(0),
    ...int32(0),
    ...int32(120),
    ...int16(1),
    ...int16(0),
    0,
    effect,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
  ]);
}

function animatedTextShape(): number[] {
  return container(0xf004, [...anchoredTextShapeBody(), container(4116, [animationAtom(0x06)])]);
}

function hyperlinkTextShape(): number[] {
  return container(0xf004, [
    record(0xf00a, [...int32(2050), ...int32(544)]),
    record(0xf010, [...int16(0), ...int16(576), ...int16(1728), ...int16(1152)]),
    container(0xf00d, [
      record(3999, int32(0)),
      record(4000, utf16("Slide text")),
      container(4082, [
        record(4083, [...int32(0), ...int32(1), ...Array.from<number>({ length: 8 }).fill(0)]),
      ]),
      record(4063, [...int32(0), ...int32(10)]),
    ]),
  ]);
}

function mergedTableDrawing(): number[][] {
  const cells: readonly (readonly [number, number, number, number])[] = [
    [0, 0, 200, 50],
    [200, 0, 300, 50],
    [0, 50, 100, 100],
    [100, 50, 200, 100],
    [200, 50, 300, 100],
    [0, 100, 100, 150],
    [100, 100, 200, 200],
    [200, 100, 300, 150],
    [0, 150, 100, 200],
    [200, 150, 300, 200],
  ];
  return nestedGroupDrawing(
    cells.map((bounds, index) =>
      childTextShape(4400 + index, index, bounds[0]!, bounds[1]!, bounds[2]!, bounds[3]!),
    ),
    200,
  );
}

function buildDocument(
  shapeContainers: readonly number[][] = [anchoredTextShape()],
  textValues: readonly string[] = ["First\rSecond"],
  extras: { notes?: boolean; animation?: boolean; hyperlink?: boolean } = {},
): { document: Uint8Array; currentUser: Uint8Array } {
  const slideList = container(4080, [
    record(1011, [...int32(2), ...int32(4), ...int32(1), ...int32(257), ...int32(0), ...int32(0)]),
    ...textValues.flatMap((value) => [
      record(3999, int32(0)),
      record(4008, [...new TextEncoder().encode(value)]),
    ]),
  ]);
  const documentChildren = [
    record(1001, [...int32(5760), ...int32(4320), ...Array.from<number>({ length: 40 }).fill(0)]),
    slideList,
  ];
  const notesList = container(
    4080,
    [record(1011, [...int32(3), ...int32(0), ...int32(0), ...int32(123), ...int32(0)])],
    2,
  );
  if (extras.notes) documentChildren.push(notesList);
  if (extras.hyperlink) {
    documentChildren.push(
      container(1033, [
        container(4055, [
          record(4051, int32(1)),
          record(4026, utf16("Tooltip"), { instance: 0 }),
          record(4026, utf16("https://example.com"), { instance: 16 }),
        ]),
      ]),
    );
  }
  const document = container(1000, documentChildren);

  const slideAtom = record(1007, [
    ...int32(5760),
    ...int32(4320),
    ...int32(0),
    ...int32(0),
    ...int32(0),
    ...int32(extras.notes ? 123 : 0),
    ...int16(0),
    ...int16(0),
  ]);
  const drawing = container(1036, [container(0xf002, [container(0xf003, shapeContainers)])]);
  const slide = container(1006, [slideAtom, drawing]);

  const notes = container(1008, [
    record(1009, [...int32(123), ...int16(0), ...int16(0)]),
    container(1036, [
      container(0xf002, [container(0xf003, [embeddedTextShape(5000, "Speaker note")])]),
    ]),
  ]);

  const pointerBody = [
    ...packedPersist(1, 1),
    ...int32(0),
    ...packedPersist(2, 1),
    ...int32(document.length),
  ];
  if (extras.notes) {
    pointerBody.push(...packedPersist(3, 1), ...int32(document.length + slide.length));
  }
  const pointers = record(6002, pointerBody);
  const pointerOffset = document.length + slide.length + (extras.notes ? notes.length : 0);
  const editOffset = pointerOffset + pointers.length;
  const userEdit = userEditAtom(pointerOffset);
  const paddedDocument = extras.notes
    ? [...document, ...slide, ...notes, ...pointers, ...userEdit]
    : [...document, ...slide, ...pointers, ...userEdit];
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

  it("preserves vertical-tab line separators in one paragraph", () => {
    const { document, currentUser } = buildDocument([anchoredTextShape()], ["First\vSecond"]);
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
    ]);
    const child = parsePresentation(fixture).slides![0]!.children![0]!;
    if (!("shape" in child)) throw new TypeError("Expected a shape child");
    expect(child.shape?.textBody?.paragraphs).toEqual([{ children: [{ text: "First\nSecond" }] }]);
  });

  it("projects notes associated by SlideAtom.notesIdRef", () => {
    const { document, currentUser } = buildDocument([anchoredTextShape()], ["Slide text"], {
      notes: true,
    });
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
    ]);
    expect(parsePresentation(fixture).slides![0]!.notes).toBe("Speaker note");
  });

  it("projects legacy shape animation effect, target, trigger, and delay", () => {
    const { document, currentUser } = buildDocument([animatedTextShape()], ["Slide text"], {
      animation: true,
    });
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
    ]);
    const slide = parsePresentation(fixture).slides![0]!;
    expect(slide.animations).toEqual([
      { type: "fade", class: "entrance", trigger: "onClick", delay: 120, shapeId: 2050 },
    ]);
  });

  it("projects legacy external text hyperlinks", () => {
    const { document, currentUser } = buildDocument([hyperlinkTextShape()], ["Slide text"], {
      hyperlink: true,
    });
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
    ]);
    const child = parsePresentation(fixture).slides![0]!.children![0]!;
    if (!("shape" in child)) throw new TypeError("Expected a shape child");
    expect(child.shape?.textBody?.paragraphs?.[0]).toEqual({
      children: [
        {
          text: "Slide text",
          hyperlink: { url: "https://example.com", tooltip: "Tooltip" },
        },
      ],
    });
  });

  it("projects direct OfficeArt picture records and blip references", () => {
    const png = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
    const pictures = record(0xf01e, [...Array.from<number>({ length: 17 }).fill(0), ...png], {
      version: 2,
      instance: 0x6e0,
    });
    const pictureStream = new Uint8Array(4096);
    pictureStream.set(pictures, 0);
    const { document, currentUser } = buildDocument([pictureShape(2051)], ["Picture"]);
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
      { name: "Pictures", data: pictureStream },
    ]);
    const child = parsePresentation(fixture).slides![0]!.children![0]!;
    if (!("picture" in child)) throw new TypeError("Expected a picture child");
    expect(child.picture?.id).toBe(2051);
    expect(child.picture?.x).toBe(317_500);
    expect(child.picture?.y).toBe(158_750);
    expect(child.picture?.width).toBe(476_250);
    expect(child.picture?.height).toBe(476_250);
    expect(child.picture?.type).toBe("png");
    expect(child.picture?.data).toEqual(new Uint8Array(png));
  });

  it("reads 32-bit child anchors in nested groups", () => {
    const { document, currentUser } = buildDocument(
      nestedGroupDrawing([
        childTextShape(4100, 0, 10, 20, 110, 170),
        childTextShape(4101, 1, 20, 40, 120, 190),
      ]),
      ["Alpha", "Beta"],
    );
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
    ]);
    const child = parsePresentation(fixture).slides![0]!.children![0]!;
    if (!("group" in child)) throw new TypeError("Expected a group child");
    expect(child.group?.x).toBe(0);
    expect(child.group?.y).toBe(0);
    expect(child.group?.width).toBe(476_250);
    expect(child.group?.height).toBe(238_125);
    expect(child.group?.children).toHaveLength(2);
    const shape = child.group?.children[0];
    if (!shape || !("shape" in shape)) throw new TypeError("Expected a nested shape");
    expect(shape.shape?.x).toBe(15_875);
    expect(shape.shape?.y).toBe(31_750);
    expect(shape.shape?.width).toBe(158_750);
    expect(shape.shape?.height).toBe(238_125);
  });

  it("projects an embedded BSE picture and trims its name", () => {
    const png = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
    const blip = record(0xf01e, [...Array.from<number>({ length: 17 }).fill(0), ...png], {
      version: 2,
      instance: 0x6e0,
    });
    const bseBody = Array.from<number>({ length: 36 }).fill(0);
    bseBody[33] = 2;
    const bse = record(0xf007, [...bseBody, 0, 0, ...blip], { version: 2, instance: 0x6e0 });
    const pictureStream = new Uint8Array(4096);
    pictureStream.set(container(0xf001, [bse]), 0);
    const { document, currentUser } = buildDocument([pictureShape(2051)], ["Picture"]);
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
      { name: "Pictures", data: pictureStream },
    ]);
    const child = parsePresentation(fixture).slides![0]!.children![0]!;
    if (!("picture" in child)) throw new TypeError("Expected a picture child");
    expect(child.picture?.type).toBe("png");
    expect(child.picture?.data).toEqual(new Uint8Array(png));
  });

  it("keeps text on a shape with an embedded picture fill", () => {
    const png = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
    const blip = record(0xf01e, [...Array.from<number>({ length: 17 }).fill(0), ...png], {
      version: 2,
      instance: 0x6e0,
    });
    const bseBody = Array.from<number>({ length: 36 }).fill(0);
    bseBody[33] = 2;
    const bse = record(0xf007, [...bseBody, 0, 0, ...blip], { version: 2, instance: 0x6e0 });
    const pictureStream = new Uint8Array(4096);
    pictureStream.set(container(0xf001, [bse]), 0);
    const { document, currentUser } = buildDocument([pictureTextShape(2051)], ["Filled"]);
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
      { name: "Pictures", data: pictureStream },
    ]);
    const child = parsePresentation(fixture).slides![0]!.children![0]!;
    if (!("shape" in child)) throw new TypeError("Expected a shape child");
    expect(child.shape?.properties?.fill).toEqual({
      type: "blip",
      data: new Uint8Array(png),
      imageType: "png",
    });
    expect(child.shape?.textBody?.paragraphs).toEqual([{ children: [{ text: "Filled" }] }]);
  });

  it("isolates a corrupt picture beside a valid embedded picture", () => {
    const png = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
    const bseBody = Array.from<number>({ length: 36 }).fill(0);
    bseBody[33] = 0;
    const corruptBlip = record(0xf01f, Array.from<number>({ length: 10 }).fill(0), {
      version: 2,
      instance: 0x7a8,
    });
    const corruptBse = record(0xf007, [...bseBody, ...corruptBlip], { version: 2 });
    const validBlip = record(0xf01e, [...Array.from<number>({ length: 17 }).fill(0), ...png], {
      version: 2,
      instance: 0x6e0,
    });
    bseBody[33] = 2;
    const validBse = record(0xf007, [...bseBody, 0, 0, ...validBlip], {
      version: 2,
      instance: 0x6e0,
    });
    const pictureStream = new Uint8Array(4096);
    pictureStream.set(container(0xf001, [corruptBse, validBse]), 0);
    const { document, currentUser } = buildDocument(
      [pictureShapeWithReference(2051, 2)],
      ["Picture"],
    );
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
      { name: "Pictures", data: pictureStream },
    ]);

    const presentation = parsePresentation(fixture);
    expect(presentation.slides).toHaveLength(1);
    const child = presentation.slides![0]!.children![0]!;
    if (!("picture" in child)) throw new TypeError("Expected a picture child");
    expect(child.picture?.id).toBe(2051);
    expect(child.picture?.type).toBe("png");
    expect(child.picture?.data).toEqual(new Uint8Array(png));
  });

  it("projects a complete grouped-text grid as a table", () => {
    const { document, currentUser } = buildDocument(
      groupedTableDrawing(),
      Array.from({ length: 6 }, (_, index) => `Cell ${index}`),
    );
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
    ]);
    const child = parsePresentation(fixture).slides![0]!.children![0]!;
    if (!("table" in child)) throw new TypeError("Expected a table child");
    expect(child.table?.x).toBe(0);
    expect(child.table?.y).toBe(0);
    expect(child.table?.width).toBe(476_250);
    expect(child.table?.height).toBe(238_125);
    expect(child.table?.columnWidths).toEqual([158_750, 158_750, 158_750]);
    expect(child.table?.rows.map((row) => row.height)).toEqual([79_375, 79_375]);
    expect(child.table?.rows.map((row) => row.cells.map((cell) => cell.text))).toEqual([
      ["Cell 0", "Cell 1", "Cell 2"],
      ["Cell 3", "Cell 4", "Cell 5"],
    ]);
  });

  it("preserves multiple paragraphs in grouped table cells", () => {
    const { document, currentUser } = buildDocument(groupedTableDrawing(), [
      "First\rSecond",
      "Cell 1",
      "Cell 2",
      "Cell 3",
      "Cell 4",
      "Cell 5",
    ]);
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
    ]);
    const child = parsePresentation(fixture).slides![0]!.children![0]!;
    if (!("table" in child)) throw new TypeError("Expected a table child");
    expect(child.table?.rows[0]?.cells[0]).toEqual({
      children: [{ children: [{ text: "First" }] }, { children: [{ text: "Second" }] }],
    });
    expect(child.table?.rows[0]?.cells[1]?.text).toBe("Cell 1");
  });

  it("projects a compact top-level text grid as a table", () => {
    const { document, currentUser } = buildDocument(
      compactGroupedTableDrawing(),
      Array.from({ length: 6 }, (_, index) => `Cell ${index}`),
    );
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
    ]);
    const child = parsePresentation(fixture).slides![0]!.children![0]!;
    if (!("table" in child)) throw new TypeError("Expected a table child");
    expect(child.table?.width).toBe(476_250);
    expect(child.table?.rows).toHaveLength(2);
  });

  it("unwraps a top-level table group wrapper", () => {
    const { document, currentUser } = buildDocument(
      wrappedCompactGroupedTableDrawing(),
      Array.from({ length: 6 }, (_, index) => `Cell ${index}`),
    );
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
    ]);
    const child = parsePresentation(fixture).slides![0]!.children![0]!;
    expect(Object.keys(child)).toEqual(["table"]);
  });

  it("projects merged grouped text rectangles as table spans", () => {
    const { document, currentUser } = buildDocument(mergedTableDrawing(), [
      "Wide",
      "Right",
      "A",
      "B",
      "C",
      "D",
      "Tall",
      "F",
      "G",
      "H",
    ]);
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
    ]);
    const child = parsePresentation(fixture).slides![0]!.children![0]!;
    if (!("table" in child)) throw new TypeError("Expected a table child");
    expect(child.table?.columnWidths).toEqual([158_750, 158_750, 158_750]);
    expect(
      child.table?.rows.map((row) =>
        row.cells.map((cell) => ({
          text: cell.text,
          columnSpan: cell.columnSpan,
          rowSpan: cell.rowSpan,
        })),
      ),
    ).toEqual([
      [
        { text: "Wide", columnSpan: 2, rowSpan: undefined },
        { text: "Right", columnSpan: undefined, rowSpan: undefined },
      ],
      [
        { text: "A", columnSpan: undefined, rowSpan: undefined },
        { text: "B", columnSpan: undefined, rowSpan: undefined },
        { text: "C", columnSpan: undefined, rowSpan: undefined },
      ],
      [
        { text: "D", columnSpan: undefined, rowSpan: undefined },
        { text: "Tall", columnSpan: undefined, rowSpan: 2 },
        { text: "F", columnSpan: undefined, rowSpan: undefined },
      ],
      [
        { text: "G", columnSpan: undefined, rowSpan: undefined },
        { text: "H", columnSpan: undefined, rowSpan: undefined },
      ],
    ]);
  });

  it("maps vertical tabs to line breaks without splitting paragraphs", () => {
    const { document, currentUser } = buildDocument(undefined, ["First\vSecond"]);
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
    ]);
    const child = parsePresentation(fixture).slides![0]!.children![0]!;
    if (!("shape" in child)) throw new TypeError("Expected a shape child");
    expect(child.shape?.textBody?.paragraphs).toEqual([{ children: [{ text: "First\nSecond" }] }]);
  });

  it("maps line feeds to line breaks without splitting paragraphs", () => {
    const { document, currentUser } = buildDocument(undefined, ["First\nSecond"]);
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
    ]);
    const child = parsePresentation(fixture).slides![0]!.children![0]!;
    if (!("shape" in child)) throw new TypeError("Expected a shape child");
    expect(child.shape?.textBody?.paragraphs).toEqual([{ children: [{ text: "First\nSecond" }] }]);
  });

  it("projects unreferenced slide-list text", () => {
    const { document, currentUser } = buildDocument([], ["Orphan"]);
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
    ]);
    const child = parsePresentation(fixture).slides![0]!.children![0]!;
    if (!("shape" in child)) throw new TypeError("Expected a shape child");
    expect(child.shape?.textBody?.paragraphs).toEqual([{ children: [{ text: "Orphan" }] }]);
  });

  it("resolves date and slide-number field placeholders", () => {
    const { document, currentUser } = buildDocument(
      [fieldShape(2051, 4006), fieldShape(2052, 4056)],
      [],
    );
    const fixture = buildCfb([
      { name: "PowerPoint Document", data: document },
      { name: "Current User", data: currentUser },
    ]);
    const children = parsePresentation(fixture).slides![0]!.children!;
    expect(
      children.map((child) => {
        if (!("shape" in child)) throw new TypeError("Expected shape children");
        const paragraph = child.shape?.textBody?.paragraphs?.[0];
        if (typeof paragraph !== "object") throw new TypeError("Expected a paragraph object");
        const run = paragraph.children?.[0];
        if (typeof run !== "object" || !("text" in run)) throw new TypeError("Expected a text run");
        return run.text;
      }),
    ).toEqual(["1/1/1", "1"]);
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
