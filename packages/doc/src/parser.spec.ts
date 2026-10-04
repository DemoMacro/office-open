import type { ParagraphOptions, SectionChild } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { DocParseError } from "./errors";
import { parseDocument } from "./index";
import type { LegacyDocumentOptions } from "./records/models";

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

function parseLegacyDocument(data: Uint8Array): LegacyDocumentOptions {
  return parseDocument(data) as LegacyDocumentOptions;
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

function buildDocument(
  options: {
    characterBinTableLength?: number;
    secondCp?: number;
    fields?: boolean;
    hyperlink?: boolean;
    annotations?: boolean;
    bookmarks?: boolean;
    drawing?: boolean | "embedded";
    fieldTable?: boolean;
    lists?: boolean;
    paragraphGrpprl?: readonly number[];
    sectionProperties?: boolean;
    stylesheet?: boolean;
    textboxes?: { malformed?: boolean; reusableFirst?: boolean };
    notes?: boolean;
    headers?: boolean;
  } = {},
): {
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
  const storyFixture = options.textboxes ?? options.notes ?? options.headers;
  const fieldInstruction = options.hyperlink ? 'HYPERLINK "https://example.com"' : "INST";
  const fieldBytes =
    options.fields || options.hyperlink
      ? new Uint8Array([
          0x41,
          0x13,
          ...new TextEncoder().encode(fieldInstruction),
          0x14,
          ...new TextEncoder().encode("RES"),
          0x15,
          0x0d,
        ])
      : undefined;
  wordView.setUint32(76, fieldBytes?.byteLength ?? (storyFixture ? 4 : 7), true);
  wordView.setUint32(80, options.notes ? 4 : 0, true);
  wordView.setUint32(84, options.headers ? 4 : 0, true);
  wordView.setUint32(100, options.textboxes ? 4 : 0, true);
  wordView.setUint16(152, 75, true);

  const setFibPair = (index: number, offset: number, length: number): void => {
    wordView.setUint32(154 + index * 8, offset, true);
    wordView.setUint32(154 + index * 8 + 4, length, true);
  };
  setFibPair(12, 0, options.characterBinTableLength ?? 12);
  setFibPair(13, 12, 12);
  setFibPair(33, 512, 33);
  setFibPair(56, 640, options.textboxes?.malformed ? 57 : 56);
  setFibPair(3, 720, options.notes ? 16 : 0);
  setFibPair(11, 720, options.headers ? 56 : 0);
  setFibPair(1, options.stylesheet ? 768 : 0, options.stylesheet ? 28 : 0);
  setFibPair(4, options.annotations ? 944 : 0, options.annotations ? 8 : 0);
  setFibPair(6, options.sectionProperties ? 800 : 0, options.sectionProperties ? 20 : 0);
  setFibPair(16, options.fieldTable ? 912 : 0, options.fieldTable ? 16 : 0);
  setFibPair(21, options.bookmarks ? 832 : 0, options.bookmarks ? 8 : 0);
  setFibPair(22, options.bookmarks ? 880 : 0, options.bookmarks ? 12 : 0);
  setFibPair(23, options.bookmarks ? 896 : 0, options.bookmarks ? 8 : 0);
  setFibPair(
    50,
    options.drawing ? 1024 : 0,
    options.drawing ? (options.drawing === "embedded" ? 114 : 60) : 0,
  );
  setFibPair(73, options.lists ? 1216 : 0, options.lists ? 30 : 0);
  setFibPair(74, options.lists ? 1280 : 0, options.lists ? 10 : 0);

  word.set(new TextEncoder().encode("Hi\r"), 512);
  if (fieldBytes) {
    word.set(fieldBytes, 512);
  }
  const unicode = new Uint8Array(storyFixture ? 8 : 6);
  const unicodeView = new DataView(unicode.buffer);
  if (storyFixture) {
    unicodeView.setUint16(0, 0x0041, true);
    unicodeView.setUint16(2, 0x000d, true);
    unicodeView.setUint16(4, 0x0042, true);
    unicodeView.setUint16(6, 0x000d, true);
  } else {
    unicodeView.setUint16(0, 0x4e16, true);
    unicodeView.setUint16(2, 0x754c, true);
    unicodeView.setUint16(4, 0x000d, true);
  }
  word.set(unicode, fieldBytes ? 576 : 516);

  tableView.setUint8(512, 2);
  tableView.setUint32(513, 28, true);
  tableView.setUint32(517, 0, true);
  tableView.setUint32(521, options.secondCp ?? (fieldBytes ? fieldBytes.byteLength : 4), true);
  tableView.setUint32(525, fieldBytes ? fieldBytes.byteLength + 3 : storyFixture ? 8 : 7, true);
  tableView.setUint32(531, 0x40000400, true);
  tableView.setUint32(539, fieldBytes ? 576 : 516, true);

  tableView.setUint32(0, 512, true);
  tableView.setUint32(4, 528, true);
  tableView.setUint32(8, 2, true);
  tableView.setUint32(12, 512, true);
  tableView.setUint32(16, 528, true);
  tableView.setUint32(20, 3, true);

  if (options.textboxes) {
    tableView.setUint32(640, 0, true);
    tableView.setUint32(644, 2, true);
    tableView.setUint32(648, 4, true);
    tableView.setUint16(660, options.textboxes.reusableFirst ? 1 : 0);
    tableView.setUint16(682, 1);
  }

  if (options.notes) {
    [0, 2, 3, 0].forEach((value, index) => tableView.setUint32(720 + index * 4, value, true));
  }

  if (options.headers) {
    [0, 0, 0, 0, 0, 0, 0, 0, 3, 3, 3, 3, 3, 0].forEach((value, index) =>
      tableView.setUint32(720 + index * 4, value, true),
    );
  }

  const characterPage = 1024;
  const paragraphPage = 1536;
  wordView.setUint32(characterPage, 512, true);
  wordView.setUint32(characterPage + 4, 528, true);
  wordView.setUint8(characterPage + 8, 20);
  wordView.setUint8(characterPage + 511, 1);
  const grpprl = [
    0x35, 0x08, 0x01, 0x00, 0xc0, 0x02, 0xaa, 0xbb, 0x43, 0x4a, 0x18, 0x00, 0x00, 0x00, 0x00,
  ];
  word.set(Uint8Array.from(grpprl), characterPage + 41);
  wordView.setUint8(characterPage + 40, grpprl.length);

  if (options.paragraphGrpprl) {
    const papxGrpprl = Uint8Array.from(options.paragraphGrpprl);
    const papxCb = Math.ceil(papxGrpprl.byteLength / 2) + 1;
    const paddedGrpprl =
      papxGrpprl.byteLength % 2 === 0 ? papxGrpprl : Uint8Array.from([...papxGrpprl, 0]);
    word.set(Uint8Array.from([0, papxCb, 0, 0, ...paddedGrpprl, 0, 0]), paragraphPage + 40);
  }

  if (options.sectionProperties) {
    tableView.setUint32(800, 0, true);
    tableView.setUint32(804, 1, true);
    tableView.setUint32(808, 0, true);
    tableView.setUint32(810, 2048, true);
    tableView.setUint32(816, 0, true);
    word.set(
      Uint8Array.from([
        24, 0, 0x1f, 0xf6, 0xa0, 0x0f, 0x20, 0xf6, 0xd0, 0x07, 0x21, 0xf6, 0x40, 0x00, 0x22, 0xf6,
        0x80, 0x00, 0x23, 0xf6, 0x60, 0x00,
      ]),
      2048,
    );
  }

  if (options.stylesheet) {
    tableView.setUint16(768, 1, true);
    tableView.setUint8(772, 0);
    tableView.setUint8(773, 0);
    tableView.setUint16(774, 0, true);
    tableView.setUint16(776, 4095, true);
    tableView.setUint16(778, 10, true);
    tableView.setUint16(780, 0, true);
    tableView.setUint16(782, 6, true);
    "Normal"
      .split("")
      .forEach((character, index) =>
        tableView.setUint16(784 + index * 2, character.charCodeAt(0), true),
      );
  }

  if (options.bookmarks) {
    tableView.setUint32(832, 1, true);
    tableView.setUint16(836, 4, true);
    "Mark".split("").forEach((character, index) => {
      tableView.setUint8(838 + index, character.charCodeAt(0));
    });
    tableView.setUint32(880, 0, true);
    tableView.setUint32(884, 3, true);
    tableView.setUint32(888, 0, true);
    tableView.setUint32(896, 0, true);
    tableView.setUint32(900, 3, true);
  }

  if (options.fieldTable) {
    [1, 2, 3, 4].forEach((value, index) => tableView.setUint32(912 + index * 4, value, true));
    tableView.setUint8(924, 0x13);
    tableView.setUint8(925, 0x21);
    tableView.setUint8(926, 0x14);
    tableView.setUint8(927, 0xff);
    tableView.setUint8(928, 0x15);
    tableView.setUint8(929, 0xff);
  }

  if (options.annotations) {
    [0, 3].forEach((value, index) => tableView.setUint32(944 + index * 4, value, true));
    tableView.setUint8(952, 0);
    tableView.setUint8(953, 0);
  }

  if (options.drawing) {
    const drawingView = new DataView(table.buffer);
    const embedded = options.drawing === "embedded";
    drawingView.setUint16(1024, 0x000f, true);
    drawingView.setUint16(1026, 0xf000, true);
    drawingView.setUint32(1028, embedded ? 106 : 52, true);
    drawingView.setUint16(1032, 0x000f, true);
    drawingView.setUint16(1034, 0xf001, true);
    drawingView.setUint32(1036, embedded ? 98 : 44, true);
    drawingView.setUint16(1040, embedded ? 0x06e0 : 0x0000, true);
    drawingView.setUint16(1042, 0xf007, true);
    drawingView.setUint32(1044, embedded ? 90 : 36, true);
    tableView.setUint8(1058, 1);
    drawingView.setUint32(1076, embedded ? 0xffffffff : 0, true);
    if (embedded) {
      drawingView.setUint16(1084, 0x06e0, true);
      drawingView.setUint16(1086, 0xf01e, true);
      drawingView.setUint32(1088, 46, true);
      drawingView.setUint8(1108, 0xff);
      [
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0,
        0, 4, 0, 0, 0, 3, 8, 6, 0, 0, 0,
      ].forEach((byte, index) => {
        drawingView.setUint8(1109 + index, byte);
      });
    }
  }

  if (options.lists) {
    tableView.setUint16(1216, 1, true);
    tableView.setUint32(1218, 100, true);
    tableView.setUint8(1244, 1);
    tableView.setUint32(1280, 1, true);
    tableView.setUint32(1284, 100, true);
    tableView.setUint8(1288, 0);
  }

  wordView.setUint32(paragraphPage, 512, true);
  wordView.setUint32(paragraphPage + 4, 528, true);
  wordView.setUint8(paragraphPage + 8, 20);
  wordView.setUint8(paragraphPage + 511, 1);
  if (!options.paragraphGrpprl) {
    word.set(
      Uint8Array.from([
        0,
        4, // cb = 0, cb' = 4 (8-byte GrpPrlAndIstd)
        0,
        0, // istd
        0x00,
        0x24,
        0x01, // GrpPrl
        0x00,
        0x24,
        0x01, // GrpPrl
      ]),
      paragraphPage + 40,
    );
  }

  const data = buildContainer([
    { path: "WordDocument", data: word },
    { path: "1Table", data: table },
    ...(options.drawing
      ? [
          {
            path: "Data",
            data: options.drawing === "embedded" ? new Uint8Array() : drawingStream(),
          },
        ]
      : []),
  ]);
  return { data, word, table };
}

function drawingStream(): Uint8Array {
  const data = new Uint8Array(STREAM_SIZE);
  const view = new DataView(data.buffer);
  view.setUint16(0, 0x06e0, true);
  view.setUint16(2, 0xf01e, true);
  view.setUint32(4, 46, true);
  view.setUint8(24, 0xff);
  data.set(
    Uint8Array.from([
      0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0,
      4, 0, 0, 0, 3, 8, 6, 0, 0, 0,
    ]),
    25,
  );
  return data;
}

function buildLegacyDocument(): Uint8Array {
  const word = new Uint8Array(STREAM_SIZE);
  const view = new DataView(word.buffer);
  view.setUint16(0, 0xa5dc, true);
  view.setUint16(2, 101, true);
  view.setUint32(24, 768, true);
  view.setUint32(28, 812, true);
  view.setUint32(52, 44, true);
  word.set(new TextEncoder().encode("The quick brown fox jumps over the lazy dog\r"), 768);
  return buildContainer([{ path: "WordDocument", data: word }]);
}

describe("legacy DOC parser", () => {
  it("reports structured low-level parse failures", () => {
    const actions: readonly (readonly [string, () => unknown])[] = [
      ["truncated-header", () => parseDocument(buildDocument().data.slice(0, 256))],
      [
        "bad-record-length",
        () => {
          const { word, table } = buildDocument();
          new DataView(table.buffer).setUint32(513, 0xffffffff, true);
          return parseDocument(
            buildContainer([
              { path: "WordDocument", data: word },
              { path: "1Table", data: table },
            ]),
          );
        },
      ],
      [
        "invalid-container-traversal",
        () => {
          const data = new Uint8Array(buildDocument().data);
          new DataView(data.buffer).setUint32(1024 + 76, 0xfffffff0, true);
          return parseDocument(data);
        },
      ],
      [
        "impossible-offset",
        () => {
          const { table, word } = buildDocument();
          new DataView(table.buffer).setUint32(531, 0x7fffffff, true);
          return parseDocument(
            buildContainer([
              { path: "WordDocument", data: word },
              { path: "1Table", data: table },
            ]),
          );
        },
      ],
      [
        "unsupported-required-structure",
        () => {
          const { word, table } = buildDocument();
          table[512] = 9;
          return parseDocument(
            buildContainer([
              { path: "WordDocument", data: word },
              { path: "1Table", data: table },
            ]),
          );
        },
      ],
      [
        "encrypted-unsupported",
        () => {
          const { word, table } = buildDocument();
          const view = new DataView(word.buffer);
          view.setUint16(10, view.getUint16(10, true) | 0x0100, true);
          return parseDocument(
            buildContainer([
              { path: "WordDocument", data: word },
              { path: "1Table", data: table },
            ]),
          );
        },
      ],
    ];

    for (const [reason, action] of actions) {
      let error: DocParseError | undefined;
      try {
        action();
      } catch (thrown) {
        error = thrown as DocParseError;
      }
      expect(error, reason).toBeInstanceOf(DocParseError);
      expect(error!.context.format, reason).toBe("doc");
      expect(error!.context.reason, reason).toMatch(
        /invalid-container|invalid-record-length|invalid-file-character-position|out-of-range|unknown-required-record|unsupported-required-structure|encrypted-unsupported/,
      );
      expect(error!.context.path, reason).toBeDefined();
    }

    try {
      parseDocument(buildDocument().data.slice(0, 256));
    } catch (error) {
      expect((error as DocParseError).context).toMatchObject({
        format: "doc",
        part: "container",
        path: "/",
      });
    }
  });

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

  it("keeps field results without field instructions", () => {
    const { data } = buildDocument({ fields: true });
    const children = parseDocument(data).sections[0]!.children;
    const first = children[0]!;
    if (!("paragraph" in first)) throw new TypeError("Expected a paragraph child");
    const paragraph = first.paragraph;
    const text =
      typeof paragraph === "string"
        ? paragraph
        : (paragraph?.children ?? [])
            .map((child: unknown) => {
              if (typeof child === "string") return child;
              return typeof child === "object" &&
                child !== null &&
                "text" in child &&
                typeof child.text === "string"
                ? child.text
                : "";
            })
            .join("");
    expect(text).toBe("ARES");
    expect(JSON.stringify(children)).not.toContain("INST");
  });

  it("projects HYPERLINK field results as hyperlink children", () => {
    const { data } = buildDocument({ hyperlink: true, fieldTable: true });
    const children = parseDocument(data).sections[0]!.children;
    const first = children[0]!;
    if (!("paragraph" in first) || typeof first.paragraph === "string")
      throw new TypeError("Expected a paragraph child");
    expect(first.paragraph.children).toEqual([
      { text: "A", bold: true, size: 12 },
      {
        hyperlink: { url: "https://example.com" },
        children: [{ text: "RES" }],
      },
    ]);
  });

  it("degrades malformed character bin tables without losing text", () => {
    const { data } = buildDocument({ characterBinTableLength: 13 });
    const children = parseDocument(data).sections[0]!.children;
    expect(children).toHaveLength(2);
  });

  it("projects actual textbox boundaries and skips reusable FTXBXS records", () => {
    const { data } = buildDocument({ textboxes: {} });
    const children = parseDocument(data).sections[0]!.children;
    const textbox = children.at(-1);
    if (!textbox || !("textbox" in textbox)) throw new TypeError("Expected a textbox child");

    expect(children).toHaveLength(2);
    expect(textbox.textbox.children).toEqual([
      { paragraph: { children: [{ text: "A", bold: true, size: 12 }] } },
    ]);
  });

  it("ignores an FTXBXS record marked reusable before the final spare", () => {
    const { data } = buildDocument({ textboxes: { reusableFirst: true } });
    const children = parseDocument(data).sections[0]!.children;

    expect(children).toHaveLength(1);
    expect("textbox" in children[0]!).toBe(false);
  });

  it("degrades malformed textbox boundary tables without losing text", () => {
    const { data } = buildDocument({ textboxes: { malformed: true } });
    const children = parseDocument(data).sections[0]!.children;
    expect(children).toHaveLength(1);
    expect("textbox" in children[0]!).toBe(false);
  });

  it("projects individual footnote boundaries", () => {
    const { data } = buildDocument({ notes: true });
    const document = parseDocument(data);

    expect(document.footnotes).toEqual([
      { id: 1, children: [{ paragraph: { children: [{ text: "A", bold: true, size: 12 }] } }] },
      { id: 2, children: [{ paragraph: { children: [{ text: "B", bold: true, size: 12 }] } }] },
    ]);
  });

  it("projects the default header story slot", () => {
    const { data } = buildDocument({ headers: true });
    const section = parseDocument(data).sections[0]!;

    expect(section.headers?.default).toEqual([
      { paragraph: { children: [{ text: "A", bold: true, size: 12 }] } },
      { paragraph: { children: [{ text: "B", bold: true, size: 12 }] } },
    ]);
    expect(section.footers).toBeUndefined();
  });

  it("decodes full CHPX run formatting", () => {
    const { data } = buildDocument({
      characterBinTableLength: 12,
      paragraphGrpprl: [0x03, 0x24, 0x01],
    });
    const paragraph = parseDocument(data).sections[0]!.children[0]!;
    const run = extractParagraph(paragraph).children?.[0];
    expect(run).toMatchObject({ text: "Hi", bold: true, size: 12 });
  });

  it("decodes PAPX alignment, indentation, spacing and outline level", () => {
    const { data } = buildDocument({
      paragraphGrpprl: [
        0x03, 0x24, 0x01, 0x5e, 0x84, 0xd0, 0x02, 0x5d, 0x84, 0x40, 0x00, 0x11, 0x84, 0x00, 0x00,
        0x13, 0xa4, 0x20, 0x00, 0x40, 0x24, 0x01,
      ],
    });
    const paragraph = extractParagraph(parseDocument(data).sections[0]!.children[0]!);
    expect(paragraph.alignment).toBe("center");
    expect(paragraph.indent).toEqual({ left: 720, right: 64, firstLine: 0 });
    expect(paragraph.spacing).toEqual({ before: 32 });
    expect(paragraph.outlineLevel).toBe(1);
  });

  it("maps SEPX geometry to DocumentOptions.sections", () => {
    const { data } = buildDocument({ sectionProperties: true });
    const section = parseDocument(data).sections[0]!;
    expect(section.properties?.pageSize).toMatchObject({ width: 4000, height: 2000 });
    expect(section.properties?.pageMargin).toMatchObject({ left: 64, right: 128, top: 96 });
  });

  it("builds a table row from in-table PAPX and TAP", () => {
    const { data } = buildDocument({
      paragraphGrpprl: [
        0x16, 0x24, 0x01, 0x17, 0x24, 0x01, 0x08, 0xd6, 0x0d, 0x02, 0x00, 0x00, 0x00, 0x40, 0x00,
        0x80, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
      ],
    });
    const children = parseDocument(data).sections[0]!.children;
    expect(children[0]).toEqual({
      table: {
        columnWidths: [64, 64],
        rows: [
          {
            cells: [
              {
                children: [
                  expect.objectContaining({
                    paragraph: { children: [{ text: "Hi", bold: true, size: 12 }] },
                  }),
                ],
              },
              {
                children: [
                  expect.objectContaining({
                    paragraph: { children: [{ text: "世界", bold: true, size: 12 }] },
                  }),
                ],
              },
            ],
          },
        ],
      },
    });
  });

  it("projects bookmark names and ranges", () => {
    const { data } = buildDocument({ bookmarks: true });
    const children = parseDocument(data).sections[0]!.children;
    expect(children[0]).toEqual({ bookmarkStart: { id: 1, name: "Mark" } });
    expect(children.at(-1)).toEqual({ bookmarkEnd: { id: 1 } });
  });

  it("rejects revisions that cannot preserve required change metadata", () => {
    const { data } = buildDocument({ annotations: true });
    expect(() => parseLegacyDocument(data)).toThrow(DocParseError);
  });

  it("projects main-document field instructions into canonical field runs", () => {
    const { data } = buildDocument({ fieldTable: true });
    const children = parseLegacyDocument(data).sections[0]!.children;
    expect(children[0]).toMatchObject({
      paragraph: { children: [{ complexField: { instruction: expect.any(String) } }] },
    });
  });

  it("projects Escher BStore pictures into canonical inline pictures", () => {
    const { data } = buildDocument({ drawing: true });
    const children = parseLegacyDocument(data).sections[0]!.children;
    expect(children[0]).toMatchObject({
      paragraph: {
        children: [{ picture: { type: "png", transformation: { width: 4, height: 3 } } }],
      },
    });
  });

  it("projects embedded BLIP pictures without consulting the Data stream", () => {
    const { data } = buildDocument({ drawing: "embedded" });
    const children = parseLegacyDocument(data).sections[0]!.children;
    expect(children[0]).toMatchObject({
      paragraph: { children: [{ picture: { type: "png" } }] },
    });
  });

  it("converts PlfLst and PlfLfo to numbering definitions", () => {
    const { data } = buildDocument({ lists: true });
    expect(parseDocument(data).numbering?.abstractNumberings).toHaveLength(1);
    expect(parseDocument(data).numbering?.abstractNumberings[0]!.levels).toHaveLength(1);
  });

  it("converts STSH style names and based-on relationships", () => {
    const { data } = buildDocument({ stylesheet: true });
    expect(parseDocument(data).styles?.paragraphStyles).toEqual([
      { id: "style-0", name: "Normal", basedOn: undefined },
    ]);
  });

  it("extracts Word 6.0/95 text without modern FIB extensions", () => {
    const document = parseDocument(buildLegacyDocument());
    expect(document.sections[0]!.children).toEqual([
      { paragraph: "The quick brown fox jumps over the lazy dog" },
      { paragraph: "" },
    ]);
  });
});

function extractParagraph(child: SectionChild): ParagraphOptions {
  if (!("paragraph" in child) || typeof child.paragraph === "string") {
    throw new TypeError("Expected a paragraph object");
  }
  return child.paragraph;
}
