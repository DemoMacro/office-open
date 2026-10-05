import { OOXML_PACKAGE_FORMATS, type OoxmlPackageFormat } from "@office-open/core";
import { zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import {
  generate,
  generateOfficeDocument,
  type GenerateOptionsMap,
  type GenerateType,
} from "./generate";
import { parseOfficeDocument } from "./parse";

const SECTOR_SIZE = 512;
const CFB_SIGNATURE = Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
const END_OF_CHAIN = 0xfffffffe;
const FAT_SECTOR = 0xfffffffd;
const FREE_SECTOR = 0xffffffff;
const NO_STREAM = 0xffffffff;

const OOXML_FIXTURES = {
  wordprocessing: { sections: [{ children: [{ paragraph: "Dispatch" }] }] },
  presentation: {
    slides: [
      {
        children: [{ shape: { x: 0, y: 0, width: 10, height: 4, textBody: { text: "Dispatch" } } }],
      },
    ],
  },
  spreadsheet: { worksheets: [{ rows: [{ cells: [{ value: "Dispatch" }] }] }] },
};

const ODF_FIXTURES = {
  odt: OOXML_FIXTURES.wordprocessing,
  ods: OOXML_FIXTURES.spreadsheet,
  odp: OOXML_FIXTURES.presentation,
};

const OUTPUT_TYPES = [
  "base64",
  "string",
  "text",
  "binarystring",
  "array",
  "uint8array",
  "arraybuffer",
  "blob",
  "nodebuffer",
] as const;

function concatenate(parts: readonly Uint8Array[]): Uint8Array {
  const bytes = new Uint8Array(parts.reduce((total, part) => total + part.byteLength, 0));
  let offset = 0;
  for (const part of parts) {
    bytes.set(part, offset);
    offset += part.byteLength;
  }
  return bytes;
}

function uint16Record(code: number, body: Uint8Array = new Uint8Array()): Uint8Array {
  const bytes = new Uint8Array(4 + body.byteLength);
  const view = new DataView(bytes.buffer);
  view.setUint16(0, code, true);
  view.setUint16(2, body.byteLength, true);
  bytes.set(body, 4);
  return bytes;
}

function uint16Body(values: readonly number[]): Uint8Array {
  const bytes = new Uint8Array(values.length * 2);
  const view = new DataView(bytes.buffer);
  values.forEach((value, index) => view.setUint16(index * 2, value, true));
  return bytes;
}

function biff5String(value: string): Uint8Array {
  const bytes = new Uint8Array(1 + value.length);
  bytes[0] = value.length;
  for (const [index, character] of Array.from(value).entries()) {
    bytes[index + 1] = character.charCodeAt(0);
  }
  return bytes;
}

function biff5LabelString(value: string): Uint8Array {
  const bytes = new Uint8Array(2 + value.length);
  new DataView(bytes.buffer).setUint16(0, value.length, true);
  for (const [index, character] of Array.from(value).entries()) {
    bytes[index + 2] = character.charCodeAt(0);
  }
  return bytes;
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
  bytes[offset + 67] = 1;
  view.setUint32(offset + 68, NO_STREAM, true);
  view.setUint32(offset + 72, rightSiblingId, true);
  view.setUint32(offset + 76, childId, true);
  view.setUint32(offset + 116, startSector, true);
  view.setUint32(offset + 120, size, true);
}

function buildCfb(streams: readonly { name: string; data: Uint8Array }[]): Uint8Array {
  const sectorCounts = streams.map((stream) =>
    Math.max(1, Math.ceil(stream.data.byteLength / SECTOR_SIZE)),
  );
  const sectorCount = 2 + sectorCounts.reduce((total, count) => total + count, 0);
  const bytes = new Uint8Array(SECTOR_SIZE + sectorCount * SECTOR_SIZE);
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
  let sector = 2;
  const startSectors = streams.map((stream, index) => {
    const start = sector;
    const count = sectorCounts[index]!;
    for (let current = start; current < start + count - 1; current += 1) {
      fat[current] = current + 1;
    }
    fat[start + count - 1] = END_OF_CHAIN;
    bytes.set(stream.data, SECTOR_SIZE + start * SECTOR_SIZE);
    sector += count;
    return start;
  });
  fat.forEach((value, current) => view.setUint32(SECTOR_SIZE + current * 4, value, true));
  for (let index = 1; index < 109; index += 1) {
    view.setUint32(76 + index * 4, FREE_SECTOR, true);
  }
  writeDirectoryEntry(bytes, SECTOR_SIZE * 2, "Root Entry", 5, NO_STREAM, 1, END_OF_CHAIN, 0);
  streams.forEach((stream, index) => {
    writeDirectoryEntry(
      bytes,
      SECTOR_SIZE * 2 + (index + 1) * 128,
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

function legacyDocument(): Uint8Array {
  const word = new Uint8Array(SECTOR_SIZE * 8);
  const view = new DataView(word.buffer);
  view.setUint16(0, 0xa5dc, true);
  view.setUint16(2, 101, true);
  view.setUint32(24, 768, true);
  view.setUint32(28, 812, true);
  view.setUint32(52, 44, true);
  word.set(new TextEncoder().encode("Legacy DOC dispatch\r"), 768);
  return buildCfb([{ name: "WordDocument", data: word }]);
}

function legacyWorkbook(): Uint8Array {
  const begin = uint16Record(0x0809, Uint8Array.from([0, 5, 0x10, 0]));
  const eof = uint16Record(0x000a);
  const numberCell = uint16Record(
    0x0203,
    concatenate([
      uint16Body([0, 0, 0]),
      (() => {
        const value = new Uint8Array(8);
        new DataView(value.buffer).setFloat64(0, 12.5, true);
        return value;
      })(),
    ]),
  );
  const labelCell = uint16Record(
    0x0204,
    concatenate([uint16Body([0, 1, 1]), biff5LabelString("XLS dispatch")]),
  );
  const worksheet = concatenate([begin, numberCell, labelCell, eof]);
  const dimensions = uint16Record(0x0200, uint16Body([0, 0, 1, 1, 0]));
  const boundSheet = uint16Record(
    0x0085,
    concatenate([
      (() => {
        const body = new Uint8Array(6);
        new DataView(body.buffer).setUint32(0, begin.byteLength + 16 + eof.byteLength, true);
        return body;
      })(),
      biff5String("Sheet"),
    ]),
  );
  const workbook = concatenate([begin, boundSheet, eof, dimensions, worksheet]);
  const padded = new Uint8Array(SECTOR_SIZE * 8);
  padded.set(workbook);
  return buildCfb([{ name: "Workbook", data: padded }]);
}

function presentationRecord(
  type: number,
  body: readonly number[],
  options: { instance?: number; version?: number } = {},
): number[] {
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

function littleEndian32(value: number): number[] {
  return [value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff];
}

function presentationContainer(
  type: number,
  children: readonly number[][],
  instance = 0,
): number[] {
  return presentationRecord(type, children.flat(), { version: 15, instance });
}

function legacyPresentation(): Uint8Array {
  const textShape = presentationContainer(0xf004, [
    presentationRecord(0xf00a, [...littleEndian32(2050), ...littleEndian32(544)]),
    presentationRecord(0xf010, [
      ...(() => [0, 0x02, 0x40, 0x00, 0, 0x03, 0x00, 0])(),
      ...(() => [0, 0x04, 0x00, 0x00, 0, 0x05, 0x00, 0])(),
    ]),
    presentationContainer(0xf00d, [presentationRecord(3998, littleEndian32(0))]),
  ]);
  const slideList = presentationContainer(4080, [
    presentationRecord(1011, [
      ...littleEndian32(2),
      ...littleEndian32(4),
      ...littleEndian32(1),
      ...littleEndian32(257),
      ...littleEndian32(0),
      ...littleEndian32(0),
    ]),
    presentationRecord(3999, littleEndian32(0)),
    presentationRecord(4008, Array.from(new TextEncoder().encode("PPT dispatch"))),
  ]);
  const document = presentationContainer(1000, [
    presentationRecord(1001, [
      ...littleEndian32(5760),
      ...littleEndian32(4320),
      ...Array.from<number>({ length: 40 }).fill(0),
    ]),
    slideList,
  ]);
  const slide = presentationContainer(1006, [
    presentationRecord(1007, [
      ...littleEndian32(5760),
      ...littleEndian32(4320),
      ...littleEndian32(0),
      ...littleEndian32(0),
      ...littleEndian32(0),
      ...littleEndian32(0),
      0,
      0,
      0,
      0,
    ]),
    presentationContainer(1036, [
      presentationContainer(0xf002, [presentationContainer(0xf003, [textShape])]),
    ]),
  ]);
  const pointers = presentationRecord(6002, [
    ...littleEndian32(((1 & 0xfff) << 20) | 1),
    ...littleEndian32(0),
    ...littleEndian32(((1 & 0xfff) << 20) | 2),
    ...littleEndian32(document.length),
  ]);
  const userEdit = presentationRecord(4085, [
    ...littleEndian32(1),
    ...littleEndian32(8),
    ...littleEndian32(0),
    ...littleEndian32(document.length + slide.length),
    ...littleEndian32(1),
    ...littleEndian32(2),
    1,
    0,
    0,
    0,
  ]);
  const presentationBytes = Uint8Array.from([...document, ...slide, ...pointers, ...userEdit]);
  const documentStream = new Uint8Array(SECTOR_SIZE * 8);
  documentStream.set(presentationBytes);
  const currentUser = new Uint8Array(SECTOR_SIZE * 8);
  currentUser.set(
    Uint8Array.from(
      presentationRecord(4086, [
        ...littleEndian32(20),
        ...littleEndian32(0xe391c05f),
        ...littleEndian32(document.length + slide.length + pointers.length),
        0,
        0,
        0xf4,
        0x03,
        3,
        0,
        ...littleEndian32(8),
      ]),
    ),
  );
  return buildCfb([
    { name: "PowerPoint Document", data: documentStream },
    { name: "Current User", data: currentUser },
  ]);
}

describe("format dispatch", () => {
  it("round-trips every writable OOXML package format", async () => {
    for (const type of Object.keys(OOXML_PACKAGE_FORMATS) as OoxmlPackageFormat[]) {
      const bytes = (await generate({
        type,
        options: OOXML_FIXTURES[
          OOXML_PACKAGE_FORMATS[type].family
        ] as GenerateOptionsMap[GenerateType],
        outputType: "uint8array",
      })) as Uint8Array;
      const parsed = await parseOfficeDocument(bytes);
      expect(parsed.type).toBe(type);
    }
  });

  it("round-trips all ODF root formats", async () => {
    for (const type of ["odt", "ods", "odp"] as const) {
      const bytes = (await generateOfficeDocument(
        type,
        ODF_FIXTURES[type],
        "uint8array",
      )) as Uint8Array;
      const parsed = await parseOfficeDocument(bytes);
      expect(parsed.type).toBe(type);
      expect(
        "sections" in parsed.options ||
          "worksheets" in parsed.options ||
          "slides" in parsed.options,
      ).toBe(true);
    }
  });

  it("generates and parses the RTF root format", async () => {
    const bytes = (await generateOfficeDocument(
      "rtf",
      ODF_FIXTURES.odt,
      "uint8array",
    )) as Uint8Array;
    expect(new TextDecoder().decode(bytes)).toContain("Dispatch");
    const parsed = await parseOfficeDocument(bytes);
    expect(parsed.type).toBe("rtf");
    const options = parsed.options as { sections?: { children: unknown[] }[] };
    expect(options.sections?.[0]?.children[0]).toEqual({
      paragraph: { children: [{ text: "Dispatch" }] },
    });
  });

  it("dispatches legacy parse-only roots to family parsers", async () => {
    const document = await parseOfficeDocument(legacyDocument());
    const workbook = await parseOfficeDocument(legacyWorkbook());
    const presentation = await parseOfficeDocument(legacyPresentation());
    expect([document.type, workbook.type, presentation.type]).toEqual(["doc", "xls", "ppt"]);
    if (document.type !== "doc" || workbook.type !== "xls" || presentation.type !== "ppt") {
      throw new Error("Unexpected legacy dispatch types");
    }
    expect(document.options.sections).toHaveLength(1);
    expect(workbook.options.worksheets).toHaveLength(1);
    expect(presentation.options.slides).toHaveLength(1);
  });

  it.each(OUTPUT_TYPES)("returns the declared %s output type", async (outputType) => {
    const ooxml = await generate({ type: "docx", options: ODF_FIXTURES.odt, outputType });
    const odf = await generateOfficeDocument("odt", ODF_FIXTURES.odt, outputType);
    if (outputType === "base64") {
      expect(typeof ooxml).toBe("string");
      expect(typeof odf).toBe("string");
    } else if (outputType === "array") {
      expect(ooxml).toBeTypeOf("object");
      expect(odf).toBeTypeOf("object");
    } else if (outputType === "uint8array") {
      expect(ooxml).toBeInstanceOf(Uint8Array);
      expect(odf).toBeInstanceOf(Uint8Array);
    } else if (outputType === "arraybuffer") {
      expect(ooxml).toBeInstanceOf(ArrayBuffer);
      expect(odf).toBeInstanceOf(ArrayBuffer);
    } else if (outputType === "blob") {
      expect(ooxml).toBeInstanceOf(Blob);
      expect(odf).toBeInstanceOf(Blob);
    } else if (outputType === "nodebuffer") {
      expect(ooxml).toBeInstanceOf(Uint8Array);
      expect(odf).toBeInstanceOf(Uint8Array);
      expect((ooxml as Uint8Array).byteLength).toBeGreaterThan(100);
      expect((odf as Uint8Array).byteLength).toBeGreaterThan(100);
    } else {
      expect(typeof ooxml).toBe("string");
      expect(typeof odf).toBe("string");
    }
  });

  it("rejects the unsupported reproducible option for ODF roots", async () => {
    const expected = {
      odp: "Reproducible generation is not supported for ODP",
      ods: "Reproducible generation is not supported for ODS",
      odt: "Reproducible generation is not supported for ODT",
    };
    for (const type of ["odt", "ods", "odp"] as const) {
      await expect(
        generateOfficeDocument(type, ODF_FIXTURES[type], "uint8array", {
          reproducible: {},
        } as never),
      ).rejects.toThrow(expected[type]);
    }
  });

  it("reports unsupported, invalid, and encrypted inputs", async () => {
    await expect(
      generateOfficeDocument("encrypted-ooxml" as GenerateType, {} as never),
    ).rejects.toThrow();
    await expect(parseOfficeDocument(new TextEncoder().encode("not office"))).rejects.toThrow(
      "Unable to detect",
    );
    await expect(
      parseOfficeDocument(zipSync({ "not-content": new TextEncoder().encode("<xml/>") })),
    ).rejects.toThrow("missing [Content_Types].xml");
    await expect(
      parseOfficeDocument(zipSync({ mimetype: new TextEncoder().encode("application/x-invalid") })),
    ).rejects.toThrow("Unsupported ODF mimetype");
    await expect(
      parseOfficeDocument(
        buildCfb([
          { name: "EncryptedPackage", data: new Uint8Array(SECTOR_SIZE) },
          { name: "EncryptionInfo", data: new Uint8Array(SECTOR_SIZE) },
        ]),
      ),
    ).rejects.toThrow("Encrypted OOXML");
    await expect(parseOfficeDocument(legacyDocument())).resolves.toHaveProperty("type", "doc");
    await expect(parseOfficeDocument(legacyWorkbook())).resolves.toHaveProperty("type", "xls");
    await expect(parseOfficeDocument(legacyPresentation())).resolves.toHaveProperty("type", "ppt");
  });
});
