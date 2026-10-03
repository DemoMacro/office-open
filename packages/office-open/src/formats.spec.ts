import { zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import { detectOfficeFormat } from "./formats";

const encoder = new TextEncoder();
const END_OF_CHAIN = 0xfffffffen;
const FAT_SECT = 0xfffffffdn;
const FREE_SECT = 0xffffffffn;
const NO_STREAM = 0xffffffffn;

function buildCfb(streamName: string): Uint8Array {
  const bytes = new Uint8Array(512 + 3 * 512);
  const view = new DataView(bytes.buffer);
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
  view.setUint32(60, Number(END_OF_CHAIN), true);
  view.setUint32(64, 0, true);
  view.setUint32(68, Number(END_OF_CHAIN), true);
  view.setUint32(72, 0, true);
  view.setUint32(76, 0, true);
  for (let index = 1; index < 109; index++) {
    view.setUint32(76 + index * 4, Number(FREE_SECT), true);
  }
  view.setUint32(512, Number(FAT_SECT), true);
  view.setUint32(516, Number(END_OF_CHAIN), true);
  view.setUint32(520, Number(END_OF_CHAIN), true);

  const writeEntry = (
    offset: number,
    name: string,
    objectType: number,
    childId: number,
    startSector: number,
    size: number,
  ): void => {
    for (const [index, char] of Array.from(name).entries()) {
      view.setUint16(offset + index * 2, char.charCodeAt(0), true);
    }
    view.setUint16(offset + name.length * 2, 0, true);
    view.setUint16(offset + 64, (name.length + 1) * 2, true);
    bytes[offset + 66] = objectType;
    bytes[offset + 67] = 1;
    view.setUint32(offset + 68, Number(NO_STREAM), true);
    view.setUint32(offset + 72, Number(NO_STREAM), true);
    view.setUint32(offset + 76, childId, true);
    view.setUint32(offset + 116, startSector, true);
    view.setUint32(offset + 120, size, true);
  };
  writeEntry(512 + 512, "Root Entry", 5, 1, Number(END_OF_CHAIN), 0);
  writeEntry(512 + 512 + 128, streamName, 2, Number(NO_STREAM), 2, 4096);
  return bytes;
}

function odfZip(mimeType: string): Uint8Array {
  return zipSync({ mimetype: encoder.encode(mimeType) });
}

function ooxmlZip(mainContentType: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": encoder.encode(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/x.xml" ContentType="${mainContentType}"/></Types>`,
    ),
  });
}

describe("detectOfficeFormat", () => {
  it("detects RTF text and bytes", () => {
    const expected = { format: "rtf", family: "document", container: "text" } as const;
    expect(detectOfficeFormat("{\\rtf1\\ansi}")).toEqual(expected);
    expect(detectOfficeFormat(encoder.encode("{\\rtf1\\ansi}"))).toEqual(expected);
  });

  it("detects all ODF siblings", () => {
    expect(detectOfficeFormat(odfZip("application/vnd.oasis.opendocument.text"))).toEqual({
      format: "odt",
      family: "document",
      container: "zip",
    });
    expect(detectOfficeFormat(odfZip("application/vnd.oasis.opendocument.spreadsheet"))).toEqual({
      format: "ods",
      family: "workbook",
      container: "zip",
    });
    expect(detectOfficeFormat(odfZip("application/vnd.oasis.opendocument.presentation"))).toEqual({
      format: "odp",
      family: "presentation",
      container: "zip",
    });
  });

  it("detects OOXML package variants", () => {
    expect(
      detectOfficeFormat(
        ooxmlZip(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        ),
      ),
    ).toEqual({
      format: "docx",
      family: "document",
      container: "zip",
    });
    expect(
      detectOfficeFormat(
        ooxmlZip("application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"),
      ),
    ).toEqual({
      format: "pptm",
      family: "presentation",
      container: "zip",
    });
    expect(
      detectOfficeFormat(ooxmlZip("application/vnd.ms-excel.sheet.macroEnabled.main+xml")),
    ).toEqual({
      format: "xlsm",
      family: "workbook",
      container: "zip",
    });
  });

  it("detects legacy CFB formats by their semantic streams", () => {
    expect(detectOfficeFormat(buildCfb("WordDocument"))).toEqual({
      format: "doc",
      family: "document",
      container: "cfb",
    });
    expect(detectOfficeFormat(buildCfb("Workbook"))).toEqual({
      format: "xls",
      family: "workbook",
      container: "cfb",
    });
    expect(detectOfficeFormat(buildCfb("PowerPoint Document"))).toEqual({
      format: "ppt",
      family: "presentation",
      container: "cfb",
    });
  });

  it("rejects unknown containers", () => {
    expect(() => detectOfficeFormat(new Uint8Array([1, 2, 3]))).toThrow("Unable to detect");
    expect(() => detectOfficeFormat(zipSync({ "hello.txt": new Uint8Array() }))).toThrow(
      "missing [Content_Types].xml",
    );
  });
});
