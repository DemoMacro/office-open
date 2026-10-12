import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { DocxParseError } from "../../errors";
import { fontTableDesc } from "./descriptor";
import type { FontTableInput } from "./descriptor";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as import("@office-open/core/descriptor").WriteContext;

const readCtx = {} as unknown as ReadContext;

function roundTrip(opts: FontTableInput) {
  const xml = fontTableDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return fontTableDesc.parse(el, readCtx);
}

describe("fontTableDesc round-trip", () => {
  it("round-trips a single font", () => {
    const result = roundTrip({
      fonts: [
        {
          name: "Arial",
          embedRegular: { fontKey: "abc-123", data: Buffer.from([]), embedRid: "rId1" },
        },
      ],
    });
    expect(result.fonts).toHaveLength(1);
    expect(result.fonts[0]?.name).toBe("Arial");
    // The key is canonicalized to uppercase on emit (ST_Guid canonical form)
    expect(result.fonts[0]?.embedRegular?.fontKey).toBe("ABC-123");
  });

  it("round-trips multiple fonts", () => {
    const result = roundTrip({
      fonts: [{ name: "Arial" }, { name: "Times New Roman" }, { name: "Calibri" }],
    });
    expect(result.fonts).toHaveLength(3);
    expect(result.fonts[0]?.name).toBe("Arial");
    expect(result.fonts[1]?.name).toBe("Times New Roman");
    expect(result.fonts[2]?.name).toBe("Calibri");
  });

  it("round-trips font with characterSet", () => {
    const result = roundTrip({
      fonts: [{ name: "Wingdings", characterSet: "02" }],
    });
    expect(result.fonts[0]?.characterSet).toBe("02");
  });

  it("round-trips notTrueType — bare element true, w:val=0 false", () => {
    const xml = fontTableDesc.stringify(
      {
        fonts: [
          { name: "Arial", notTrueType: true },
          { name: "Courier New", notTrueType: false },
        ],
      },
      writeCtx,
    )!;
    expect(xml).toContain("<w:notTrueType/>");
    expect(xml).toContain('<w:notTrueType w:val="0"/>');
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("no root");
    const parsed = fontTableDesc.parse(el, readCtx);
    expect(parsed.fonts[0]?.notTrueType).toBe(true);
    expect(parsed.fonts[1]?.notTrueType).toBe(false);
  });

  it("round-trips font with characterSetName (w:characterSet)", () => {
    const result = roundTrip({
      fonts: [
        {
          name: "Wingdings",
          characterSet: "02",
          characterSetName: "ISO-8859-1",
        },
      ],
    });
    expect(result.fonts[0]?.characterSet).toBe("02");
    expect(result.fonts[0]?.characterSetName).toBe("ISO-8859-1");
  });

  it("round-trips font key stripping braces", () => {
    const result = roundTrip({
      fonts: [
        {
          name: "TestFont",
          embedRegular: { fontKey: "some-key-value", data: Buffer.from([]), embedRid: "rId1" },
        },
      ],
    });
    expect(result.fonts[0]?.embedRegular?.fontKey).toBe("SOME-KEY-VALUE");
  });

  it("round-trips all embedded faces", () => {
    const result = roundTrip({
      fonts: [
        {
          name: "Family",
          embedRegular: { fontKey: "regular", data: Buffer.from([]), embedRid: "rId1" },
          embedBold: { fontKey: "bold", data: Buffer.from([]), embedRid: "rId2" },
          embedItalic: { fontKey: "italic", data: Buffer.from([]), embedRid: "rId3" },
          embedBoldItalic: { fontKey: "bolditalic", data: Buffer.from([]), embedRid: "rId4" },
        },
      ],
    });
    expect(result.fonts[0]?.embedRegular?.fontKey).toBe("REGULAR");
    expect(result.fonts[0]?.embedBold?.fontKey).toBe("BOLD");
    expect(result.fonts[0]?.embedItalic?.fontKey).toBe("ITALIC");
    expect(result.fonts[0]?.embedBoldItalic?.fontKey).toBe("BOLDITALIC");
  });

  it("round-trips empty fonts", () => {
    const result = roundTrip({ fonts: [] });
    expect(result.fonts).toHaveLength(0);
  });

  it("rejects a font declaration without required @w:name", () => {
    const xml =
      '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:font><w:pitch w:val="variable"/></w:font></w:fonts>';
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("no root");

    expect(() => fontTableDesc.parse(el, readCtx)).toThrow(DocxParseError);
  });

  it("round-trips a markup-compatibility gated font declaration", () => {
    const xml =
      '<w:fonts xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas">' +
      '<w:font w:name="Calibri"><w:pitch w:val="variable"/></w:font>' +
      '<mc:AlternateContent><mc:Choice Requires="wpc">' +
      '<w:font w:name="Calibri1"><w:pitch w:val="variable"/></w:font>' +
      "</mc:Choice><mc:Fallback></mc:Fallback></mc:AlternateContent>" +
      "</w:fonts>";
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("no root");
    const parsed = fontTableDesc.parse(el, readCtx);
    expect(parsed.fonts).toHaveLength(2);
    expect(parsed.fonts[0]?.requires).toBeUndefined();
    expect(parsed.fonts[1]?.name).toBe("Calibri1");
    expect(parsed.fonts[1]?.requires).toBe("wpc");

    const generated = fontTableDesc.stringify(parsed, writeCtx)!;
    expect(generated).toContain(
      'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"',
    );
    expect(generated).toContain('<mc:Choice Requires="wpc"><w:font w:name="Calibri1">');
    expect(generated).toContain("<mc:Fallback/></mc:AlternateContent>");
  });
});
