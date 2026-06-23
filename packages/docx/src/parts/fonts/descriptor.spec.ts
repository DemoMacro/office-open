import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

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
  const el = doc.elements![0];
  return fontTableDesc.parse(el, readCtx);
}

describe("fontTableDesc round-trip", () => {
  it("round-trips a single font", () => {
    const result = roundTrip({
      fonts: [{ name: "Arial", fontKey: "abc-123", data: Buffer.from([]), embedRid: "rId1" }],
    });
    expect(result.fonts).toHaveLength(1);
    expect(result.fonts[0].name).toBe("Arial");
    expect(result.fonts[0].fontKey).toBe("abc-123");
  });

  it("round-trips multiple fonts", () => {
    const result = roundTrip({
      fonts: [
        { name: "Arial", fontKey: "key-1", data: Buffer.from([]) },
        { name: "Times New Roman", fontKey: "key-2", data: Buffer.from([]) },
        { name: "Calibri", fontKey: "key-3", data: Buffer.from([]) },
      ],
    });
    expect(result.fonts).toHaveLength(3);
    expect(result.fonts[0].name).toBe("Arial");
    expect(result.fonts[1].name).toBe("Times New Roman");
    expect(result.fonts[2].name).toBe("Calibri");
  });

  it("round-trips font with characterSet", () => {
    const result = roundTrip({
      fonts: [{ name: "Wingdings", fontKey: "wd-key", data: Buffer.from([]), characterSet: "02" }],
    });
    expect(result.fonts[0].characterSet).toBe("02");
  });

  it("round-trips font with characterSetName (w:characterSet)", () => {
    const result = roundTrip({
      fonts: [
        {
          name: "Wingdings",
          fontKey: "wd-key",
          data: Buffer.from([]),
          characterSet: "02",
          characterSetName: "ISO-8859-1",
        },
      ],
    });
    expect(result.fonts[0].characterSet).toBe("02");
    expect(result.fonts[0].characterSetName).toBe("ISO-8859-1");
  });

  it("round-trips font key stripping braces", () => {
    const result = roundTrip({
      fonts: [
        { name: "TestFont", fontKey: "some-key-value", data: Buffer.from([]), embedRid: "rId1" },
      ],
    });
    expect(result.fonts[0].fontKey).toBe("some-key-value");
  });

  it("round-trips empty fonts", () => {
    const result = roundTrip({ fonts: [] });
    expect(result.fonts).toHaveLength(0);
  });
});
