import { parse } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { DocxReadContext } from "../../context";
import { parseSdtBlock, parseSdtProperties } from "./sdt-parse";

const readCtx = {} as unknown as DocxReadContext;

describe("parseSdtProperties", () => {
  it("reads date language from w:lid/w:val", () => {
    const doc = parse(
      '<w:sdtPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:date w:fullDate="2001-01-01T00:00:00Z"><w:lid w:val="en-US"/></w:date></w:sdtPr>',
    );
    const root = doc.elements?.[0];
    if (!root) throw new Error("parsed document has no root element");
    expect(parseSdtProperties(root, readCtx).date?.languageId).toBe("en-US");
  });

  it("rejects an unsupported multiline on/off value", () => {
    const root = parse(
      '<w:sdtPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:text w:multiLine="maybe"/></w:sdtPr>',
    ).elements?.[0];
    if (!root) throw new Error("parsed document has no root element");
    expect(() => parseSdtProperties(root, readCtx)).toThrow("Invalid SDT multiline on/off value");
  });
});

describe("parseSdtBlock", () => {
  it("distinguishes an absent properties shell from an empty one", () => {
    const root = parse(
      '<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        "<w:sdtContent><w:p/></w:sdtContent></w:sdt>",
    ).elements?.[0];
    if (!root) throw new Error("parsed document has no root element");
    expect(parseSdtBlock(root, {} as never, () => []).properties).toEqual({
      propertiesPresent: false,
    });
  });
});
