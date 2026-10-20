import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { DocxReadContext } from "../../context";
import { parseAltChunk } from "./alt-chunk-parse";

describe("parseAltChunk", () => {
  it("preserves an embedded WordprocessingML package", () => {
    const data = new Uint8Array([80, 75, 3, 4]);
    const context = {
      docx: {
        partRefs: { afChunks: new Map([["rId1", "word/afchunk2.docx"]]) },
        doc: { getRaw: (path: string) => (path === "word/afchunk2.docx" ? data : undefined) },
      },
    } as unknown as DocxReadContext;
    const element = parseXml(
      '<w:altChunk xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>',
    ).elements?.[0];
    if (!element) throw new Error("altChunk fixture is missing");

    expect(parseAltChunk(element, context)).toMatchObject({
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      data,
      extension: "docx",
      sourcePath: "afchunk2.docx",
      sourceRid: "rId1",
    });
  });
});
