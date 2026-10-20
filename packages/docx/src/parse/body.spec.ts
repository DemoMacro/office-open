import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { DocxReadContext } from "../context";
import { parseBody } from "./body";

describe("parseBody", () => {
  it("preserves tolerated block children after a body-level sectPr", () => {
    const body = parseXml(
      '<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        "<w:p><w:r><w:t>before</w:t></w:r></w:p><w:sectPr/>" +
        '<w:altChunk r:id="rId1"/></w:body>',
    ).elements?.[0];
    if (!body) throw new Error("body fixture is missing");
    const context = {
      docx: {
        partRefs: { afChunks: new Map([["rId1", "word/afchunk2.docx"]]) },
        doc: { getRaw: () => new Uint8Array([1, 2, 3]) },
      },
    } as unknown as DocxReadContext;

    expect(parseBody(body, context)).toEqual([
      {
        children: [
          { paragraph: { text: "before" } },
          {
            altChunk: {
              contentType:
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              data: new Uint8Array([1, 2, 3]),
              extension: "docx",
              sourcePath: "afchunk2.docx",
              sourceRid: "rId1",
            },
          },
        ],
        properties: expect.anything(),
      },
    ]);
  });
});
