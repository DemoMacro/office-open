import { Formatter } from "@export/formatter";
import type { FileChild } from "@file/file-child";
import { Paragraph, TextRun } from "@file/paragraph";
import { describe, expect, it } from "vite-plus/test";

import { StructuredDocumentTagBlock } from "./sdt";

describe("StructuredDocumentTagBlock", () => {
  it("should create a block SDT with properties only", () => {
    const tree = new Formatter().format(
      new StructuredDocumentTagBlock({
        properties: { alias: "Block Control" },
      }),
    );
    expect(tree).to.deep.equal({
      "w:sdt": [
        {
          "w:sdtPr": [{ "w:alias": { _attr: { "w:val": "Block Control" } } }],
        },
      ],
    });
  });

  it("should create a block SDT with children content", () => {
    const tree = new Formatter().format(
      new StructuredDocumentTagBlock({
        properties: { text: {} },
        children: [
          new Paragraph({
            children: [new TextRun("hello")],
          }),
        ],
      }),
    );
    expect(tree).to.deep.equal({
      "w:sdt": [
        { "w:sdtPr": [{ "w:text": { _attr: { "w:multiLine": false } } }] },
        {
          "w:sdtContent": [
            {
              "w:p": [
                {
                  "w:r": [
                    {
                      "w:t": [
                        {
                          _attr: {
                            "xml:space": "preserve",
                          },
                        },
                        "hello",
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    });
  });

  it("should create a block SDT with empty children array", () => {
    const tree = new Formatter().format(
      new StructuredDocumentTagBlock({
        properties: { text: {} },
        children: [],
      }),
    );
    expect(tree).to.deep.equal({
      "w:sdt": [{ "w:sdtPr": [{ "w:text": { _attr: { "w:multiLine": false } } }] }],
    });
  });

  it("should be assignable to FileChild (block-level element)", () => {
    const sdt: FileChild = new StructuredDocumentTagBlock({
      properties: { richText: true },
    });
    expect(sdt.fileChild).to.be.a("symbol");
  });
});
