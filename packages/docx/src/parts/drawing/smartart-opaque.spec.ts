import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import { generateDocument, generateDocumentSync } from "../../generate";
import { parseDocumentSync } from "../../parse";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const DRAWING =
  XML_DECL +
  '<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:spTree/></dsp:drawing>';
const MEDIA = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]);

function documentWithOpaqueSmartArt() {
  return {
    sections: [
      {
        children: [
          {
            paragraph: {
              children: [
                {
                  smartArt: {
                    altText: { id: "1", name: "Diagram" },
                    nodes: [{ text: "Item" }],
                    transformation: { width: 100, height: 100 },
                    raw: {
                      drawing: DRAWING,
                      media: [{ fileName: "image1.png", data: MEDIA }],
                    },
                  },
                },
              ],
            },
          },
        ],
      },
    ],
  };
}

describe("SmartArt opaque part round-trip", () => {
  it("rebuilds canonical diagram parts and preserves opaque companions", async () => {
    const output = await generateDocument(documentWithOpaqueSmartArt(), { type: "uint8array" });
    const files = unzipSync(output);
    const decoder = new TextDecoder();

    expect(decoder.decode(files["word/diagrams/data1.xml"]!)).toContain("<a:t>Item</a:t>");
    expect(decoder.decode(files["word/diagrams/data1.xml"]!)).toContain(
      "<dgm:dataModel xmlns:dgm=",
    );
    expect(decoder.decode(files["word/diagrams/drawing1.xml"]!)).toBe(DRAWING);
    const rels = decoder.decode(files["word/diagrams/_rels/data1.xml.rels"]!);
    expect(rels).toContain(
      'Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"',
    );
    expect(rels).toContain('Target="drawing1.xml"');
    expect(rels).toContain('Target="../media/image1.png"');
    expect(Array.from(files["word/media/image1.png"]!)).toEqual(Array.from(MEDIA));
  });

  it("keeps opaque companions stable on a second round-trip", async () => {
    const first = await generateDocument(documentWithOpaqueSmartArt(), { type: "uint8array" });
    const second = await generateDocument(parseDocumentSync(first), { type: "uint8array" });
    const firstFiles = unzipSync(first);
    const secondFiles = unzipSync(second);

    expect(new TextDecoder().decode(secondFiles["word/diagrams/drawing1.xml"]!)).toBe(DRAWING);
    expect(Array.from(secondFiles["word/media/image1.png"]!)).toEqual(
      Array.from(firstFiles["word/media/image1.png"]!),
    );
  });

  it("preserves source SmartArt data rels and embedded rIds", async () => {
    const generated = await generateDocument(documentWithOpaqueSmartArt(), { type: "uint8array" });
    const files = unzipSync(generated);
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    files["word/diagrams/data1.xml"] = encoder.encode(
      decoder
        .decode(files["word/diagrams/data1.xml"]!)
        .replace(
          /<dgm:spPr\s*\/>/,
          '<dgm:spPr><a:blipFill><a:blip r:embed="rId9"/>' +
            "<a:stretch><a:fillRect/></a:stretch></a:blipFill></dgm:spPr>",
        ),
    );
    files["word/diagrams/_rels/data1.xml.rels"] = encoder.encode(
      XML_DECL +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId4" Type="http://example.com/relationships/diagramDrawing" ' +
        'Target="/word/diagrams/drawing1.xml"/>' +
        '<Relationship Id="rId9" Type="http://example.com/relationships/image" ' +
        'Target="/word/media/image1.png"/>' +
        "</Relationships>",
    );

    const output = unzipSync(
      generateDocumentSync(parseDocumentSync(zipSync(files)), { type: "uint8array" }),
    );
    expect(decoder.decode(output["word/diagrams/data1.xml"]!)).toContain('r:embed="rId9"');
    const rels = decoder.decode(output["word/diagrams/_rels/data1.xml.rels"]!);
    expect(rels).toContain(
      'Id="rId4" Type="http://example.com/relationships/diagramDrawing" ' +
        'Target="/word/diagrams/drawing1.xml"',
    );
    expect(rels).toContain(
      'Id="rId9" Type="http://example.com/relationships/image" Target="/word/media/image1.png"',
    );
  });
});
