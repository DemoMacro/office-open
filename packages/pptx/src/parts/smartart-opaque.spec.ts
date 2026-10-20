import { toUint8Array } from "@office-open/core";
import { unzipSync, zipSync } from "@office-open/core/opc";
import { describe, expect, it } from "vitest";

import { generatePresentation } from "../generate";
import { parsePresentation } from "../parse";
import { getLayoutXml } from "./smartart";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const DRAWING =
  XML_DECL +
  '<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:spTree/></dsp:drawing>';
const IMAGE = new Uint8Array([1, 2, 3, 4]);

describe("SmartArt opaque part round-trip", () => {
  it("preserves every source data-part relationship", async () => {
    const generated = await generatePresentation({
      slides: [
        {
          children: [
            {
              smartart: {
                x: 0,
                y: 0,
                width: "100px",
                height: "100px",
                nodes: [{ text: "Source" }],
              },
            },
          ],
        },
      ],
    });
    const archive = unzipSync(generated);
    archive["ppt/diagrams/_rels/data1.xml.rels"] = new TextEncoder().encode(
      XML_DECL +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ' +
        'Target="https://example.com/" TargetMode="External"/></Relationships>',
    );
    const parsed = await parsePresentation(zipSync(archive));
    const child = parsed.slides?.[0]?.children?.[0];
    if (!child || !("smartart" in child)) throw new Error("expected a SmartArt child");
    expect(child.smartart.sourceDataRelationships).toEqual([
      {
        id: "rId1",
        relationshipType:
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        target: "https://example.com/",
        targetMode: "External",
      },
    ]);

    const regenerated = await generatePresentation(parsed);
    const rels = new TextDecoder().decode(
      unzipSync(regenerated)["ppt/diagrams/_rels/data1.xml.rels"]!,
    );
    expect(rels).toContain('Id="rId1"');
  });

  it("preserves the drawing cache while rebuilding canonical diagram parts", async () => {
    const generated = await generatePresentation({
      slides: [
        {
          children: [
            {
              smartart: {
                x: 0,
                y: 0,
                width: "100px",
                height: "100px",
                nodes: [{ text: "Source" }],
                raw: { drawing: DRAWING },
              },
            },
          ],
        },
      ],
    });
    const options = await parsePresentation(generated);
    const child = options.slides?.[0]?.children?.[0];
    if (!child || !("smartart" in child)) throw new Error("expected a SmartArt child");
    const drawing = child.smartart.raw?.drawing;
    if (!drawing) throw new Error("expected the opaque drawing cache");
    expect(new TextDecoder().decode(toUint8Array(drawing))).toBe(DRAWING);
    expect(child.smartart.nodes).toEqual([{ text: "Source" }]);
  });

  it("registers companion media and rebuilds drawing relationships", async () => {
    const generated = await generatePresentation({
      slides: [
        {
          children: [
            {
              smartart: {
                x: 0,
                y: 0,
                width: "100px",
                height: "100px",
                nodes: [{ text: "Spacer" }],
              },
            },
            {
              smartart: {
                x: 0,
                y: 0,
                width: "100px",
                height: "100px",
                nodes: [{ text: "Source" }],
                raw: { drawing: DRAWING, media: [{ fileName: "source.png", data: IMAGE }] },
              },
            },
          ],
        },
      ],
    });
    const files = unzipSync(generated);
    const rels = new TextDecoder().decode(files["ppt/diagrams/_rels/data2.xml.rels"]);
    expect(rels).toContain(
      'Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"',
    );
    expect(rels).toContain('Target="drawing2.xml"');
    expect(rels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"',
    );
    expect(rels).toContain('Target="../media/source.png"');
    expect(files["ppt/media/source.png"]).toEqual(IMAGE);
    expect(new TextDecoder().decode(files["ppt/diagrams/data2.xml"])).toContain(
      "<a:t>Source</a:t>",
    );
    expect(new TextDecoder().decode(files["ppt/diagrams/data2.xml"])).toContain(
      "<dgm:dataModel xmlns:dgm=",
    );
    expect(new TextDecoder().decode(files["ppt/diagrams/layout2.xml"])).toBe(
      getLayoutXml("default"),
    );
  });
});
