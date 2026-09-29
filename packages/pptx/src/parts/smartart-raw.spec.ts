import { toUint8Array } from "@office-open/core";
import { createDataModel } from "@office-open/core/smartart";
import { describe, expect, it } from "vitest";

import { generatePresentation } from "../generate";
import { parsePresentation } from "../parse";
import { getColorXml, getLayoutXml, getStyleXml } from "./smartart";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const RELS_DECL =
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
const DRAWING =
  '<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:spTree/></dsp:drawing>';
const DATA_RELS =
  XML_DECL +
  RELS_DECL +
  '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="drawing1.xml"/>' +
  "</Relationships>";

describe("SmartArt raw round-trip", () => {
  it("preserves source diagram parts instead of rebuilding them", async () => {
    const raw = {
      data: XML_DECL + createDataModel([{ text: "Source" }], "default", "simple1", "accent1_2"),
      layout: XML_DECL + getLayoutXml("default"),
      style: XML_DECL + getStyleXml("simple1"),
      color: XML_DECL + getColorXml("accent1_2"),
      drawing: XML_DECL + DRAWING,
      dataRels: DATA_RELS,
    };
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
                raw,
              },
            },
          ],
        },
      ],
    });
    const options = await parsePresentation(generated);
    const child = options.slides?.[0]?.children?.[0];
    if (!child || !("smartart" in child)) throw new Error("expected a SmartArt child");
    const saved = child.smartart.raw;
    if (!saved?.data || !saved.drawing || !saved.dataRels) throw new Error("expected raw parts");
    const decoder = new TextDecoder();
    expect(decoder.decode(toUint8Array(saved.data))).toBe(raw.data);
    expect(decoder.decode(toUint8Array(saved.drawing))).toBe(raw.drawing);
    expect(decoder.decode(toUint8Array(saved.dataRels))).toBe(DATA_RELS);
    expect(child.smartart.nodes).toEqual([{ text: "Source" }]);
  });
});
