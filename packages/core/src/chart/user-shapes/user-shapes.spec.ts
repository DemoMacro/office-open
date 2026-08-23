import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parse, stringify } from "../../descriptor";
import type { ReadContext } from "../../descriptor";
import { userShapesDesc } from "./user-shapes";
import type { UserShapesOptions } from "./user-shapes";

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: UserShapesOptions) {
  const xml = stringify(userShapesDesc, opts, {} as never)!;
  const el = parseXml(xml).elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return { xml, result: parse(userShapesDesc, el, readCtx) };
}

describe("userShapesDesc", () => {
  it("round-trips a relative-anchored shape with text body", () => {
    const { xml, result } = roundTrip({
      anchors: [
        {
          from: { x: 0.1, y: 0.2 },
          to: { x: 0.9, y: 0.8 },
          object: {
            type: "shape",
            id: 2,
            nonVisualProperties: { name: "Label" },
            textBox: true,
            textLink: "Sheet1!$A$1",
            locksText: false,
            shapeProperties: { width: 914400, height: 457200, geometry: "rect" },
            textBody: { paragraphs: [{ children: [{ text: "Overlaid label" }] }] },
          },
        },
      ],
    });
    expect(xml).toContain("<cdr:relSizeAnchor>");
    expect(xml).toContain('<cdr:cNvPr id="2" name="Label"/>');
    expect(xml).toContain('txBox="1"');
    expect(xml).toContain('textlink="Sheet1!$A$1"');
    expect(xml).toContain('fLocksText="0"');
    // the text body rides inside a cdr:txBody wrapper, not spread bare
    // into cdr:sp (CT_Shape content model)
    expect(xml).toContain("<cdr:txBody><a:bodyPr");

    const anchor = result.anchors[0]!;
    expect("to" in anchor && anchor.to).toEqual({ x: 0.9, y: 0.8 });
    if ("to" in anchor && anchor.object.type === "shape") {
      expect(anchor.object.id).toBe(2);
      expect(anchor.object.nonVisualProperties?.name).toBe("Label");
      expect(anchor.object.textBox).toBe(true);
      expect(anchor.object.textLink).toBe("Sheet1!$A$1");
      expect(anchor.object.locksText).toBe(false);
      expect(anchor.object.textBody?.paragraphs).toHaveLength(1);
    } else {
      throw new Error("expected a shape object");
    }
  });

  it("round-trips an absolute-anchored connector and picture", () => {
    const { xml, result } = roundTrip({
      anchors: [
        {
          from: { x: 0, y: 0 },
          extent: { width: 1828800, height: 914400 },
          object: {
            type: "connector",
            id: 3,
            shapeProperties: { width: 1828800, height: 914400 },
          },
        },
        {
          from: { x: 0.5, y: 0.5 },
          extent: { width: 914400, height: 914400 },
          object: {
            type: "picture",
            id: 4,
            referenceId: "image1.png",
            shapeProperties: { width: 914400, height: 914400 },
            published: true,
          },
        },
      ],
    });
    expect(xml).toContain("<cdr:absSizeAnchor>");
    expect(xml).toContain("<cdr:cxnSp>");
    expect(xml).toContain('r:embed="{image1.png}"');
    expect(xml).toContain('fPublished="1"');

    const connector = result.anchors[0]!;
    if (!("extent" in connector) || connector.object.type !== "connector") {
      throw new Error("expected a connector anchor");
    }
    expect(connector.extent).toEqual({ width: 1828800, height: 914400 });
    const picture = result.anchors[1]!;
    if (!("extent" in picture) || picture.object.type !== "picture") {
      throw new Error("expected a picture anchor");
    }
    expect(picture.object.referenceId).toBe("image1.png");
    expect(picture.object.published).toBe(true);
  });

  it("round-trips a graphic frame with a raw graphic payload", () => {
    const graphic =
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/></a:graphicData></a:graphic>';
    const { xml, result } = roundTrip({
      anchors: [
        {
          from: { x: 0, y: 0 },
          extent: { width: 914400, height: 914400 },
          object: {
            type: "graphicFrame",
            id: 5,
            graphicFrameLocks: { noChangeAspect: true },
            transform: { x: 0, y: 0, width: 914400, height: 914400 },
            graphic,
          },
        },
      ],
    });
    expect(xml).toContain("<cdr:xfrm>");
    expect(xml).toContain('noChangeAspect="1"');
    expect(xml).toContain('r:id="rId1"');

    const anchor = result.anchors[0]!;
    if (!("extent" in anchor) || anchor.object.type !== "graphicFrame") {
      throw new Error("expected a graphicFrame anchor");
    }
    expect(anchor.object.transform.width).toBe(914400);
    expect(anchor.object.graphic).toContain('r:id="rId1"');
  });

  it("round-trips a nested group", () => {
    const { xml, result } = roundTrip({
      anchors: [
        {
          from: { x: 0, y: 0 },
          extent: { width: 1828800, height: 1828800 },
          object: {
            type: "group",
            id: 6,
            groupShapeProperties: {
              width: 1828800,
              height: 1828800,
              childExtentWidth: 1828800,
              childExtentHeight: 1828800,
            },
            children: [
              {
                type: "shape",
                id: 7,
                shapeProperties: { width: 914400, height: 914400 },
              },
            ],
          },
        },
      ],
    });
    expect(xml).toContain("<cdr:grpSp>");
    expect(xml).toContain("<cdr:grpSpPr>");

    const anchor = result.anchors[0]!;
    if (!("extent" in anchor) || anchor.object.type !== "group") {
      throw new Error("expected a group anchor");
    }
    expect(anchor.object.children).toHaveLength(1);
    expect(anchor.object.children[0]!.type).toBe("shape");
  });
});
