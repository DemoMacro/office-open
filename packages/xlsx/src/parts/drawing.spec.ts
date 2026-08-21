import { unzipSync } from "@office-open/core";
import type { HyperlinkTarget, ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook } from "../generate";
import { drawingDesc } from "./drawing";
import type { DrawingOptions } from "./drawing";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
  addHyperlink: () => {},
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: DrawingOptions) {
  const xml = drawingDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return drawingDesc.parse(el, readCtx);
}

describe("drawingDesc round-trip", () => {
  it("returns undefined for empty images and charts", () => {
    const xml = drawingDesc.stringify({ images: [], charts: [] }, writeCtx);
    expect(xml).toBeUndefined();
  });

  it("round-trips single image", () => {
    const opts: DrawingOptions = {
      images: [{ col: 2, row: 3, rId: "rId1" }],
    };
    const result = roundTrip(opts);
    const images = result.images!;

    expect(images).toHaveLength(1);
    expect(images[0]?.col).toBe(2);
    expect(images[0]?.row).toBe(3);
    expect(images[0]?.rId).toBe("rId1");
  });

  it("round-trips image with offsets", () => {
    const opts: DrawingOptions = {
      images: [{ col: 1, row: 1, colOffset: 50000, rowOffset: 25000, rId: "rId1" }],
    };
    const result = roundTrip(opts);
    const images = result.images!;

    expect(images[0]?.colOffset).toBe(50000);
    expect(images[0]?.rowOffset).toBe(25000);
  });

  it("round-trips image locksWithSheet and printsWithSheet", () => {
    const opts: DrawingOptions = {
      images: [{ col: 1, row: 1, rId: "rId1", locksWithSheet: false, printsWithSheet: false }],
    };
    const result = roundTrip(opts);
    const images = result.images!;

    expect(images[0]?.locksWithSheet).toBe(false);
    expect(images[0]?.printsWithSheet).toBe(false);
  });

  it("round-trips multiple images", () => {
    const opts: DrawingOptions = {
      images: [
        { col: 1, row: 1, rId: "rId1" },
        { col: 5, row: 10, rId: "rId2" },
      ],
    };
    const result = roundTrip(opts);
    const images = result.images!;

    expect(images).toHaveLength(2);
    expect(images[1]?.col).toBe(5);
    expect(images[1]?.row).toBe(10);
  });

  it("round-trips smartArt relIds and anchor", () => {
    const opts: DrawingOptions = {
      smartArts: [
        {
          dataRId: "rId1",
          layoutRId: "rId2",
          quickStyleRId: "rId3",
          colorsRId: "rId4",
          col: 2,
          row: 3,
          toCol: 8,
          toRow: 12,
          name: "Diagram 1",
        },
      ],
    };
    const xml = drawingDesc.stringify(opts, writeCtx)!;
    expect(xml).toContain('uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"');
    expect(xml).toMatch(/r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"/);

    const result = roundTrip(opts);
    const smartArt = result.smartArts![0]!;
    expect(smartArt.dataRId).toBe("rId1");
    expect(smartArt.layoutRId).toBe("rId2");
    expect(smartArt.quickStyleRId).toBe("rId3");
    expect(smartArt.colorsRId).toBe("rId4");
    expect(smartArt.col).toBe(2);
    expect(smartArt.row).toBe(3);
    expect(smartArt.toCol).toBe(8);
    expect(smartArt.toRow).toBe(12);
    expect(smartArt.name).toBe("Diagram 1");
  });
  it("round-trips chart", () => {
    const opts: DrawingOptions = {
      charts: [{ col: 1, row: 1, rId: "rId3" }],
    };
    const result = roundTrip(opts);
    const charts = result.charts!;

    expect(charts).toHaveLength(1);
    expect(charts[0]?.rId).toBe("rId3");
  });

  it("round-trips chart with offsets", () => {
    const opts: DrawingOptions = {
      charts: [{ col: 3, row: 5, colOffset: 10000, rowOffset: 20000, rId: "rId4" }],
    };
    const result = roundTrip(opts);
    const charts = result.charts!;

    expect(charts[0]?.col).toBe(3);
    expect(charts[0]?.row).toBe(5);
    expect(charts[0]?.colOffset).toBe(10000);
    expect(charts[0]?.rowOffset).toBe(20000);
  });

  it("round-trips chart to corner, editAs and cNvPr", () => {
    const opts: DrawingOptions = {
      charts: [
        {
          col: 2,
          row: 3,
          toCol: 15,
          toRow: 18,
          toColOffset: 5000,
          editAs: "twoCell",
          name: "SalesChart",
          description: "Q3 sales",
          rId: "rId7",
        },
      ],
    };
    const result = roundTrip(opts);
    const chart = result.charts![0]!;

    expect(chart.toCol).toBe(15);
    expect(chart.toRow).toBe(18);
    expect(chart.toColOffset).toBe(5000);
    expect(chart.editAs).toBe("twoCell");
    expect(chart.name).toBe("SalesChart");
    expect(chart.description).toBe("Q3 sales");
  });

  it("round-trips mixed images and charts", () => {
    const opts: DrawingOptions = {
      images: [{ col: 1, row: 1, rId: "rId1" }],
      charts: [{ col: 5, row: 5, rId: "rId2" }],
    };
    const result = roundTrip(opts);

    expect(result.images).toHaveLength(1);
    expect(result.charts).toHaveLength(1);
  });
});

describe("drawingDesc — anchored shapes", () => {
  it("round-trips a shape with geometry and text body", () => {
    const opts: DrawingOptions = {
      shapes: [
        {
          col: 2,
          row: 3,
          name: "TextBox 1",
          spPr: { x: 100, y: 200, width: 1000, height: 500, geometry: "rect" },
          textBody: { paragraphs: [{ text: "Hello" }] },
        },
      ],
    };
    const result = roundTrip(opts);
    const shape = result.shapes![0]!;

    expect(result.shapes).toHaveLength(1);
    expect(shape.col).toBe(2);
    expect(shape.row).toBe(3);
    expect(shape.name).toBe("TextBox 1");
    expect(shape.spPr).toMatchObject({ x: 100, y: 200, width: 1000, height: 500 });
    expect(shape.spPr.geometry).toMatchObject({ preset: "rect" });
    expect(shape.textBody?.paragraphs).toHaveLength(1);
  });

  it("registers drawing shape text-hyperlink runs via ctx.addHyperlink", () => {
    const registered: Array<{ key: string; url: string; tooltip?: string }> = [];
    const ctx = {
      addRelationship: () => "rId1",
      addMedia: () => "",
      addHyperlink: (key: string, target: HyperlinkTarget) => {
        registered.push({ key, url: target.url ?? "", tooltip: target.tooltip });
      },
    } as unknown as WriteContext;

    const opts: DrawingOptions = {
      shapes: [
        {
          col: 0,
          row: 0,
          spPr: { geometry: "rect" },
          textBody: {
            paragraphs: [
              {
                children: [
                  { text: "Link", hyperlink: { url: "https://example.com", tooltip: "Example" } },
                ],
              },
            ],
          },
        },
      ],
    };
    const xml = drawingDesc.stringify(opts, ctx)!;

    // The DrawingML text hyperlink is registered through ctx.addHyperlink with a
    // placeholder key; the compiler resolves {hlink:key} → a real rId later.
    expect(registered).toHaveLength(1);
    expect(registered[0]?.url).toBe("https://example.com");
    expect(registered[0]?.tooltip).toBe("Example");
    expect(xml).toContain("a:hlinkClick");
    expect(xml).toContain(`r:id="{hlink:${registered[0]?.key}}"`);
  });

  it("round-trips shape macro and textlink attributes", () => {
    const opts: DrawingOptions = {
      shapes: [{ col: 1, row: 1, spPr: { geometry: "rect" }, macro: "Click()", textlink: "rId1" }],
    };
    const result = roundTrip(opts);
    expect(result.shapes![0]?.macro).toBe("Click()");
    expect(result.shapes![0]?.textlink).toBe("rId1");
  });

  it("round-trips a oneCellAnchor shape with extent", () => {
    const opts: DrawingOptions = {
      shapes: [
        {
          col: 1,
          row: 1,
          anchorType: "oneCell",
          extentCx: 2000,
          extentCy: 1000,
          spPr: { geometry: "ellipse" },
        },
      ],
    };
    const result = roundTrip(opts);
    const shape = result.shapes![0]!;
    expect(shape.anchorType).toBe("oneCell");
    expect(shape.extentCx).toBe(2000);
    expect(shape.extentCy).toBe(1000);
  });
});

describe("drawingDesc — anchored connectors", () => {
  it("round-trips a connector with line geometry", () => {
    const opts: DrawingOptions = {
      connectors: [
        {
          col: 1,
          row: 1,
          toCol: 5,
          toRow: 5,
          name: "Arrow 1",
          spPr: { geometry: "line" },
        },
      ],
    };
    const result = roundTrip(opts);
    const conn = result.connectors![0]!;
    expect(result.connectors).toHaveLength(1);
    expect(conn.name).toBe("Arrow 1");
    expect(conn.toCol).toBe(5);
    expect(conn.spPr.geometry).toMatchObject({ preset: "line" });
  });

  it("round-trips connector locking and endpoint connections", () => {
    const opts: DrawingOptions = {
      connectors: [
        {
          col: 1,
          row: 1,
          toCol: 5,
          toRow: 5,
          spPr: { geometry: "line" },
          locking: { noAdjustHandles: true, noChangeShapeType: true },
          startConnection: { id: 1, index: 0 },
          endConnection: { id: 2, index: 3 },
        },
      ],
    };
    const result = roundTrip(opts);
    const conn = result.connectors![0]!;
    expect(conn.locking).toEqual({ noAdjustHandles: true, noChangeShapeType: true });
    expect(conn.startConnection).toEqual({ id: 1, index: 0 });
    expect(conn.endConnection).toEqual({ id: 2, index: 3 });
  });
});

describe("drawingDesc — anchored groups", () => {
  it("round-trips a group with nested shapes and connectors", () => {
    const opts: DrawingOptions = {
      groups: [
        {
          col: 1,
          row: 1,
          toCol: 10,
          toRow: 10,
          name: "Group 1",
          grpSpPr: {
            x: 0,
            y: 0,
            width: 5000,
            height: 5000,
            childOffsetX: 0,
            childOffsetY: 0,
            childExtentWidth: 5000,
            childExtentHeight: 5000,
          },
          shapes: [{ name: "Child 1", spPr: { geometry: "rect" } }],
          connectors: [{ name: "Child Line", spPr: { geometry: "line" } }],
        },
      ],
    };
    const result = roundTrip(opts);
    const group = result.groups![0]!;
    expect(result.groups).toHaveLength(1);
    expect(group.name).toBe("Group 1");
    expect(group.grpSpPr.childExtentWidth).toBe(5000);
    expect(group.shapes).toHaveLength(1);
    expect(group.shapes![0]?.name).toBe("Child 1");
    expect(group.connectors).toHaveLength(1);
    expect(group.connectors![0]?.name).toBe("Child Line");
  });
});

describe("drawingDesc — anchored content parts", () => {
  it("round-trips a content part reference", () => {
    const opts: DrawingOptions = {
      contentParts: [{ col: 1, row: 1, toCol: 3, toRow: 3, rId: "rId9" }],
    };
    const result = roundTrip(opts);
    const cp = result.contentParts![0]!;
    expect(result.contentParts).toHaveLength(1);
    expect(cp.rId).toBe("rId9");
    expect(cp.toCol).toBe(3);
  });
});

describe("drawing shape hyperlink — compiler resolution", () => {
  it("resolves {hlink:key} placeholders to real rIds and emits External hyperlink rels", async () => {
    const buffer = (await generateWorkbook(
      {
        worksheets: [
          {
            name: "Sheet1",
            shapes: [
              {
                col: 0,
                row: 0,
                spPr: { geometry: "rect" },
                textBody: {
                  paragraphs: [
                    {
                      children: [
                        {
                          text: "Link",
                          hyperlink: { url: "https://example.com", tooltip: "Example" },
                        },
                      ],
                    },
                  ],
                },
              },
            ],
          },
        ],
      },
      { type: "uint8array" },
    )) as Uint8Array;

    const unzipped = unzipSync(buffer);
    const drawingXml = new TextDecoder().decode(unzipped["xl/drawings/drawing1.xml"]);
    const drawingRels = new TextDecoder().decode(unzipped["xl/drawings/_rels/drawing1.xml.rels"]);

    // Placeholder replaced with a real rId on the drawing shape's a:hlinkClick.
    expect(drawingXml).toContain("a:hlinkClick");
    expect(drawingXml).not.toContain("{hlink:");
    expect(drawingXml).toMatch(/r:id="rId\d+"/);
    // External hyperlink relationship emitted on the drawing rels.
    expect(drawingRels).toContain("/relationships/hyperlink");
    expect(drawingRels).toContain('TargetMode="External"');
    expect(drawingRels).toContain("https://example.com");
  });
});

describe("drawing picture cNvPr — compiler passthrough", () => {
  it("threads worksheet picture name/description/title/hidden through to the drawing cNvPr", async () => {
    const buffer = (await generateWorkbook(
      {
        worksheets: [
          {
            name: "Sheet1",
            images: [
              {
                type: "png",
                data: "AAAA",
                col: 1,
                row: 1,
                name: "Logo",
                description: "Company logo",
                title: "Logo title",
                hidden: true,
              },
            ],
          },
        ],
      },
      { type: "uint8array" },
    )) as Uint8Array;

    const unzipped = unzipSync(buffer);
    const drawingXml = new TextDecoder().decode(unzipped["xl/drawings/drawing1.xml"]);

    expect(drawingXml).toContain('name="Logo"');
    expect(drawingXml).toContain('descr="Company logo"');
    expect(drawingXml).toContain('title="Logo title"');
    expect(drawingXml).toContain('hidden="1"');
  });
});

describe("linked-only picture (a:blip @r:link)", () => {
  it("round-trips a linked-only blip as rId-less linkRId", () => {
    const opts: DrawingOptions = {
      images: [{ col: 1, row: 1, rId: "", linkRId: "rId2" }],
    };
    const xml = drawingDesc.stringify(opts, writeCtx)!;
    expect(xml).toContain('<a:blip r:link="rId2"/>');
    expect(xml).not.toContain("r:embed");

    const result = roundTrip(opts);
    const image = result.images![0]!;
    expect(image.rId).toBe("");
    expect(image.linkRId).toBe("rId2");
  });

  it("emits an External image relationship and no media part from sourceUrl", async () => {
    const buffer = (await generateWorkbook(
      {
        worksheets: [
          {
            name: "Sheet1",
            images: [
              {
                type: "png",
                sourceUrl: "https://example.com/logo.png",
                col: 1,
                row: 1,
              },
            ],
          },
        ],
      },
      { type: "uint8array" },
    )) as Uint8Array;

    const unzipped = unzipSync(buffer);
    const drawingXml = new TextDecoder().decode(unzipped["xl/drawings/drawing1.xml"]);
    const drawingRels = new TextDecoder().decode(unzipped["xl/drawings/_rels/drawing1.xml.rels"]);

    // Linked-only: the blip carries r:link alone — no r:embed, no media bytes.
    expect(drawingXml).toMatch(/<a:blip r:link="rId\d+"\/>/);
    expect(drawingXml).not.toContain("r:embed");
    expect(drawingRels).toContain("/relationships/image");
    expect(drawingRels).toContain('TargetMode="External"');
    expect(drawingRels).toContain("https://example.com/logo.png");
    expect(Object.keys(unzipped).some((p) => p.startsWith("xl/media/"))).toBe(false);
  });
});
