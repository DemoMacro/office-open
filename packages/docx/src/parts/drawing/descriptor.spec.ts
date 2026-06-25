import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { BodyContext } from "../../context";
import { drawingDesc, resetDrawingIdGen } from "./descriptor";
import type { DrawingDescriptorOptions } from "./descriptor";
import type { Floating } from "./floating";
import { TextWrappingType } from "./text-wrap";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
  stringifyChild: (child: unknown) => (typeof child === "string" ? child : ""),
  file: {
    media: {
      addImage: () => {},
    },
  },
  fileData: {} as never,
  viewWrapper: {
    relationships: {
      addRelationship: () => {},
    },
  },
} as unknown as BodyContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
  docx: {
    partRefs: {
      media: new Map(),
      charts: new Map(),
      diagramData: new Map(),
    },
    doc: {
      get: () => undefined,
      getRaw: () => undefined,
    },
  },
} as unknown as ReadContext;

function makeImageMediaData() {
  return {
    type: "png" as const,
    fileName: "image1.png",
    data: new Uint8Array([1, 2, 3]),
    transformation: {
      pixels: { x: 0, y: 0 },
      emus: { x: 914400, y: 914400 },
    },
  };
}

// readCtx with media wired so parseImageRun can resolve the blip embed
// ({fileName} placeholder) and read image bytes.
const mediaMap = new Map([["{image1.png}", "word/media/image1.png"]]);
const mediaReadCtx = {
  // parseImageRun resolves the blip embed via resolveRelationship (per-part
  // rels, falling back to partRefs.media) — mirror that here.
  resolveRelationship: (rId: string) => mediaMap.get(rId),
  getPart: () => undefined,
  getRaw: () => undefined,
  docx: {
    partRefs: {
      media: mediaMap,
      charts: new Map(),
      diagramData: new Map(),
    },
    doc: {
      get: () => undefined,
      getRaw: () => new Uint8Array([1, 2, 3]),
    },
  },
} as unknown as ReadContext;

function stringify(opts: DrawingDescriptorOptions) {
  resetDrawingIdGen();
  return drawingDesc.stringify(opts, writeCtx)!;
}

function roundTrip(opts: DrawingDescriptorOptions) {
  const xml = stringify(opts);
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return drawingDesc.parse(el, readCtx);
}

describe("drawingDesc round-trip", () => {
  it("stringifies inline image with w:drawing root", () => {
    const xml = stringify({
      mediaData: makeImageMediaData(),
    });
    expect(xml).toContain("<w:drawing>");
    expect(xml).toContain("<wp:inline");
    expect(xml).toContain("</w:drawing>");
  });

  it("stringifies inline image with correct extent", () => {
    const xml = stringify({
      mediaData: makeImageMediaData(),
    });
    // 914400 EMU = 96 pixels
    expect(xml).toContain('cx="914400"');
    expect(xml).toContain('cy="914400"');
  });

  it("stringifies blip fill with correct embed reference", () => {
    const xml = stringify({
      mediaData: makeImageMediaData(),
    });
    expect(xml).toContain('r:embed="{image1.png}"');
  });

  it("stringifies docPr with custom properties", () => {
    const xml = stringify({
      mediaData: makeImageMediaData(),
      docProperties: {
        name: "MyImage",
        description: "A test image",
        title: "Test Title",
      },
    });
    expect(xml).toContain('name="MyImage"');
    expect(xml).toContain('descr="A test image"');
    expect(xml).toContain('title="Test Title"');
  });

  it("stringifies floating image as anchor", () => {
    const xml = stringify({
      mediaData: makeImageMediaData(),
      floating: {
        horizontalPosition: { align: "center" },
        verticalPosition: { offset: 100000 },
      },
    });
    expect(xml).toContain("<wp:anchor");
    expect(xml).toContain("<wp:positionH");
    expect(xml).toContain("<wp:positionV");
  });

  it("stringifies chart media data", () => {
    const xml = stringify({
      mediaData: {
        type: "chart" as const,
        chartKey: "chart1",
        transformation: {
          pixels: { x: 0, y: 0 },
          emus: { x: 5000000, y: 3000000 },
        },
      },
    });
    expect(xml).toContain("c:chart");
    expect(xml).toContain("chart1");
  });

  it("stringifies smartart media data", () => {
    const xml = stringify({
      mediaData: {
        type: "smartart" as const,
        smartArtKey: "smartart1",
        transformation: {
          pixels: { x: 0, y: 0 },
          emus: { x: 5000000, y: 3000000 },
        },
      },
    });
    expect(xml).toContain("dgm:relIds");
    expect(xml).toContain("smartart1");
  });

  it("parse returns an object for inline image", () => {
    const result = roundTrip({
      mediaData: makeImageMediaData(),
    });
    // parse returns {} when mediaPath can't be resolved
    expect(result).toBeDefined();
  });

  it("stringifies graphic with xmlns:a namespace", () => {
    const xml = stringify({
      mediaData: makeImageMediaData(),
    });
    expect(xml).toContain('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"');
  });

  it("stringifies pic:spPr with preset geometry rect", () => {
    const xml = stringify({
      mediaData: makeImageMediaData(),
    });
    expect(xml).toContain('prst="rect"');
    expect(xml).toContain("<pic:spPr");
  });

  it("stringifies wps shape with preset geometry (not hardcoded rect)", () => {
    const xml = stringify({
      mediaData: {
        type: "wps" as const,
        transformation: { pixels: { x: 0, y: 0 }, emus: { x: 914400, y: 914400 } },
        data: {
          children: [],
          presetGeometry: { preset: "roundRect" },
        },
      },
    });
    expect(xml).toContain('prst="roundRect"');
    expect(xml).not.toContain('prst="rect"');
  });

  it("round-trips wps shape preset geometry", () => {
    const xml = stringify({
      mediaData: {
        type: "wps" as const,
        transformation: { pixels: { x: 0, y: 0 }, emus: { x: 914400, y: 914400 } },
        data: {
          children: [],
          presetGeometry: { preset: "roundRect" },
        },
      },
    });
    const doc = parseXml(xml);
    const el = doc.elements![0];
    const result = drawingDesc.parse(el, mediaReadCtx) as {
      wpsShape?: { presetGeometry?: { preset?: string } };
    };
    expect(result.wpsShape?.presetGeometry?.preset).toBe("roundRect");
  });

  it("round-trips floating image margins/flags/relativeFrom/wrap", () => {
    // parseImageRun must read all Floating fields the anchor stringify writes:
    // margins (distT-D), relativeFrom, allowOverlap/behindDoc/locked/
    // layoutInCell/relativeHeight, and wrap (type number + side).
    const xml = stringify({
      mediaData: makeImageMediaData(),
      floating: {
        horizontalPosition: { relative: "column", align: "center" },
        verticalPosition: { relative: "page", offset: 100000 },
        margins: { top: 50000, bottom: 60000, left: 70000, right: 80000 },
        allowOverlap: false,
        behindDocument: true,
        layoutInCell: false,
        lockAnchor: true,
        zIndex: 200000,
        wrap: { type: TextWrappingType.SQUARE, side: "left" },
      },
    });
    const doc = parseXml(xml);
    const el = doc.elements![0];
    const result = drawingDesc.parse(el, mediaReadCtx) as { image?: { floating?: Floating } };
    const floating = result.image?.floating;
    expect(floating).toBeDefined();
    expect(floating!.margins).toEqual({ top: 50000, bottom: 60000, left: 70000, right: 80000 });
    expect(floating!.horizontalPosition.relative).toBe("column");
    expect(floating!.horizontalPosition.align).toBe("center");
    expect(floating!.verticalPosition.relative).toBe("page");
    expect(floating!.verticalPosition.offset).toBe(100000);
    expect(floating!.allowOverlap).toBe(false);
    expect(floating!.behindDocument).toBe(true);
    expect(floating!.layoutInCell).toBe(false);
    expect(floating!.lockAnchor).toBe(true);
    expect(floating!.zIndex).toBe(200000);
    expect(floating!.wrap?.type).toBe(TextWrappingType.SQUARE);
    expect(floating!.wrap?.side).toBe("left");
  });
});
