import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { MediaTransformation } from "@shared/media/data";
import { describe, expect, it } from "vite-plus/test";

import type { BodyContext } from "../../context";
import { DocxParseError } from "../../errors";
import { drawingDesc, resetDrawingIdGen } from "./descriptor";
import type { DrawingDescriptorOptions } from "./descriptor";
import type { Floating } from "./floating";
import type { WrapPolygon } from "./text-wrap";
import { TextWrappingType } from "./text-wrap";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
  addHyperlink: () => {},
  stringifyChild: undefined as unknown as BodyContext["stringifyChild"],
  fileData: {} as never,
  viewWrapper: {
    relationships: {
      addRelationship: () => 1,
      add: () => 1,
      hasId: () => false,
      idOf: () => undefined,
    },
  },
} as unknown as BodyContext;

const mediaMap = new Map([["{image1.png}", "word/media/image1.png"]]);
const mediaReadCtx = {
  currentPart: "word/document.xml",
  resolveRelationship: (relationshipId: string) => mediaMap.get(relationshipId),
  getPart: () => undefined,
  getRaw: () => undefined,
  docx: {
    partRefs: { media: mediaMap, charts: new Map(), diagramData: new Map() },
    doc: { get: () => undefined, getRaw: () => new Uint8Array([1, 2, 3]) },
  },
} as unknown as ReadContext;

type ParsedPicture = {
  picture?: {
    transformation?: {
      inlineDistances?: unknown;
      effectExtent?: unknown;
    };
    floating?: Floating;
  };
};

function makeImageMediaData(): {
  type: "png";
  fileName: string;
  data: Uint8Array;
  transformation: MediaTransformation;
} {
  return {
    type: "png" as const,
    fileName: "image1.png",
    data: new Uint8Array([1, 2, 3]),
    transformation: {
      width: 914400,
      height: 914400,
    },
  };
}

function stringify(opts: DrawingDescriptorOptions) {
  resetDrawingIdGen();
  return drawingDesc.stringify(opts, writeCtx)!;
}

function parseDrawing(el: Element) {
  return drawingDesc.parse(el, mediaReadCtx) as ParsedPicture;
}

function parseXmlDrawing(xml: string) {
  const root = parseXml(xml).elements?.[0];
  if (!root) throw new Error("parsed drawing has no root element");
  return root;
}

function anchorFrom(xml: string) {
  const root = parseXmlDrawing(xml);
  const anchor = root.elements?.find((element) => element.name === "wp:anchor");
  if (!anchor) throw new Error("drawing fixture has no wp:anchor");
  return { root, anchor };
}

describe("WordprocessingDrawing wrapper model", () => {
  it("round-trips all wp:inline distances", () => {
    const mediaData = makeImageMediaData();
    mediaData.transformation.inlineDistances = {
      top: 1000,
      bottom: 2000,
      left: 3000,
      right: 4000,
    };
    const xml = stringify({ mediaData });
    expect(xml).toContain('distT="1000"');
    expect(xml).toContain('distB="2000"');
    expect(xml).toContain('distL="3000"');
    expect(xml).toContain('distR="4000"');

    const result = parseDrawing(parseXmlDrawing(xml));
    expect(result.picture?.transformation?.inlineDistances).toEqual({
      top: 1000,
      bottom: 2000,
      left: 3000,
      right: 4000,
    });
  });

  it("keeps explicit false values for hidden and simplePos", () => {
    const { root, anchor } = anchorFrom(
      stringify({
        mediaData: makeImageMediaData(),
        floating: {
          horizontalPosition: { relative: "page", offset: 0 },
          verticalPosition: { relative: "page", offset: 0 },
          simplePos: { x: 0, y: 0 },
        },
      }),
    );
    anchor.attributes = { ...anchor.attributes, hidden: "0", simplePos: "0" };
    const result = parseDrawing(root);
    expect(result.picture?.floating?.hidden).toBe(false);
    expect(result.picture?.floating?.simplePos).toEqual({ enabled: false, x: 0, y: 0 });

    const roundTrip = stringify({
      mediaData: makeImageMediaData(),
      floating: result.picture!.floating!,
    });
    expect(roundTrip).toContain('simplePos="0"');
    expect(roundTrip).toContain('hidden="0"');
  });

  it("parses all effectExtent coordinates", () => {
    const mediaData = makeImageMediaData();
    mediaData.transformation.effectExtent = { left: -1, top: 2, right: 3, bottom: 4 };
    const result = parseDrawing(parseXmlDrawing(stringify({ mediaData })));
    expect(result.picture?.transformation?.effectExtent).toEqual({
      left: -1,
      top: 2,
      right: 3,
      bottom: 4,
    });
  });

  it("round-trips wrapSquare margins and its nested effectExtent", () => {
    const xml = stringify({
      mediaData: makeImageMediaData(),
      floating: {
        horizontalPosition: { relative: "page", align: "center" },
        verticalPosition: { relative: "page", offset: 0 },
        wrap: {
          type: TextWrappingType.SQUARE,
          side: "largest",
          margins: { distT: 11, distB: 22, distL: 33, distR: 44 },
          effectExtent: { left: 55, top: 66, right: 77, bottom: 88 },
        },
      },
    });
    expect(xml).toContain('distT="11" distB="22" distL="33" distR="44"');
    expect(xml).toContain('<wp:effectExtent l="55" t="66" r="77" b="88"/>');

    const result = parseDrawing(parseXmlDrawing(xml));
    expect(result.picture?.floating?.wrap).toMatchObject({
      type: TextWrappingType.SQUARE,
      side: "largest",
      margins: { distT: 11, distB: 22, distL: 33, distR: 44 },
      effectExtent: { left: 55, top: 66, right: 77, bottom: 88 },
    });
  });

  it("round-trips wrapTopAndBottom margins and its nested effectExtent", () => {
    const xml = stringify({
      mediaData: makeImageMediaData(),
      floating: {
        horizontalPosition: { relative: "page", align: "center" },
        verticalPosition: { relative: "page", offset: 0 },
        wrap: {
          type: TextWrappingType.TOP_AND_BOTTOM,
          margins: { distT: 12, distB: 34 },
          effectExtent: { left: 56, top: 78, right: 90, bottom: 12 },
        },
      },
    });
    expect(xml).toContain('<wp:wrapTopAndBottom distT="12" distB="34">');
    const result = parseDrawing(parseXmlDrawing(xml));
    expect(result.picture?.floating?.wrap).toMatchObject({
      type: TextWrappingType.TOP_AND_BOTTOM,
      margins: { distT: 12, distB: 34 },
      effectExtent: { left: 56, top: 78, right: 90, bottom: 12 },
    });
  });

  it("round-trips wrapTight distances and polygon", () => {
    const polygon: WrapPolygon = {
      edited: false,
      points: [
        { x: 1, y: 2 },
        { x: 3, y: 4 },
        { x: 5, y: 6 },
      ],
    };
    const xml = stringify({
      mediaData: makeImageMediaData(),
      floating: {
        horizontalPosition: { relative: "page", align: "center" },
        verticalPosition: { relative: "page", offset: 0 },
        wrap: {
          type: TextWrappingType.TIGHT,
          side: "right",
          margins: { distL: 10, distR: 20 },
          polygon,
        },
      },
    });
    expect(xml).toContain('<wp:wrapTight wrapText="right" distL="10" distR="20">');
    expect(xml).toContain('<wp:wrapPolygon edited="0">');
    expect(xml).toContain('<wp:start x="1" y="2"/>');
    expect(xml.match(/<wp:lineTo /g)).toHaveLength(2);

    const result = parseDrawing(parseXmlDrawing(xml));
    expect(result.picture?.floating?.wrap).toMatchObject({
      type: TextWrappingType.TIGHT,
      side: "right",
      margins: { distL: 10, distR: 20 },
      polygon,
    });
  });

  it("round-trips wrapThrough distances and polygon", () => {
    const xml = stringify({
      mediaData: makeImageMediaData(),
      floating: {
        horizontalPosition: { relative: "page", align: "center" },
        verticalPosition: { relative: "page", offset: 0 },
        wrap: {
          type: TextWrappingType.THROUGH,
          side: "left",
          margins: { distL: 30, distR: 40 },
          polygon: {
            edited: true,
            points: [
              { x: 7, y: 8 },
              { x: 9, y: 10 },
              { x: 11, y: 12 },
            ],
          },
        },
      },
    });
    const result = parseDrawing(parseXmlDrawing(xml));
    expect(result.picture?.floating?.wrap).toMatchObject({
      type: TextWrappingType.THROUGH,
      side: "left",
      margins: { distL: 30, distR: 40 },
      polygon: {
        edited: true,
        points: [
          { x: 7, y: 8 },
          { x: 9, y: 10 },
          { x: 11, y: 12 },
        ],
      },
    });
  });

  it("parses explicit zero position offsets and percentages", () => {
    const xml = stringify({
      mediaData: makeImageMediaData(),
      floating: {
        horizontalPosition: { relative: "page", percentOffset: 0 },
        verticalPosition: { relative: "page", offset: 0 },
      },
    });
    const result = parseDrawing(parseXmlDrawing(xml));
    expect(result.picture?.floating?.horizontalPosition.percentOffset).toBe(0);
    expect(result.picture?.floating?.verticalPosition.offset).toBe(0);
  });

  it("rejects wrap polygons that violate CT_WrapPath", () => {
    const source = stringify({
      mediaData: makeImageMediaData(),
      floating: {
        horizontalPosition: { relative: "page", align: "center" },
        verticalPosition: { relative: "page", offset: 0 },
        wrap: {
          type: TextWrappingType.TIGHT,
          polygon: {
            points: [
              { x: 1, y: 2 },
              { x: 3, y: 4 },
              { x: 5, y: 6 },
            ],
          },
        },
      },
    });
    const invalidXmlCases = [
      source.replace(/<wp:wrapPolygon>.*<\/wp:wrapPolygon>/, ""),
      source.replace('<wp:lineTo x="5" y="6"/>', ""),
      source.replace('<wp:start x="1" y="2"/>', '<wp:start x="NaN" y="2"/>'),
    ];

    for (const xml of invalidXmlCases) {
      try {
        parseDrawing(parseXmlDrawing(xml));
        throw new Error("expected invalid wrap polygon to throw");
      } catch (error) {
        expect(error).toBeInstanceOf(DocxParseError);
        const parseError = error as DocxParseError;
        expect(parseError.reason).toBe("invalid-drawing-model");
        expect(parseError.part).toBe("word/document.xml");
        expect(parseError.path).toContain("wp:wrapPolygon");
      }
    }
  });
});
