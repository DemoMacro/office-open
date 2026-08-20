import { Relationships } from "@office-open/core";
import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { stringifyBodyChild } from "../../body";
import type { BodyContext } from "../../context";
import { parseSectionChild } from "../../parse/body";
import { setBodyParseChild } from "../bodychildren";
import { drawingDesc, resetDrawingIdGen } from "./descriptor";
import type { DrawingDescriptorOptions } from "./descriptor";
import type { Floating } from "./floating";
import { TextWrappingType } from "./text-wrap";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
  stringifyChild: undefined as unknown as BodyContext["stringifyChild"],
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
writeCtx.stringifyChild = (child) => stringifyBodyChild(child, writeCtx);
setBodyParseChild(parseSectionChild);

const readContextData = {
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
};
const readCtx = readContextData as unknown as ReadContext;

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

// readCtx with media wired so parsePictureRun can resolve the blip embed
// ({fileName} placeholder) and read image bytes.
const mediaMap = new Map([["{image1.png}", "word/media/image1.png"]]);
const mediaReadCtx = {
  // parsePictureRun resolves the blip embed via resolveRelationship (per-part
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
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
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

  it("stringifies image effects through the shared blip serializer", () => {
    const xml = stringify({
      mediaData: makeImageMediaData(),
      blipEffects: {
        luminance: { bright: 20 },
        duotone: {
          color1: { value: "002060" },
          color2: { value: "D0CECE" },
        },
      },
    });
    expect(xml).toContain('<a:lum bright="20000"/>');
    expect(xml).toContain(
      '<a:duotone><a:srgbClr val="002060"/><a:srgbClr val="D0CECE"/></a:duotone>',
    );
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

  it("round-trips wps linked text box and normalEastAsianFlow", () => {
    const xml = stringify({
      mediaData: {
        type: "wps" as const,
        transformation: { pixels: { x: 0, y: 0 }, emus: { x: 914400, y: 914400 } },
        data: {
          children: [],
          linkedTextBox: { id: 7, sequence: 2 },
          normalEastAsianFlow: true,
        },
      },
    });
    expect(xml).toContain('<wps:wsp normalEastAsianFlow="1">');
    expect(xml).toContain('<wps:linkedTxbx id="7" seq="2"/>');

    const doc = parseXml(xml);
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = drawingDesc.parse(el, readCtx) as {
      wpsShape?: {
        linkedTextBox?: { id: number; sequence: number };
        normalEastAsianFlow?: boolean;
      };
    };
    expect(result.wpsShape?.linkedTextBox).toEqual({ id: 7, sequence: 2 });
    expect(result.wpsShape?.normalEastAsianFlow).toBe(true);
  });

  it("round-trips a relationship-based wps text box part", () => {
    const relationships = new Relationships();
    const relationCtx = {
      ...writeCtx,
      viewWrapper: { relationships },
    } as BodyContext;
    relationCtx.stringifyChild = (child) => stringifyBodyChild(child, relationCtx);
    const xml = drawingDesc.stringify(
      {
        mediaData: {
          type: "wps" as const,
          transformation: { pixels: { x: 0, y: 0 }, emus: { x: 914400, y: 914400 } },
          data: {
            children: [],
            textBoxPart: { path: "word/txbx1.xml", sequence: 0 },
          },
        },
      },
      relationCtx,
    )!;

    expect(xml).toContain('<wps:txbx r:txbx="rId1" txbxSeq="0"/>');
    expect(relationships.serialize()).toContain(
      'Type="http://schemas.microsoft.com/office/2006/relationships/txbx" Target="txbx1.xml"',
    );

    const doc = parseXml(xml);
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const relationReadCtx = {
      ...readCtx,
      currentPart: "word/document.xml",
      docx: {
        ...readContextData.docx,
        partRefs: {
          ...readContextData.docx.partRefs,
          partTextBoxes: new Map([["word/document.xml", new Map([["rId1", "word/txbx1.xml"]])]]),
        },
      },
    } as unknown as ReadContext;
    const parsed = drawingDesc.parse(el, relationReadCtx) as {
      wpsShape?: { children: unknown[]; textBoxPart?: { path: string; sequence: number } };
    };
    expect(parsed.wpsShape?.children).toEqual([]);
    expect(parsed.wpsShape?.textBoxPart).toEqual({ path: "word/txbx1.xml", sequence: 0 });
  });

  it("round-trips block SDT content in a wps text box", () => {
    const xml = stringify({
      mediaData: {
        type: "wps" as const,
        transformation: { pixels: { x: 0, y: 0 }, emus: { x: 914400, y: 914400 } },
        data: {
          children: [
            {
              sdt: {
                properties: {
                  alias: "Title",
                  id: -958338334,
                  showingPlaceholder: true,
                  dataBinding: {
                    prefixMappings:
                      "xmlns:ns0='http://schemas.openxmlformats.org/package/2006/metadata/core-properties'",
                    xpath: "/ns0:coreProperties[1]/dc:title[1]",
                    storeItemID: "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}",
                  },
                  text: {},
                },
                endProperties: {},
                children: [{ paragraph: { children: [{ text: "     " }] } }],
              },
            },
          ],
        },
      },
    });
    expect(xml).toContain("<wps:txbx><w:txbxContent><w:sdt>");
    expect(xml).toContain('<w:alias w:val="Title"/>');
    expect(xml).toContain('<w:id w:val="-958338334"/>');
    expect(xml).toContain("<w:showingPlcHdr/>");
    expect(xml).toContain('w:storeItemID="{6C3C8BC8-F283-45AE-878A-BAB7291924A1}"');
    expect(xml).toContain('w:prefixMappings="xmlns:ns0=&apos;');
    expect(xml).toContain('<w:text w:multiLine="false"/>');
    expect(xml).toContain("<w:sdtEndPr/>");
    expect(xml).toContain('<w:t xml:space="preserve">     </w:t>');

    const doc = parseXml(xml);
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const parsed = drawingDesc.parse(el, readCtx) as {
      wpsShape?: {
        children: Array<{
          sdt?: {
            properties: { alias?: string; id?: number; showingPlaceholder?: boolean };
            children?: unknown[];
          };
        }>;
      };
    };
    const sdt = parsed.wpsShape?.children[0]?.sdt;
    expect(sdt?.properties).toMatchObject({
      alias: "Title",
      id: -958338334,
      showingPlaceholder: true,
    });
    expect(sdt?.children).toHaveLength(1);
  });

  it("parses every block-level child in a wps text box", () => {
    const xml = stringify({
      mediaData: {
        type: "wps" as const,
        transformation: { pixels: { x: 0, y: 0 }, emus: { x: 914400, y: 914400 } },
        data: {
          children: [{ rawXml: "<w:unknown/>" }, { paragraph: "text" }],
        },
      },
    });
    const doc = parseXml(xml);
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const parsed = drawingDesc.parse(el, readCtx) as {
      wpsShape?: { children: Array<Record<string, unknown>> };
    };
    expect(parsed.wpsShape?.children[0]).toEqual({ rawXml: "<w:unknown/>" });
    expect(parsed.wpsShape?.children[1]).toEqual({ text: "text" });
  });

  it("stringifies a chart group child as wpg:graphicFrame", () => {
    const xml = stringify({
      mediaData: {
        type: "wpg" as const,
        transformation: { pixels: { x: 0, y: 0 }, emus: { x: 1828800, y: 914400 } },
        children: [
          {
            type: "chart" as const,
            chartKey: "chart_1",
            transformation: { pixels: { x: 0, y: 0 }, emus: { x: 914400, y: 914400 } },
            nonVisualProperties: { id: 6, name: "Chart 6" },
          },
        ],
      },
    });
    expect(xml).toContain("<wpg:graphicFrame>");
    expect(xml).toContain('<wpg:cNvPr id="6" name="Chart 6"/>');
    expect(xml).toContain("<wpg:cNvFrPr/>");
    expect(xml).toContain('r:id="{chart:chart_1}"');
  });

  it("round-trips a content part (wpg:contentPart) with cpLocks", () => {
    const xml = stringify({
      mediaData: {
        type: "wpg" as const,
        transformation: { pixels: { x: 0, y: 0 }, emus: { x: 1828800, y: 914400 } },
        children: [
          {
            type: "contentPart" as const,
            referenceId: "rId9",
            transformation: { pixels: { x: 0, y: 0 }, emus: { x: 914400, y: 914400 } },
            nonVisualProperties: {
              id: 3,
              name: "Video",
              contentPart: { isComment: false, locks: { noChangeAspect: true } },
            },
            blackWhiteMode: "auto" as const,
          },
        ],
      },
    });
    expect(xml).toContain('<wpg:contentPart r:id="rId9" bwMode="auto">');
    expect(xml).toContain('<wpg:cNvContentPartPr isComment="0">');
    expect(xml).toContain('<a:cpLocks noChangeAspect="true"/>');

    const doc = parseXml(xml);
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = drawingDesc.parse(el, readCtx) as {
      contentPart?: {
        referenceId?: string;
        blackWhiteMode?: string;
        nonVisualProperties?: {
          contentPart?: { isComment?: boolean; locks?: Record<string, boolean> };
        };
      };
    };
    expect(result.contentPart?.referenceId).toBe("rId9");
    expect(result.contentPart?.blackWhiteMode).toBe("auto");
    expect(result.contentPart?.nonVisualProperties?.contentPart?.isComment).toBe(false);
    expect(result.contentPart?.nonVisualProperties?.contentPart?.locks?.noChangeAspect).toBe(true);
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
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = drawingDesc.parse(el, mediaReadCtx) as {
      wpsShape?: { presetGeometry?: { preset?: string } };
    };
    expect(result.wpsShape?.presetGeometry?.preset).toBe("roundRect");
  });

  it("round-trips wp14 percentage positioning", () => {
    const xml = stringify({
      mediaData: makeImageMediaData(),
      floating: {
        horizontalPosition: { relative: "page", percentOffset: 45.5 },
        verticalPosition: { relative: "page", percentOffset: 66 },
      },
    });
    expect(xml).toContain("<wp14:pctPosHOffset>45500</wp14:pctPosHOffset>");
    expect(xml).toContain("<wp14:pctPosVOffset>66000</wp14:pctPosVOffset>");

    const doc = parseXml(xml);
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = drawingDesc.parse(el, mediaReadCtx) as { picture?: { floating?: Floating } };
    expect(result.picture?.floating?.horizontalPosition).toEqual({
      relative: "page",
      percentOffset: 45.5,
    });
    expect(result.picture?.floating?.verticalPosition).toEqual({
      relative: "page",
      percentOffset: 66,
    });
  });

  it("parses percentage positioning from mc:AlternateContent choices", () => {
    const xml = stringify({
      mediaData: makeImageMediaData(),
      floating: {
        horizontalPosition: { relative: "page", offset: 4576445 },
        verticalPosition: { relative: "page", offset: 5129530 },
      },
    })
      .replace(
        '<wp:positionH relativeFrom="page"><wp:posOffset>4576445</wp:posOffset></wp:positionH>',
        '<mc:AlternateContent><mc:Choice Requires="wp14"><wp:positionH relativeFrom="page"><wp14:pctPosHOffset>45500</wp14:pctPosHOffset></wp:positionH></mc:Choice><mc:Fallback><wp:positionH relativeFrom="page"><wp:posOffset>4576445</wp:posOffset></wp:positionH></mc:Fallback></mc:AlternateContent>',
      )
      .replace(
        '<wp:positionV relativeFrom="page"><wp:posOffset>5129530</wp:posOffset></wp:positionV>',
        '<mc:AlternateContent><mc:Choice Requires="wp14"><wp:positionV relativeFrom="page"><wp14:pctPosVOffset>66000</wp14:pctPosVOffset></wp:positionV></mc:Choice><mc:Fallback><wp:positionV relativeFrom="page"><wp:posOffset>5129530</wp:posOffset></wp:positionV></mc:Fallback></mc:AlternateContent>',
      );

    const doc = parseXml(xml);
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = drawingDesc.parse(el, mediaReadCtx) as { picture?: { floating?: Floating } };
    expect(result.picture?.floating?.horizontalPosition).toEqual({
      relative: "page",
      percentOffset: 45.5,
      offset: 4576445,
    });
    expect(result.picture?.floating?.verticalPosition).toEqual({
      relative: "page",
      percentOffset: 66,
      offset: 5129530,
    });

    const roundTripXml = stringify({
      mediaData: makeImageMediaData(),
      floating: result.picture!.floating,
    });
    expect(roundTripXml.match(/<mc:AlternateContent>/g)).toHaveLength(2);
    expect(roundTripXml).toContain("<wp14:pctPosHOffset>45500</wp14:pctPosHOffset>");
    expect(roundTripXml).toContain("<wp:posOffset>4576445</wp:posOffset>");
  });

  it("round-trips floating image margins/flags/relativeFrom/wrap", () => {
    // parsePictureRun must read all Floating fields the anchor stringify writes:
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
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = drawingDesc.parse(el, mediaReadCtx) as { picture?: { floating?: Floating } };
    const floating = result.picture?.floating;
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

  it("round-trips image rotation via pic:spPr/a:xfrm/@rot", () => {
    // parsePictureRun must read pic:spPr/a:xfrm/@rot (ST_Angle, 1/60000 deg) and
    // convert to degrees — otherwise rotated images lose orientation on round-trip.
    const base = makeImageMediaData();
    const xml = stringify({
      mediaData: { ...base, transformation: { ...base.transformation, rotation: 270 } },
    });
    expect(xml).toContain('rot="16200000"');
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const result = drawingDesc.parse(el, mediaReadCtx) as {
      picture?: { transformation?: { rotation?: number } };
    };
    // 16200000 / 60000 = 270 degrees
    expect(result.picture?.transformation?.rotation).toBe(270);
  });
});
