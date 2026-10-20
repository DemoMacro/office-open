import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { DocxReadContext } from "../../context";
import { parseDrawingRun } from "./drawing-parse";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const CHART_SPACE =
  "<c:chartSpace><c:chart><c:plotArea><c:barChart>" +
  '<c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit>' +
  '<c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>' +
  "</c:barChart></c:plotArea></c:chart></c:chartSpace>";

function chartContext(): DocxReadContext {
  const chartEl = parseXml(`<root ${NS}>${CHART_SPACE}</root>`).elements?.[0]?.elements?.[0];
  return {
    docx: {
      partRefs: {
        charts: new Map([["rId1", "word/charts/chart1.xml"]]),
        hyperlinks: new Map([
          ["rId2", "https://example.invalid/chart"],
          ["rId3", "https://example.invalid/preview"],
        ]),
      },
      doc: {
        get: (path: string) => (path === "word/charts/chart1.xml" ? chartEl : undefined),
      },
    },
  } as unknown as DocxReadContext;
}

function drawingXml(docPrAttrs?: string, docPrChildren = ""): string {
  const docPr =
    docPrAttrs === undefined ? "" : `<wp:docPr ${docPrAttrs}>${docPrChildren}</wp:docPr>`;
  return (
    `<w:drawing ${NS}><wp:inline><wp:extent cx="5486400" cy="3200400"/>` +
    '<wp:effectExtent l="19050" t="0" r="19050" b="0"/>' +
    docPr +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
    `<c:chart r:id="rId1"/></a:graphicData></a:graphic></wp:inline></w:drawing>`
  );
}

function parseDrawing(docPrAttrs?: string, docPrChildren = "") {
  const el = parseXml(drawingXml(docPrAttrs, docPrChildren)).elements?.[0];
  if (!el) throw new Error("parsed drawing has no root element");
  return parseDrawingRun(el, chartContext());
}

describe("parseChartDrawing alt text", () => {
  it("carries wp:docPr alt text onto the chart options", () => {
    const result = parseDrawing('id="1" name="Chart 1" descr="Sales chart" title="Sales"');
    expect(result).toMatchObject({
      chart: { altText: { name: "Chart 1", description: "Sales chart", title: "Sales" } },
    });
  });

  it("carries docPr hyperlink tooltips", () => {
    const result = parseDrawing(
      'id="1" name="Chart 1"',
      '<a:hlinkClick r:id="rId2" tgtFrame="_blank" tooltip="Open chart"/>' +
        '<a:hlinkHover r:id="rId3" tgtFrame="_parent" tooltip="Preview chart"/>',
    );
    expect(result).toMatchObject({
      chart: {
        altText: {
          hyperlink: {
            clickTooltip: "Open chart",
            clickTargetFrame: "_blank",
            hoverTooltip: "Preview chart",
            hoverTargetFrame: "_parent",
          },
        },
      },
    });
  });

  it("resolves drawing hover hyperlinks from the current part's rels", () => {
    const chartEl = parseXml(`<root ${NS}>${CHART_SPACE}</root>`).elements?.[0]?.elements?.[0];
    if (!chartEl) throw new Error("chart fixture missing");
    const context = {
      currentPart: "word/document.xml",
      docx: {
        partRefs: {
          charts: new Map([["rId1", "word/charts/chart1.xml"]]),
          hyperlinks: new Map(),
          partHyperlinks: new Map([
            [
              "word/document.xml",
              new Map([
                ["rId2", "https://example.invalid/click"],
                ["rId3", "https://example.invalid/hover"],
              ]),
            ],
          ]),
        },
        doc: { get: () => chartEl },
      },
    } as unknown as DocxReadContext;
    const el = parseXml(
      drawingXml('id="1" name="Chart 1"', '<a:hlinkClick r:id="rId2"/><a:hlinkHover r:id="rId3"/>'),
    ).elements?.[0];
    if (!el) throw new Error("parsed drawing has no root element");
    const result = parseDrawingRun(el, context);
    expect(result).toMatchObject({
      chart: {
        altText: {
          hyperlink: {
            click: "https://example.invalid/click",
            clickRelationshipId: "rId2",
            hover: "https://example.invalid/hover",
            hoverRelationshipId: "rId3",
          },
        },
      },
    });
  });

  it("leaves altText undefined when the drawing carries no wp:docPr", () => {
    const result = parseDrawing();
    const chart = (result as { chart?: { altText?: unknown } } | undefined)?.chart;
    expect(chart?.altText).toBeUndefined();
  });

  it("carries wp:effectExtent onto the chart transformation", () => {
    const result = parseDrawing();
    expect(result).toMatchObject({
      chart: {
        transformation: {
          effectExtent: { left: 19050, top: 0, right: 19050, bottom: 0 },
        },
      },
    });
  });
});

describe("parsePictureRun blip extensions", () => {
  it("preserves the source-specific useLocalDpi URI", () => {
    const xml =
      `<w:drawing ${NS} xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ` +
      'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">' +
      '<wp:inline><wp:extent cx="9525" cy="9525"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="rId1"><a:extLst>' +
      '<a:ext uri="28A0092B-C50C-407e-A947-70E740481C1C">' +
      '<a14:useLocalDpi val="0"/></a:ext></a:extLst></a:blip></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>";
    const context = {
      docx: {
        doc: { getRaw: () => new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]) },
      },
      resolveRelationship: () => "word/media/image.png",
    } as unknown as DocxReadContext;
    const result = parseDrawingRun(parseXml(xml).elements![0]!, context);
    expect(result).toMatchObject({
      picture: {
        useLocalDpi: false,
        useLocalDpiUri: "28A0092B-C50C-407e-A947-70E740481C1C",
      },
    });
  });
});

describe("group and canvas DrawingML extensions", () => {
  const mediaContext = {
    docx: { doc: { getRaw: () => new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]) } },
    resolveRelationship: () => "word/media/image.png",
  } as unknown as DocxReadContext;

  it("preserves useLocalDpi on canvas picture children", () => {
    const xml =
      `<w:drawing ${NS} xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ` +
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
      'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">' +
      '<wp:inline><wp:extent cx="9525" cy="9525"/><a:graphic>' +
      '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas">' +
      '<wpc:wpc><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Picture"/>' +
      '<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1" cstate="print"><a:extLst>' +
      '<a:ext uri="28A0092B-C50C-407e-A947-70E740481C1C"><a14:useLocalDpi val="0"/>' +
      "</a:ext></a:extLst></a:blip></pic:blipFill></pic:pic></wpc:wpc>" +
      "</a:graphicData></a:graphic></wp:inline></w:drawing>";
    const result = parseDrawingRun(parseXml(xml).elements![0]!, mediaContext);
    expect(result).toMatchObject({
      wordprocessingCanvas: {
        children: [
          {
            compression: "print",
            useLocalDpi: false,
            useLocalDpiUri: "28A0092B-C50C-407e-A947-70E740481C1C",
          },
        ],
      },
    });
  });

  it("preserves group property extensions", () => {
    const xml =
      `<w:drawing ${NS} xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ` +
      'xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main">' +
      '<wp:inline><wp:extent cx="9525" cy="9525"/><a:graphic>' +
      '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
      '<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/>' +
      '<a:ext cx="9525" cy="9525"/></a:xfrm><a:extLst><a:ext uri="{group-flag}">' +
      '<ma14:wrappingTextBoxFlag val="1"/></a:ext></a:extLst></wpg:grpSpPr></wpg:wgp>' +
      "</a:graphicData></a:graphic></wp:inline></w:drawing>";
    const result = parseDrawingRun(parseXml(xml).elements![0]!, mediaContext);
    expect(result).toMatchObject({
      wpgGroup: { extensions: [{ uri: "{group-flag}" }] },
    });
  });

  it("preserves inline wps text box ids", () => {
    const xml =
      `<w:drawing ${NS} xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
      '<wp:inline><wp:extent cx="9525" cy="9525"/><a:graphic>' +
      '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/>' +
      '<a:ext cx="9525" cy="9525"/></a:xfrm></wps:spPr>' +
      '<wps:txbx id="7"><w:txbxContent><w:p><w:r><w:t>Box</w:t></w:r></w:p>' +
      "</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>" +
      "</a:graphicData></a:graphic></wp:inline></w:drawing>";
    const result = parseDrawingRun(parseXml(xml).elements![0]!, mediaContext);
    expect(result).toMatchObject({ wpsShape: { textBoxId: 7 } });
  });
});
