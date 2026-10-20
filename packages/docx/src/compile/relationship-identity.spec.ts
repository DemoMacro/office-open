import { describe, expect, it } from "vite-plus/test";

import { DocxWriteContext } from "../context";
import type { MediaData } from "../shared/media";
import { compileDocumentEntries } from "./document";
import { compileChartExParts, compileSmartArtParts } from "./drawings";

const MEDIA = new Uint8Array([1, 2, 3, 4]);

function decode(data: string | Uint8Array): string {
  return typeof data === "string" ? data : new TextDecoder().decode(data);
}

describe("drawing relationship identity", () => {
  it("claims equivalent document media target spellings", () => {
    for (const target of ["../media/image.png", "media/image.png", "/media/image.png"]) {
      const ctx = new DocxWriteContext({
        sections: [],
        primaryPartPath: "document.xml",
        passthroughRelationships: [
          {
            source: "document.xml",
            relationshipType: "http://example.com/relationships/image",
            target,
            rId: "rId9",
          },
        ],
      });
      ctx.media.addMedia(
        MEDIA,
        "png",
        (fileName): MediaData => ({
          type: "png",
          data: MEDIA,
          fileName,
          partPath: "media/image.png",
          transformation: { width: 1, height: 1 },
        }),
        "image.png",
      );

      const { Document, Relationships } = compileDocumentEntries(ctx, "<w:p>{image.png}</w:p>", 1);

      expect(decode(Document.data)).toBe("<w:p>rId9</w:p>");
      expect(decode(Relationships.data)).toContain(
        `Id="rId9" Type="http://example.com/relationships/image" Target="${target}"`,
      );
      expect(decode(Relationships.data).split(`Target="${target}"`).length - 1).toBe(1);
    }
  });

  it("normalizes duplicate image relationship ids to the claimed source id", () => {
    const imageType = "http://example.com/relationships/image";
    const ctx = new DocxWriteContext({
      sections: [],
      passthroughRelationships: [
        {
          source: "word/document.xml",
          relationshipType: imageType,
          target: "media/image.png",
          rId: "rId2",
        },
        {
          source: "word/document.xml",
          relationshipType: imageType,
          target: "media/image.png",
          rId: "rId3",
        },
      ],
    });
    ctx.media.addMedia(
      MEDIA,
      "png",
      (fileName): MediaData => ({
        type: "png",
        data: MEDIA,
        fileName,
        transformation: { width: 1, height: 1 },
      }),
      "image.png",
    );

    const { Document, Relationships } = compileDocumentEntries(ctx, "<w:p>{image.png}</w:p>", 1);

    expect(decode(Document.data)).toContain("rId2");
    expect(decode(Relationships.data)).toContain(
      'Id="rId2" Type="http://example.com/relationships/image" Target="media/image.png"',
    );
    expect(decode(Relationships.data)).not.toContain('Id="rId3"');
  });

  it("reuses one ChartEx source relationship across identity spellings", () => {
    const ctx = new DocxWriteContext({
      sections: [],
      passthroughRelationships: [
        {
          source: "word/charts/chartEx1.xml",
          relationshipType: "http://example.com/relationships/chartStyle",
          target: "/word/charts/style1.xml",
          rId: "rId6",
        },
      ],
    });
    ctx.chartExes.addChartEx("chartEx_1", {
      key: "chartEx_1",
      chartExXml: "<cx:chartSpace/>",
      sourcePath: "word/charts/chartEx1.xml",
      sourceRelationships: [
        {
          relationshipType: "http://example.com/relationships/chartStyle",
          target: "style1.xml",
          rId: "rId5",
        },
      ],
    });

    const chartExParts = compileChartExParts(ctx).ChartEx ?? [];
    const rels = chartExParts.find((part) => part.path === "word/charts/_rels/chartEx1.xml.rels");
    expect(rels).toBeDefined();
    const xml = decode(rels?.data ?? "");
    expect(xml).toContain(
      'Id="rId5" Type="http://example.com/relationships/chartStyle" Target="/word/charts/style1.xml"',
    );
    expect(xml.match(/<Relationship /g)).toHaveLength(1);
  });

  it("preserves a ChartEx part declared by a chart relationship", () => {
    const ctx = new DocxWriteContext({
      sections: [],
      passthroughRelationships: [
        {
          source: "word/document.xml",
          relationshipType:
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
          target: "charts/chart4.xml",
          rId: "rId8",
        },
      ],
    });
    ctx.chartExes.addChartEx("chartEx_1", {
      key: "chartEx_1",
      chartExXml: "<cx:chartSpace/>",
      sourcePath: "word/charts/chart4.xml",
    });

    const { Document, Relationships } = compileDocumentEntries(
      ctx,
      "<w:p>{chartex:chartEx_1}</w:p>",
      0,
    );
    const xml = decode(Relationships.data);
    expect(decode(Document.data)).toContain("rId8");
    expect(xml).toContain(
      'Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"',
    );
    expect(xml).not.toContain("/relationships/chartEx");
  });

  it("preserves SmartArt data-part relationship ids and embedded rIds", () => {
    const ctx = new DocxWriteContext({ sections: [] });
    ctx.smartArts.addSmartArt("smartart_1", {
      key: "smartart_1",
      dataModelXml: '<dgm:dataModel><a:blip r:embed="{image.png}"/></dgm:dataModel>',
      layout: "default",
      style: "simple1",
      color: "accent1_2",
      sourcePaths: {
        data: "word/diagrams/data1.xml",
        drawing: "word/diagrams/drawing1.xml",
      },
      sourceRelationships: [
        {
          relationshipType: "http://example.com/relationships/diagramDrawing",
          target: "drawing1.xml",
          rId: "rId4",
        },
        {
          relationshipType: "http://example.com/relationships/image",
          target: "/word/media/image.png",
          rId: "rId7",
        },
      ],
      raw: {
        drawing: MEDIA,
        drawingRels: new TextEncoder().encode(
          '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image.png"/></Relationships>',
        ),
        drawingRelationship: true,
        media: [{ fileName: "image.png", data: MEDIA, partPath: "word/media/image.png" }],
      },
    });

    const parts = compileSmartArtParts(ctx);
    const data = parts.DiagramData?.[0];
    const rels = parts.SmartArtDataRels?.[0];
    expect(data).toBeDefined();
    expect(rels).toBeDefined();
    expect(decode(data?.data ?? "")).toContain('r:embed="rId7"');
    const relsXml = decode(rels?.data ?? "");
    expect(relsXml).toContain(
      'Id="rId4" Type="http://example.com/relationships/diagramDrawing" Target="drawing1.xml"',
    );
    expect(relsXml).toContain(
      'Id="rId7" Type="http://example.com/relationships/image" Target="/word/media/image.png"',
    );
    expect(relsXml.match(/<Relationship /g)).toHaveLength(2);
    const drawingParts = parts.DiagramDrawing ?? [];
    expect(drawingParts.map((part) => part.path)).toEqual([
      "word/diagrams/drawing1.xml",
      "word/diagrams/_rels/drawing1.xml.rels",
    ]);
    expect(decode(drawingParts[1]?.data ?? "")).toContain('Id="rId1"');
  });

  it("preserves SmartArt document relationship ids in body and rels", () => {
    const passthroughRelationships = [
      {
        source: "word/document.xml",
        relationshipType:
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
        target: "diagrams/data1.xml",
        rId: "rId20",
      },
      {
        source: "word/document.xml",
        relationshipType:
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
        target: "diagrams/layout1.xml",
        rId: "rId21",
      },
      {
        source: "word/document.xml",
        relationshipType:
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
        target: "diagrams/quickStyle1.xml",
        rId: "rId22",
      },
      {
        source: "word/document.xml",
        relationshipType:
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
        target: "diagrams/colors1.xml",
        rId: "rId23",
      },
    ];
    const ctx = new DocxWriteContext({ sections: [], passthroughRelationships });
    ctx.smartArts.addSmartArt("smartart_1", {
      key: "smartart_1",
      dataModelXml: "<dgm:dataModel/>",
      layout: "default",
      style: "simple1",
      color: "accent1_2",
      sourcePaths: {
        data: "word/diagrams/data1.xml",
        layout: "word/diagrams/layout1.xml",
        quickStyle: "word/diagrams/quickStyle1.xml",
        colors: "word/diagrams/colors1.xml",
      },
    });

    const { Document, Relationships } = compileDocumentEntries(
      ctx,
      "<w:p>{smartart:smartart_1}{smartart-lo:smartart_1}{smartart-qs:smartart_1}{smartart-cs:smartart_1}</w:p>",
      1,
    );
    const body = decode(Document.data);
    const rels = decode(Relationships.data);
    for (const rId of ["rId20", "rId21", "rId22", "rId23"]) {
      expect(body).toContain(rId);
      expect(rels).toContain(`Id="${rId}"`);
    }
  });
});
