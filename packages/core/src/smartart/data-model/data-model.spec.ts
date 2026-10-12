import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseDataModelOptions, stringifyDataModelOptions } from "./data-model";
import type { DataModelOptions } from "./data-model";

function roundTrip(model: DataModelOptions): DataModelOptions {
  const xml = parseXml(stringifyDataModelOptions(model));
  const root = xml.elements?.[0];
  expect(root?.name).toBe("dgm:dataModel");
  return parseDataModelOptions(root)!;
}

describe("SmartArt data model", () => {
  it("round-trips doc, node, transition, and connection surfaces", () => {
    const model: DataModelOptions = {
      points: [
        {
          modelId: "0",
          type: "doc",
          propertySet: {
            presentationAssociationId: "pres-assoc",
            presentationName: "presentation",
            presentationStyleLabel: "node0",
            presentationStyleIndex: 1,
            presentationStyleCount: 2,
            placeholderText: "text",
            placeholder: true,
            customAngle: 30,
            customFlipVertical: true,
            customFlipHorizontal: false,
            customSizeX: 10,
            customSizeY: 20,
            customScaleX: 30,
            customScaleY: 40,
            customText: true,
            customLinearFactorX: 50,
            customLinearFactorY: 25,
            customLinearFactorNeighborX: 75,
            customLinearFactorNeighborY: 80,
            customRadialScaleRadius: 150,
            customRadialScaleIncrement: 250,
            coherent3DOffset: true,
            presentationLayoutVariables: {
              orgChart: { val: true },
              maxChildren: { val: 4 },
              bulletEnabled: { val: false },
              direction: { val: "reversed" },
              resizeHandles: { val: "relative" },
            },
            style: {
              lineReference: { index: 2, color: { value: "accent1" } },
              fillReference: { index: 1 },
              effectReference: { index: 0 },
              fontReference: { collection: "minor" },
            },
            layoutTypeId: "layout/1",
            layoutCategoryId: "layout/cat",
            quickStyleTypeId: "style/1",
            quickStyleCategoryId: "style/cat",
            colorStyleTypeId: "color/1",
            colorStyleCategoryId: "color/cat",
          },
          shapeProperties: { x: 10, y: 20, width: 30, height: 40 },
          textBody: {
            bodyProperties: {},
            paragraphs: [{ endParagraphProperties: false, text: "Document" }],
          },
          extensions: [{ uri: "urn:doc-ext", content: "<foo:payload/>" }],
        },
        {
          modelId: "1",
          textBody: {
            bodyProperties: {},
            paragraphs: [{ endParagraphProperties: false, text: "One" }],
          },
          shapeProperties: {
            geometry: { preset: "rect" },
            fill: { type: "solid", color: { value: "FF0000" } },
          },
        },
        {
          modelId: "2",
          type: "parTrans",
          connectionId: "3",
          textBody: {
            bodyProperties: {},
            paragraphs: [{ endParagraphProperties: { lang: "en-US" } }],
          },
        },
      ],
      connections: [
        {
          modelId: "3",
          sourceId: "0",
          destinationId: "1",
          type: "presParOf",
          sourceOrder: 2,
          destinationOrder: 4,
          parentTransitionId: "5",
          siblingTransitionId: "6",
          presentationId: "pres-1",
          extensions: [{ uri: "urn:cxn-ext" }],
        },
      ],
      background: {
        fill: { type: "solid", color: { value: "00FF00" } },
        effects: { blur: { radius: 100 } },
      },
      whole: {
        outline: { width: 12700, type: "solidFill", color: { value: "0000FF" } },
        effects: { glow: { radius: 200, color: { value: "FFFFFF" } } },
      },
      extensions: [{ uri: "urn:model-ext" }],
    };

    expect(roundTrip(model)).toEqual(model);
  });

  it("preserves an empty point property set", () => {
    const model: DataModelOptions = {
      points: [
        { modelId: "0", type: "doc" },
        { modelId: "1", propertySet: {} },
      ],
    };
    expect(roundTrip(model)).toEqual(model);
  });

  it("round-trips non-finite rule values and decimal percentages", () => {
    const model: DataModelOptions = {
      points: [
        {
          modelId: "0",
          propertySet: { customLinearFactorNeighborX: 34.137, customLinearFactorNeighborY: 12.5 },
        },
      ],
    };
    expect(roundTrip(model)).toEqual(model);
  });

  it("preserves empty category ids, black-white mode, and image fills", () => {
    const image = new Uint8Array([1, 2, 3]);
    const source = parseXml(
      '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" ' +
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<dgm:ptLst><dgm:pt modelId="1"><dgm:prSet loCatId=""/>' +
        '<dgm:spPr bwMode="gray"><a:blipFill rotWithShape="0">' +
        '<a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>' +
        "</dgm:spPr></dgm:pt></dgm:ptLst></dgm:dataModel>",
    ).elements?.[0];
    const model = parseDataModelOptions(source, {
      resolveRelationship: () => "word/media/image.png",
      getPart: () => undefined,
      getRaw: () => image,
    })!;

    expect(model.points?.[0]?.propertySet?.layoutCategoryId).toBe("");
    expect(model.points?.[0]?.blackWhiteMode).toBe("gray");
    expect(model.points?.[0]?.shapeProperties?.fill).toMatchObject({
      type: "blip",
      rotWithShape: false,
      fillRectangle: {},
    });
    expect(stringifyDataModelOptions(model)).toContain('loCatId=""');
    expect(stringifyDataModelOptions(model)).toContain('<dgm:spPr bwMode="gray">');
  });
});
