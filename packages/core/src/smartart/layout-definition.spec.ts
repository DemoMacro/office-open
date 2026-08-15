import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { CustomDescriptor, ReadContext, WriteContext } from "../descriptor";
import { colorsDefDesc, type ColorDefinitionOptions } from "./color-definition";
import { layoutDefDesc, type LayoutDefinitionOptions } from "./layout-definition";
import { styleDefDesc, type StyleDefinitionOptions } from "./style-definition";

function roundTrip<T>(desc: CustomDescriptor<T>, opts: T): T {
  const xml = desc.stringify(opts, {} as WriteContext);
  if (!xml) throw new Error("stringify returned empty");
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return desc.parse(el, {} as ReadContext);
}

describe("layoutDefDesc", () => {
  it("round-trips a full layout definition", () => {
    const opts: LayoutDefinitionOptions = {
      uniqueId: "urn:microsoft.com/office/officeart/2005/8/layout/custom1",
      minVer: "http://schemas.openxmlformats.org/drawingml/2006/diagram",
      defaultStyle: "nofill",
      titles: [{ lang: "en-US", val: "Custom Process" }],
      descriptions: [{ val: "A custom process layout" }],
      categories: [{ type: "process", pri: 1000 }],
      sampleData: {
        dataModel: {
          points: [
            { modelId: "1", text: "One", type: "node" },
            { modelId: "2", text: "Two", type: "node" },
          ],
          connections: [
            {
              modelId: "3",
              sourceId: "1",
              destinationId: "2",
              type: "parOf",
              sourceOrder: 0,
              destinationOrder: 0,
            },
          ],
        },
      },
      layoutNode: {
        name: "diagram",
        styleLabel: "node1",
        childOrder: "t",
        children: [
          {
            variables: {
              organizationChart: true,
              direction: "rev",
              animateOne: "branch",
            },
          },
          {
            algorithm: {
              type: "snake",
              revision: 2,
              parameters: [
                { type: "grDir", value: "tL" },
                { type: "numCol", value: 3 },
                { type: "animBg", value: true },
              ],
            },
          },
          {
            constraints: [
              { type: "w", for: "ch", factor: 0.5 },
              { type: "h", referenceType: "w", operation: "equ", value: 100 },
            ],
          },
          { rules: [{ type: "primFontSz", value: 50, maximum: 90 }] },
          {
            forEach: {
              name: "nodes",
              axis: "ch",
              pointType: "node",
              start: "0",
              count: "5",
              step: "1",
              children: [
                {
                  layoutNode: {
                    name: "node",
                    moveWith: "diagram",
                    children: [
                      {
                        shape: {
                          type: "roundRect",
                          rotation: 15.5,
                          zOrderOffset: 2,
                          hideGeometry: true,
                          adjustments: [{ idx: 1, val: 0.2 }],
                        },
                      },
                      { presentationOf: { axis: "des", pointType: "node" } },
                    ],
                  },
                },
              ],
            },
          },
          {
            choose: {
              name: "branch",
              conditions: [
                {
                  function: "var",
                  argument: "dir",
                  operator: "equ",
                  value: "1",
                  children: [{ layoutNode: { name: "horizontal" } }],
                },
              ],
              otherwise: { children: [{ layoutNode: { name: "vertical" } }] },
            },
          },
        ],
      },
    };
    expect(roundTrip(layoutDefDesc, opts)).toEqual(opts);
  });

  it("round-trips a minimal layout definition", () => {
    const opts: LayoutDefinitionOptions = { layoutNode: { children: [] } };
    expect(roundTrip(layoutDefDesc, opts)).toEqual({ layoutNode: {} });
  });
});

describe("styleDefDesc", () => {
  it("round-trips style labels with 3D and style-matrix content", () => {
    const scene = {
      camera: { preset: "orthographicFront" },
      lightRig: { rig: "threePt", direction: "tl" },
    };
    const style = {
      lineReference: { index: 2, color: { value: "accent1" } },
      fillReference: { index: 1 },
      effectReference: { index: 0 },
      fontReference: { collection: "minor" as const },
    };
    const opts: StyleDefinitionOptions = {
      uniqueId: "urn:microsoft.com/office/officeart/2005/8/quickstyle/custom1",
      titles: [{ val: "Custom Style" }],
      categories: [{ type: "simple", pri: 100 }],
      styleLabels: [
        {
          name: "node0",
          scene3d: scene,
          shape3d: { z: 100, extrusionH: 500 },
          textProperties: { flatText: 0 },
          style,
        },
        { name: "node1", textProperties: { shape3d: { z: 50 } } },
      ],
    };
    expect(roundTrip(styleDefDesc, opts)).toEqual(opts);
  });
});

describe("colorsDefDesc", () => {
  it("round-trips color style labels with color lists", () => {
    const opts: ColorDefinitionOptions = {
      uniqueId: "urn:microsoft.com/office/officeart/2005/8/colors/custom1",
      titles: [{ lang: "en-US", val: "Custom Colors" }],
      styleLabels: [
        {
          name: "node0",
          fillColorList: {
            meth: "cycle",
            hueDir: "ccw",
            colors: [{ value: "accent1" }, { value: "#FF0000" }],
          },
          lineColorList: { colors: [{ value: "lt1" }] },
          textFillColorList: { meth: "repeat" },
        },
      ],
    };
    expect(roundTrip(colorsDefDesc, opts)).toEqual(opts);
  });
});
