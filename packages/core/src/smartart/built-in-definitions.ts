/**
 * SmartArt built-in definitions — layout, quick style, and color transforms.
 *
 * Emits layoutDef, styleDef, and colorsDef parts from typed constants through
 * the structured descriptors. All entries carry only the uniqueId plus a
 * schema-valid minimal body — Office applications resolve these to built-in
 * definitions; the "default" layout ships its full layout tree.
 *
 * @module
 */

export { COLOR_CATEGORIES, LAYOUT_CATEGORIES, STYLE_CATEGORIES } from "./categories";

import type { ColorListOptions } from "../drawing/diagram/diagram-style";
import { COLOR_CATEGORIES, LAYOUT_CATEGORIES, STYLE_CATEGORIES } from "./categories";
import { stringifyColorDefinitionPart, type ColorDefinitionOptions } from "./color-definition";
import { stringifyLayoutDefinitionPart, type LayoutDefinitionOptions } from "./layout-definition";
import { stringifyStyleDefinitionPart, type StyleDefinitionOptions } from "./style-definition";

// ---------------------------------------------------------------------------
// Layout XML — full for "default", stubs for everything else
// ---------------------------------------------------------------------------

const BUILTIN_ACCENTS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const;

/** Sample-model point list shared by the layout stubs: a doc point plus placeholders. */
function samplePoints(count: number): LayoutDefinitionOptions["sampleData"] {
  const points = [
    { modelId: "0", type: "doc" },
    ...Array.from({ length: count }, (_, i) => ({
      modelId: String(i + 1),
      propertySet: { placeholder: true },
    })),
  ];
  const connections = Array.from({ length: count }, (_, i) => ({
    modelId: String(count + 1 + i),
    sourceId: "0",
    destinationId: String(i + 1),
    sourceOrder: i,
    destinationOrder: 0,
  }));
  return { dataModel: { points, connections } };
}

/**
 * Full default list layout (urn:microsoft.com/office/officeart/2005/8/layout/default).
 */
const DEFAULT_LAYOUT: LayoutDefinitionOptions = {
  uniqueId: "urn:microsoft.com/office/officeart/2005/8/layout/default",
  titles: [{ val: "" }],
  descriptions: [{ val: "" }],
  categories: [{ type: "list", pri: 400 }],
  sampleData: samplePoints(5),
  styleData: {
    dataModel: {
      points: [{ modelId: "0", type: "doc" }, { modelId: "1" }, { modelId: "2" }],
      connections: [
        { modelId: "3", sourceId: "0", destinationId: "1", sourceOrder: 0, destinationOrder: 0 },
        { modelId: "4", sourceId: "0", destinationId: "2", sourceOrder: 1, destinationOrder: 0 },
      ],
    },
  },
  colorData: {
    dataModel: {
      points: [
        { modelId: "0", type: "doc" },
        ...Array.from({ length: 6 }, (_, i) => ({ modelId: String(i + 1) })),
      ],
      connections: Array.from({ length: 6 }, (_, i) => ({
        modelId: String(7 + i),
        sourceId: "0",
        destinationId: String(i + 1),
        sourceOrder: i,
        destinationOrder: 0,
      })),
    },
  },
  layoutNode: {
    name: "diagram",
    childOrder: "b",
    children: [
      { variables: { direction: "normal", resizeHandles: "exact" } },
      {
        choose: {
          name: "Name0",
          conditions: [
            {
              name: "Name1",
              function: "var",
              argument: "dir",
              operator: "equ",
              value: "norm",
              children: [
                {
                  algorithm: {
                    type: "snake",
                    parameters: [
                      { type: "grDir", value: "tL" },
                      { type: "flowDir", value: "row" },
                      { type: "contDir", value: "sameDir" },
                      { type: "off", value: "ctr" },
                    ],
                  },
                },
              ],
            },
          ],
          otherwise: {
            name: "Name2",
            children: [
              {
                algorithm: {
                  type: "snake",
                  parameters: [
                    { type: "grDir", value: "tR" },
                    { type: "flowDir", value: "row" },
                    { type: "contDir", value: "sameDir" },
                    { type: "off", value: "ctr" },
                  ],
                },
              },
            ],
          },
        },
      },
      { shape: { blipPlaceholder: false, adjustments: [] } },
      { presentationOf: {} },
      {
        constraints: [
          { type: "w", for: "ch", forName: "node", referenceType: "w", referencePointType: "all" },
          {
            type: "h",
            for: "ch",
            forName: "node",
            referenceType: "w",
            referenceFor: "ch",
            referenceForName: "node",
            factor: 0.6,
          },
          {
            type: "w",
            for: "ch",
            forName: "sibTrans",
            referenceType: "w",
            referenceFor: "ch",
            referenceForName: "node",
            factor: 0.1,
          },
          {
            type: "sp",
            referenceType: "w",
            referenceFor: "ch",
            referenceForName: "sibTrans",
          },
          { type: "primFontSz", for: "ch", forName: "node", operation: "equ", value: 65 },
        ],
      },
      { rules: [] },
      {
        forEach: {
          name: "Name3",
          axis: "ch",
          pointType: "node",
          start: "1",
          children: [
            {
              layoutNode: {
                name: "node",
                children: [
                  { variables: { bulletEnabled: true } },
                  { algorithm: { type: "tx" } },
                  { shape: { type: "rect", adjustments: [] } },
                  { presentationOf: { axis: "desOrSelf", pointType: "node" } },
                  {
                    constraints: [
                      { type: "lMarg", referenceType: "primFontSz", factor: 0.3 },
                      { type: "rMarg", referenceType: "primFontSz", factor: 0.3 },
                      { type: "tMarg", referenceType: "primFontSz", factor: 0.3 },
                      { type: "bMarg", referenceType: "primFontSz", factor: 0.3 },
                    ],
                  },
                  // rule @fact/@max default to NaN in the XSD, so omitting them
                  // means "unbounded" — same as the explicit NaN in Office's copy.
                  { rules: [{ type: "primFontSz", value: 5 }] },
                ],
              },
            },
            {
              forEach: {
                name: "Name4",
                axis: "followSib",
                pointType: "sibTrans",
                count: "1",
                children: [
                  {
                    layoutNode: {
                      name: "sibTrans",
                      children: [
                        { algorithm: { type: "sp" } },
                        { shape: { adjustments: [] } },
                        { presentationOf: {} },
                        { constraints: [] },
                        { rules: [] },
                      ],
                    },
                  },
                ],
              },
            },
          ],
        },
      },
    ],
  },
};

/**
 * Returns layout XML. Full layout tree for "default", minimal stub for others.
 * Stubs keep a schema-required empty layoutNode; PowerPoint resolves built-in
 * definitions from the uniqueId / loTypeId in the data model.
 */
export function getLayoutXml(layoutId: string): string {
  if (layoutId === "default") return stringifyLayoutDefinitionPart(DEFAULT_LAYOUT);
  const layout: LayoutDefinitionOptions = {
    uniqueId: `urn:microsoft.com/office/officeart/2005/8/layout/${layoutId}`,
    titles: [{ val: "" }],
    descriptions: [{ val: "" }],
    categories: [{ type: LAYOUT_CATEGORIES[layoutId] ?? "list", pri: 400 }],
    sampleData: samplePoints(2),
    layoutNode: {},
  };
  return stringifyLayoutDefinitionPart(layout);
}

// ---------------------------------------------------------------------------
// Style XML
// ---------------------------------------------------------------------------

const STYLE_SCENE_3D = {
  camera: { preset: "orthographicFront" },
  lightRig: { rig: "threePt", direction: "t" },
};

/**
 * Returns style XML. Each label node0…node5 ties one accent color into the
 * theme style matrix (line in tx1, fill in the accent), so PowerPoint renders
 * SmartArt shapes with themed fills and effects.
 */
export function getStyleXml(styleId: string): string {
  const style: StyleDefinitionOptions = {
    uniqueId: `urn:microsoft.com/office/officeart/2005/8/quickstyle/${styleId}`,
    titles: [{ val: "" }],
    descriptions: [{ val: "" }],
    categories: [{ type: STYLE_CATEGORIES[styleId] ?? "simple", pri: 10100 }],
    scene3d: STYLE_SCENE_3D,
    styleLabels: BUILTIN_ACCENTS.map((accent, i) => ({
      name: `node${i}`,
      style: {
        lineReference: { index: 2, color: { value: "tx1" } },
        fillReference: { index: 1, color: { value: accent } },
        effectReference: { index: 0, color: { value: "tx1" } },
        fontReference: { collection: "minor" as const, color: { value: "tx1" } },
      },
    })),
  };
  return stringifyStyleDefinitionPart(style);
}

// ---------------------------------------------------------------------------
// Color XML
// ---------------------------------------------------------------------------

const TINTED_TEXT = { value: "tx1", transforms: { tint: 75 } };

/**
 * Returns color XML mapping the node0 label to color lists for the requested
 * color scheme (accent / colorful / dark / fallback).
 */
export function getColorXml(colorId: string): string {
  const accent = colorId.match(/^(accent\d)/)?.[1] ?? "accent1";

  let fillColorList: ColorListOptions;
  let lineColorList: ColorListOptions;
  let textFillColorList: ColorListOptions;
  if (colorId.startsWith("colorful")) {
    fillColorList = { colors: BUILTIN_ACCENTS.slice(0, 5).map((a) => ({ value: a })) };
    lineColorList = { colors: [{ value: "lt1" }] };
    textFillColorList = { colors: [{ value: "tx1" }] };
  } else if (colorId.startsWith("dark")) {
    fillColorList = { colors: [{ value: "dk1" }, { value: "dk2" }] };
    lineColorList = { colors: [{ value: "lt1" }] };
    textFillColorList = { colors: [{ value: "lt1" }] };
  } else if (colorId.startsWith("accent")) {
    fillColorList = { colors: [{ value: accent }] };
    lineColorList = { colors: [TINTED_TEXT] };
    textFillColorList = { colors: [{ value: "tx1" }] };
  } else {
    fillColorList = { colors: [{ value: "accent1" }] };
    lineColorList = { colors: [{ value: "tx1" }] };
    textFillColorList = { colors: [{ value: "tx1" }] };
  }

  const colors: ColorDefinitionOptions = {
    uniqueId: `urn:microsoft.com/office/officeart/2005/8/colors/${colorId}`,
    titles: [{ val: "" }],
    descriptions: [{ val: "" }],
    categories: [{ type: COLOR_CATEGORIES[colorId] ?? "accent1", pri: 11200 }],
    styleLabels: [
      {
        name: "node0",
        fillColorList,
        lineColorList,
        effectColorList: lineColorList,
        textFillColorList,
        textLineColorList: colorId.startsWith("accent")
          ? { colors: [TINTED_TEXT] }
          : { colors: [] },
        textEffectColorList: textFillColorList,
      },
    ],
  };
  return stringifyColorDefinitionPart(colors);
}

/** Minimal drawing cache for SmartArt (Office apps auto-regenerate this on open) */
export const DEFAULT_DRAWING_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<dsp:drawing xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  "<dsp:spTree>" +
  '<dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr>' +
  "<dsp:grpSpPr/>" +
  "</dsp:spTree>" +
  "</dsp:drawing>";
