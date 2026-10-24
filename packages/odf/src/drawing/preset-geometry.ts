/**
 * LibreOffice preset geometry names (ODF draw:type) mapped to ECMA ST_ShapeType
 * tokens. Sourced from LibreOffice core svx EnhancedCustomShapeTypeNames and
 * the ECMA preset token list; names without an OOXML counterpart (text boxes,
 * non-primitive geometry) are intentionally absent and fall back to no
 * geometry on conversion.
 *
 * @module
 */

import type { ShapeType, TextShapePreset } from "@office-open/core/drawing";
import { MSO_SHAPE_TYPE_PRESETS } from "@office-open/core/drawing";

export const PRESET_GEOMETRY_DOCX: Record<string, ShapeType> = {
  "bent-arrow": "bentArrow",
  "bent-up-arrow": "bentUpArrow",
  "block-arc": "blockArc",
  "brace-pair": "bracePair",
  "bracket-pair": "bracketPair",
  can: "can",
  chevron: "chevron",
  circle: "ellipse",
  "circular-arrow": "circularArrow",
  cloud: "cloud",
  "cloud-callout": "cloudCallout",
  forbidden: "noSmoking",
  cross: "plus",
  cube: "cube",
  "curved-down-arrow": "curvedDownArrow",
  "curved-left-arrow": "curvedLeftArrow",
  "curved-right-arrow": "curvedRightArrow",
  "curved-up-arrow": "curvedUpArrow",
  diamond: "diamond",
  "diamond-callout": "wedgeRectCallout",
  "double-wave": "doubleWave",
  "down-arrow": "downArrow",
  "down-arrow-callout": "downArrowCallout",
  "eight-pointed-star": "star8",
  ellipse: "ellipse",
  "flowchart-alternate-process": "flowChartAlternateProcess",
  "flowchart-connector": "flowChartConnector",
  "flowchart-data": "flowChartInputOutput",
  "flowchart-decision": "flowChartDecision",
  "flowchart-delay": "flowChartDelay",
  "flowchart-direct-access-storage": "flowChartMagneticDisk",
  "flowchart-display": "flowChartDisplay",
  "flowchart-document": "flowChartDocument",
  "flowchart-internal-storage": "flowChartInternalStorage",
  "flowchart-manual-operation": "flowChartManualOperation",
  "flowchart-merge": "flowChartMerge",
  "flowchart-multidocument": "flowChartMultidocument",
  "flowchart-off-page-connector": "flowChartOffpageConnector",
  "flowchart-predefined-process": "flowChartPredefinedProcess",
  "flowchart-preparation": "flowChartPreparation",
  "flowchart-process": "flowChartProcess",
  "flowchart-summing-junction": "flowChartSummingJunction",
  "flowchart-terminator": "flowChartTerminator",
  "folded-corner": "foldedCorner",
  frame: "frame",
  funnel: "funnel",
  "half-frame": "snip2DiagRect",
  heart: "heart",
  hexagon: "hexagon",
  "horizontal-scroll": "horizontalScroll",
  "col-502ad400": "frame",
  "col-60da8460": "frame",
  "isosceles-triangle": "triangle",
  "left-arrow": "leftArrow",
  "left-arrow-callout": "leftArrowCallout",
  "left-brace": "leftBrace",
  "left-bracket": "leftBracket",
  "left-right-arrow": "leftRightArrow",
  "left-right-arrow-callout": "leftRightArrowCallout",
  lightning: "lightningBolt",
  megaphone: "wedgeEllipseCallout",
  moon: "moon",
  "notched-right-arrow": "notchedRightArrow",
  octagon: "octagon",
  parallelogram: "parallelogram",
  pentagon: "pentagon",
  "pentagon-right": "homePlate",
  "quad-arrow": "quadArrow",
  "quad-arrow-callout": "quadArrowCallout",
  rectangle: "rect",
  "rectangular-callout": "wedgeRectCallout",
  ribbon: "ribbon",
  "ribbon-2": "ribbon2",
  "right-arrow": "rightArrow",
  "right-arrow-callout": "rightArrowCallout",
  "right-brace": "rightBrace",
  "right-bracket": "rightBracket",
  "right-triangle": "rtTriangle",
  ring: "donut",
  "round-rectangle": "roundRect",
  "rounded-rectangular-callout": "wedgeRoundRectCallout",
  sheet: "plaque",
  "six-pointed-star": "star6",
  "sixteen-pointed-star": "star16",
  smiley: "smileyFace",
  "smiley-face": "smileyFace",
  star24: "star24",
  star4: "star4",
  star5: "star5",
  star8: "star8",
  "striped-right-arrow": "stripedRightArrow",
  sun: "sun",
  teardrop: "teardrop",
  "ten-pointed-star": "star10",
  "thirty-two-pointed-star": "star32",
  trapezoid: "trapezoid",
  "twelve-pointed-star": "star12",
  "twenty-four-pointed-star": "star24",
  "up-arrow": "upArrow",
  "up-arrow-callout": "upArrowCallout",
  "up-down-arrow": "upDownArrow",
  "up-down-arrow-callout": "upDownArrowCallout",
  "vertical-scroll": "verticalScroll",
  wave: "wave",
};

const ODF_TEXT_WARP_GEOMETRY: Record<string, TextShapePreset> = {
  "fontwork-plain-text": "textPlain",
  "fontwork-stop": "textStop",
  "fontwork-triangle-up": "textTriangle",
  "fontwork-triangle-down": "textTriangleInverted",
  "fontwork-chevron-up": "textChevron",
  "fontwork-chevron-down": "textChevronInverted",
  "mso-spt142": "textRingInside",
  "mso-spt143": "textRingOutside",
  "fontwork-arch-up-curve": "textArchUp",
  "fontwork-arch-down-curve": "textArchDown",
  "fontwork-circle-curve": "textCircle",
  "fontwork-open-circle-curve": "textButton",
  "fontwork-arch-up-pour": "textArchUpPour",
  "fontwork-arch-down-pour": "textArchDownPour",
  "fontwork-circle-pour": "textCirclePour",
  "fontwork-open-circle-pour": "textButtonPour",
  "fontwork-curve-up": "textCurveUp",
  "fontwork-curve-down": "textCurveDown",
  "mso-spt174": "textCanUp",
  "mso-spt175": "textCanDown",
  "fontwork-wave": "textWave1",
  "mso-spt157": "textWave2",
  "mso-spt158": "textDoubleWave1",
  "mso-spt159": "textWave4",
  "fontwork-inflate": "textInflate",
  "mso-spt161": "textDeflate",
  "mso-spt162": "textInflateBottom",
  "mso-spt163": "textDeflateBottom",
  "mso-spt164": "textInflateTop",
  "mso-spt165": "textDeflateTop",
  "mso-spt166": "textDeflateInflate",
  "mso-spt167": "textDeflateInflateDeflate",
  "fontwork-fade-right": "textFadeRight",
  "fontwork-fade-left": "textFadeLeft",
  "fontwork-fade-up": "textFadeUp",
  "fontwork-fade-down": "textFadeDown",
  "fontwork-slant-up": "textSlantUp",
  "fontwork-slant-down": "textSlantDown",
  "fontwork-fade-up-and-right": "textCascadeUp",
  "fontwork-fade-up-and-left": "textCascadeDown",
  "mso-spt202": "textPlain",
};

export type OdfDrawGeometry =
  | { kind: "custom" }
  | { kind: "shape"; preset: ShapeType }
  | { kind: "text-warp"; preset: TextShapePreset };

/** Resolve an ODF draw:type into its canonical OOXML-equivalent projection. */
export function resolveOdfDrawGeometry(drawType: string): OdfDrawGeometry {
  if (drawType === "non-primitive" || drawType === "ooxml-non-primitive") return { kind: "custom" };
  if (Object.hasOwn(ODF_TEXT_WARP_GEOMETRY, drawType))
    return { kind: "text-warp", preset: ODF_TEXT_WARP_GEOMETRY[drawType]! };
  const preset = Object.hasOwn(PRESET_GEOMETRY_DOCX, drawType)
    ? PRESET_GEOMETRY_DOCX[drawType]
    : undefined;
  if (preset) return { kind: "shape", preset };
  if (drawType.startsWith("ooxml-")) {
    const token = drawType.slice("ooxml-".length);
    const alias = token === "rect" ? "rect" : token;
    if (Object.hasOwn(PRESET_GEOMETRY_DOCX, alias))
      return { kind: "shape", preset: PRESET_GEOMETRY_DOCX[alias]! };
    if ((["triangle", "ellipse", "uturnArrow", "circularArrow"] as const).includes(alias as never))
      return { kind: "shape", preset: alias as ShapeType };
  }
  const numeric = drawType.match(/^ms?o-spt(\d+)$/u);
  if (numeric) {
    const identifier = Number(numeric[1]);
    const msoPreset = MSO_SHAPE_TYPE_PRESETS[identifier];
    if (msoPreset) return { kind: "shape", preset: msoPreset };
    const textPreset = ODF_TEXT_WARP_GEOMETRY[`mso-spt${identifier}`];
    if (textPreset) return { kind: "text-warp", preset: textPreset };
  }
  return { kind: "custom" };
}

/** Return an overlay only when canonical geometry cannot restore this exact draw:type. */
export function resolveOdfDrawGeometryOverlay(
  drawType: string,
  geometry: OdfDrawGeometry,
): string | undefined {
  if (geometry.kind !== "shape") return drawType;
  return presetGeometryOdf(geometry.preset) === drawType ? undefined : drawType;
}

/** Reverse lookup: ST_ShapeType token → the first matching ODF preset name. */
export function presetGeometryOdf(token: ShapeType): string | undefined {
  for (const [odf, mapped] of Object.entries(PRESET_GEOMETRY_DOCX))
    if (mapped === token) return odf;
  return undefined;
}
