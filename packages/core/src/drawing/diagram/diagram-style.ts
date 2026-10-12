/**
 * Diagram style, color list, and style label elements.
 *
 * Reference: ISO/IEC 29500-4, dml-diagram.xsd
 * CT_Style, CT_StyleLabel, CT_Colors, CT_CTStyleLabel
 *
 * @module
 */
import { element } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { WriteContext } from "../../descriptor";
import { stringifyShapeStyle } from "../../theme/style-matrix";
import type { DefaultShapeStyleOptions } from "../../theme/theme-options";
import { xsdHueDirection } from "../../util/mappings";
import { parseColorChoiceElement } from "../color/color-descriptors";
import type { SolidFillOptions } from "../color/solid-fill";
import { createColorElement } from "../color/solid-fill";

// ---------------------------------------------------------------------------
// dgm:style — diagram style (CT_ShapeStyle from dml-main.xsd)
// ---------------------------------------------------------------------------

export type StyleMatrixIndex = "subtle" | "moderate" | "intense";

export const FontCollectionIndex = {
  MAJOR: "major",
  MINOR: "minor",
  NONE: "none",
} as const;

/**
 * Creates a dgm:style element (a:CT_ShapeStyle).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_ShapeStyle">
 *   <xsd:sequence>
 *     <xsd:element name="lnRef" type="CT_StyleMatrixReference"/>
 *     <xsd:element name="fillRef" type="CT_StyleMatrixReference"/>
 *     <xsd:element name="effectRef" type="CT_StyleMatrixReference"/>
 *     <xsd:element name="fontRef" type="CT_FontReference"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
const AUTHORING_WRITE_CONTEXT: WriteContext = {
  addRelationship: () => "",
  addMedia: () => "",
  addHyperlink: () => {},
};

export const createDiagramStyle = (options?: DefaultShapeStyleOptions): string =>
  stringifyShapeStyle(
    {
      lineReference: options?.lineReference ?? { index: 1, color: { value: "accent1" } },
      fillReference: options?.fillReference ?? { index: 1, color: { value: "accent1" } },
      effectReference: options?.effectReference ?? { index: 0, color: { value: "accent1" } },
      fontReference: options?.fontReference ?? { collection: "minor", color: { value: "tx1" } },
    },
    AUTHORING_WRITE_CONTEXT,
    "dgm:style",
  );

// ---------------------------------------------------------------------------
// Color list elements (CT_Colors) — used for fillClrLst, linClrLst, etc.
// ---------------------------------------------------------------------------

/** How the color list maps onto shapes (ST_ColorMethod): "span" interpolate across the list, "cycle" repeat in order, "repeat" alternate. */
export type ColorMethod = "span" | "cycle" | "repeat";

/** Hue rotation direction (ST_HueDir). */
export type HueDirection = "clockwise" | "counterClockwise";

export interface ColorListOptions {
  /** Color method (default: span) */
  meth?: ColorMethod;
  /** Hue direction (default: cw) */
  hueDir?: HueDirection;
  /** Colors (EG_ColorChoice items) */
  colors?: readonly SolidFillOptions[];
}

/** Creates a generic CT_Colors element with the given tag. */
export const createColorList = (tag: string, options?: ColorListOptions): string => {
  const children: string[] = [];
  if (options?.colors) {
    for (const c of options.colors) {
      children.push(createColorElement(c));
    }
  }

  const hasAttrs = options?.meth !== undefined || options?.hueDir !== undefined;
  const attrs: Record<string, string> = {};
  if (options?.meth) attrs.meth = options.meth;
  if (options?.hueDir) attrs.hueDir = xsdHueDirection.to(options.hueDir);

  return element(tag, hasAttrs ? attrs : undefined, children.length > 0 ? children : undefined);
};

/**
 * Parses a CT_Colors element (fillClrLst/linClrLst/…) back into options.
 * Colors parse in document order.
 */
export const parseColorList = (el: Element): ColorListOptions => {
  const result: ColorListOptions = {};
  if (el.attributes?.meth !== undefined) result.meth = el.attributes.meth as ColorMethod;
  if (el.attributes?.hueDir !== undefined)
    result.hueDir = xsdHueDirection.from(String(el.attributes.hueDir)) as HueDirection;
  const colors: SolidFillOptions[] = [];
  for (const child of el.elements ?? []) {
    const color = parseColorChoiceElement(child, NOOP_READ_CTX);
    if (color) colors.push(color);
  }
  if (colors.length) result.colors = colors;
  return result;
};

// Noop context: color elements never resolve relationships or parts.
const NOOP_READ_CTX = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
};

// ---------------------------------------------------------------------------
// dgm:fillClrLst / dgm:linClrLst / dgm:effectClrLst (used in CT_CTStyleLabel)
// ---------------------------------------------------------------------------

export const createFillColorList = (options?: ColorListOptions): string =>
  createColorList("dgm:fillClrLst", options);
export const createLineColorList = (options?: ColorListOptions): string =>
  createColorList("dgm:linClrLst", options);
export const createEffectColorList = (options?: ColorListOptions): string =>
  createColorList("dgm:effectClrLst", options);

// ---------------------------------------------------------------------------
// dgm:txFillClrLst / dgm:txLinClrLst / dgm:txEffectClrLst (text color lists)
// ---------------------------------------------------------------------------

export const createTextFillColorList = (options?: ColorListOptions): string =>
  createColorList("dgm:txFillClrLst", options);

export const createTextLineColorList = (options?: ColorListOptions): string =>
  createColorList("dgm:txLinClrLst", options);

export const createTextEffectColorList = (options?: ColorListOptions): string =>
  createColorList("dgm:txEffectClrLst", options);

// ---------------------------------------------------------------------------
// dgm:styleLbl — color-transform style label (CT_CTStyleLabel)
// ---------------------------------------------------------------------------

export interface ColorStyleLabelOptions {
  /** Label name referenced by dgm:layoutNode `@styleLbl` (required) */
  name: string;
  fillColorList?: ColorListOptions;
  lineColorList?: ColorListOptions;
  effectColorList?: ColorListOptions;
  textFillColorList?: ColorListOptions;
  textLineColorList?: ColorListOptions;
  textEffectColorList?: ColorListOptions;
}

/**
 * Creates a dgm:styleLbl element (CT_StyleLabel or CT_CTStyleLabel).
 *
 * ## XSD Schema (CT_CTStyleLabel, for color transforms)
 * ```xml
 * <xsd:complexType name="CT_CTStyleLabel">
 *   <xsd:sequence>
 *     <xsd:element name="fillClrLst" type="CT_Colors" minOccurs="0"/>
 *     <xsd:element name="linClrLst" type="CT_Colors" minOccurs="0"/>
 *     <xsd:element name="effectClrLst" type="CT_Colors" minOccurs="0"/>
 *     <xsd:element name="txLinClrLst" type="CT_Colors" minOccurs="0"/>
 *     <xsd:element name="txFillClrLst" type="CT_Colors" minOccurs="0"/>
 *     <xsd:element name="txEffectClrLst" type="CT_Colors" minOccurs="0"/>
 *     <xsd:element name="extLst" type="a:CT_OfficeArtExtensionList" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="name" type="xsd:string" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * ## XSD Schema (CT_StyleLabel, for quick styles)
 * ```xml
 * <xsd:complexType name="CT_StyleLabel">
 *   <xsd:sequence>
 *     <xsd:element name="scene3d" type="a:CT_Scene3D" minOccurs="0"/>
 *     <xsd:element name="sp3d" type="a:CT_Shape3D" minOccurs="0"/>
 *     <xsd:element name="txPr" type="CT_TextProps" minOccurs="0"/>
 *     <xsd:element name="style" type="a:CT_ShapeStyle" minOccurs="0"/>
 *     <xsd:element name="extLst" type="a:CT_OfficeArtExtensionList" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="name" type="xsd:string" use="required"/>
 * </xsd:complexType>
 * ```
 */
export const createStyleLabel = (options: ColorStyleLabelOptions): string => {
  const children: string[] = [];
  if (options.fillColorList) children.push(createFillColorList(options.fillColorList));
  if (options.lineColorList) children.push(createLineColorList(options.lineColorList));
  if (options.effectColorList) children.push(createEffectColorList(options.effectColorList));
  if (options.textFillColorList) children.push(createTextFillColorList(options.textFillColorList));
  if (options.textLineColorList) children.push(createTextLineColorList(options.textLineColorList));
  if (options.textEffectColorList)
    children.push(createTextEffectColorList(options.textEffectColorList));

  return element("dgm:styleLbl", { name: options.name }, children);
};
