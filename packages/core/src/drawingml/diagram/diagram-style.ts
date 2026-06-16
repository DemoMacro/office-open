/**
 * Diagram style, color list, and style label elements.
 *
 * Reference: ISO/IEC 29500-4, dml-diagram.xsd
 * CT_Style, CT_StyleLabel, CT_Colors, CT_CTStyleLabel
 *
 * @module
 */
import { element } from "@office-open/xml";

import { SchemeColor } from "../color/scheme-color";
import type { SolidFillOptions } from "../color/solid-fill";
import { createColorElement } from "../color/solid-fill";

// ---------------------------------------------------------------------------
// dgm:style — diagram style (CT_ShapeStyle from dml-main.xsd)
// ---------------------------------------------------------------------------

export const StyleMatrixIndex = {
  SUBTLE: "subtle",
  MODERATE: "moderate",
  INTENSE: "intense",
} as const;

export const FontCollectionIndex = {
  MAJOR: "major",
  MINOR: "minor",
  NONE: "none",
} as const;

/** Reference into the theme style matrix (a:lnRef/fillRef/effectRef). */
export interface DiagramStyleReferenceOptions {
  /** Index into the theme style matrix (a:*Ref @idx) */
  idx: number;
}

/** Font reference (a:fontRef) — idx is a FontCollectionIndex (major/minor/none). */
export interface DiagramFontReferenceOptions {
  /** Font collection index (FontCollectionIndex.MAJOR/MINOR/NONE) */
  idx: string;
}

export interface DiagramStyleOptions {
  /** Line reference (a:lnRef) */
  lineReference?: DiagramStyleReferenceOptions;
  /** Fill reference (a:fillRef) */
  fillReference?: DiagramStyleReferenceOptions;
  /** Effect reference (a:effectRef) */
  effectReference?: DiagramStyleReferenceOptions;
  /** Font reference (a:fontRef) */
  fontReference?: DiagramFontReferenceOptions;
}

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
export const createDiagramStyle = (options?: DiagramStyleOptions): string => {
  const children: string[] = [];

  children.push(
    element("a:lnRef", { idx: options?.lineReference?.idx ?? 1 }, [
      createColorElement({ value: SchemeColor.ACCENT1 }),
    ]),
  );
  children.push(
    element("a:fillRef", { idx: options?.fillReference?.idx ?? 1 }, [
      createColorElement({ value: SchemeColor.ACCENT1 }),
    ]),
  );
  children.push(
    element("a:effectRef", { idx: options?.effectReference?.idx ?? 0 }, [
      createColorElement({ value: SchemeColor.ACCENT1 }),
    ]),
  );
  children.push(
    element("a:fontRef", { idx: options?.fontReference?.idx ?? "minor" }, [
      createColorElement({ value: SchemeColor.TX1 }),
    ]),
  );

  return element("dgm:style", undefined, children);
};

// ---------------------------------------------------------------------------
// Color list elements (CT_Colors) — used for fillClrLst, linClrLst, etc.
// ---------------------------------------------------------------------------

export const ColorMethod = {
  SPAN: "span",
  CYCLE: "cycle",
  REPEAT: "repeat",
} as const;

export const HueDirection = {
  CLOCKWISE: "cw",
  COUNTER_CLOCKWISE: "ccw",
} as const;

export interface ColorListOptions {
  /** Color method (default: span) */
  meth?: (typeof ColorMethod)[keyof typeof ColorMethod];
  /** Hue direction (default: cw) */
  hueDir?: (typeof HueDirection)[keyof typeof HueDirection];
  /** Colors (EG_ColorChoice items) */
  colors?: readonly SolidFillOptions[];
}

/** Creates a generic CT_Colors element with the given tag. */
const createColorList = (tag: string, options?: ColorListOptions): string => {
  const children: string[] = [];
  if (options?.colors) {
    for (const c of options.colors) {
      children.push(createColorElement(c));
    }
  }

  const hasAttrs = options?.meth !== undefined || options?.hueDir !== undefined;
  const attrs: Record<string, string> = {};
  if (options?.meth) attrs.meth = options.meth;
  if (options?.hueDir) attrs.hueDir = options.hueDir;

  return element(tag, hasAttrs ? attrs : undefined, children.length > 0 ? children : undefined);
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
// dgm:styleLbl — style label (CT_StyleLabel)
// ---------------------------------------------------------------------------

export interface DiagramStyleLabelOptions {
  /** Label name (required) */
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
export const createStyleLabel = (options: DiagramStyleLabelOptions): string => {
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
