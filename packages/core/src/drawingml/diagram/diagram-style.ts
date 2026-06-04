/**
 * Diagram style, color list, and style label elements.
 *
 * Reference: ISO/IEC 29500-4, dml-diagram.xsd
 * CT_Style, CT_StyleLabel, CT_Colors, CT_CTStyleLabel
 *
 * @module
 */
import { BuilderElement, type XmlComponent } from "../../xml-components";
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

export interface DiagramStyleOptions {
  /** Line style matrix reference index */
  readonly lnIdx?: number;
  /** Fill style matrix reference index */
  readonly fillIdx?: number;
  /** Effect style matrix reference index */
  readonly effectIdx?: number;
  /** Font reference collection index */
  readonly fontIdx?: string;
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
export const createDiagramStyle = (options?: DiagramStyleOptions): XmlComponent => {
  const children: XmlComponent[] = [];

  children.push(
    new BuilderElement({
      name: "a:lnRef",
      attributes: { idx: { key: "idx", value: options?.lnIdx ?? 1 } },
      children: [createColorElement({ value: SchemeColor.ACCENT1 })],
    }),
  );
  children.push(
    new BuilderElement({
      name: "a:fillRef",
      attributes: { idx: { key: "idx", value: options?.fillIdx ?? 1 } },
      children: [createColorElement({ value: SchemeColor.ACCENT1 })],
    }),
  );
  children.push(
    new BuilderElement({
      name: "a:effectRef",
      attributes: { idx: { key: "idx", value: options?.effectIdx ?? 0 } },
      children: [createColorElement({ value: SchemeColor.ACCENT1 })],
    }),
  );
  children.push(
    new BuilderElement({
      name: "a:fontRef",
      attributes: { idx: { key: "idx", value: options?.fontIdx ?? "minor" } },
      children: [createColorElement({ value: SchemeColor.TX1 })],
    }),
  );

  return new BuilderElement({ name: "dgm:style", children });
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
  readonly meth?: (typeof ColorMethod)[keyof typeof ColorMethod];
  /** Hue direction (default: cw) */
  readonly hueDir?: (typeof HueDirection)[keyof typeof HueDirection];
  /** Colors (EG_ColorChoice items) */
  readonly colors?: readonly SolidFillOptions[];
}

/** Creates a generic CT_Colors element with the given tag. */
const createColorList = (tag: string, options?: ColorListOptions): XmlComponent => {
  const children: XmlComponent[] = [];
  if (options?.colors) {
    for (const c of options.colors) {
      children.push(createColorElement(c));
    }
  }

  const hasAttrs = options?.meth !== undefined || options?.hueDir !== undefined;

  return new BuilderElement({
    name: tag,
    attributes: hasAttrs
      ? {
          ...(options!.meth && { meth: { key: "meth", value: options!.meth } }),
          ...(options!.hueDir && { hueDir: { key: "hueDir", value: options!.hueDir } }),
        }
      : undefined,
    children,
  });
};

// ---------------------------------------------------------------------------
// dgm:fillClrLst / dgm:linClrLst / dgm:effectClrLst (used in CT_CTStyleLabel)
// ---------------------------------------------------------------------------

export const createFillClrLst = (options?: ColorListOptions): XmlComponent =>
  createColorList("dgm:fillClrLst", options);
export const createLinClrLst = (options?: ColorListOptions): XmlComponent =>
  createColorList("dgm:linClrLst", options);
export const createEffectClrLst = (options?: ColorListOptions): XmlComponent =>
  createColorList("dgm:effectClrLst", options);

// ---------------------------------------------------------------------------
// dgm:txFillClrLst / dgm:txLinClrLst / dgm:txEffectClrLst (text color lists)
// ---------------------------------------------------------------------------

export const createTxFillClrLst = (options?: ColorListOptions): XmlComponent =>
  createColorList("dgm:txFillClrLst", options);

export const createTxLinClrLst = (options?: ColorListOptions): XmlComponent =>
  createColorList("dgm:txLinClrLst", options);

export const createTxEffectClrLst = (options?: ColorListOptions): XmlComponent =>
  createColorList("dgm:txEffectClrLst", options);

// ---------------------------------------------------------------------------
// dgm:styleLbl — style label (CT_StyleLabel)
// ---------------------------------------------------------------------------

export interface DiagramStyleLblOptions {
  /** Label name (required) */
  readonly name: string;
  readonly fillClrLst?: ColorListOptions;
  readonly linClrLst?: ColorListOptions;
  readonly effectClrLst?: ColorListOptions;
  readonly txFillClrLst?: ColorListOptions;
  readonly txLinClrLst?: ColorListOptions;
  readonly txEffectClrLst?: ColorListOptions;
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
export const createStyleLbl = (options: DiagramStyleLblOptions): XmlComponent => {
  const children: XmlComponent[] = [];
  if (options.fillClrLst) children.push(createFillClrLst(options.fillClrLst));
  if (options.linClrLst) children.push(createLinClrLst(options.linClrLst));
  if (options.effectClrLst) children.push(createEffectClrLst(options.effectClrLst));
  if (options.txFillClrLst) children.push(createTxFillClrLst(options.txFillClrLst));
  if (options.txLinClrLst) children.push(createTxLinClrLst(options.txLinClrLst));
  if (options.txEffectClrLst) children.push(createTxEffectClrLst(options.txEffectClrLst));

  return new BuilderElement({
    name: "dgm:styleLbl",
    attributes: { name: { key: "name", value: options.name } },
    children,
  });
};
