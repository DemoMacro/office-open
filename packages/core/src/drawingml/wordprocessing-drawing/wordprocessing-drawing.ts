/**
 * DrawingML WordprocessingDrawing elements (wp: namespace).
 *
 * These elements define inline and anchored drawings within WordprocessingML
 * documents, including wordprocessing shapes, groups, canvases, content parts,
 * and their non-visual properties.
 *
 * Reference: ISO/IEC 29500-4, dml-wordprocessingDrawing.xsd
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";
import type { XmlComponent, IXmlableObject } from "../../xml-components";

// ── Non-Visual Properties ──

export interface WpNonVisualDrawingPropsOptions {
  readonly id?: number;
  readonly name?: string;
  readonly descr?: string;
  readonly hidden?: boolean;
}

/**
 * Creates a wp:cNvPr element (Non-Visual Drawing Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvPr" type="a:CT_NonVisualDrawingProps"/>
 * ```
 */
export const createWpCNvPr = (options?: WpNonVisualDrawingPropsOptions): XmlComponent =>
  new BuilderElement({
    attributes: {
      ...(options?.id !== undefined ? { id: { key: "id", value: options.id } } : {}),
      ...(options?.name ? { name: { key: "name", value: options.name } } : {}),
      ...(options?.descr ? { descr: { key: "descr", value: options.descr } } : {}),
      ...(options?.hidden !== undefined
        ? { hidden: { key: "hidden", value: options.hidden ? 1 : 0 } }
        : {}),
    },
    name: "wp:cNvPr",
  });

/**
 * Creates a wp:cNvSpPr element (Non-Visual Drawing Shape Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvSpPr" type="a:CT_NonVisualDrawingShapeProps"/>
 * ```
 */
export const createWpCNvSpPr = (options?: { readonly txSp?: boolean }): XmlComponent =>
  new BuilderElement({
    attributes:
      options?.txSp !== undefined
        ? { txSp: { key: "txSp", value: options.txSp ? 1 : 0 } }
        : undefined,
    name: "wp:cNvSpPr",
  });

/**
 * Creates a wp:cNvCnPr element (Non-Visual Connector Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvCnPr" type="a:CT_NonVisualConnectorProperties"/>
 * ```
 */
export const createWpCNvCnPr = (): XmlComponent => new BuilderElement({ name: "wp:cNvCnPr" });

/**
 * Creates a wp:cNvFrPr element (Non-Visual Graphic Frame Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvFrPr" type="a:CT_NonVisualGraphicFrameProperties"/>
 * ```
 */
export const createWpCNvFrPr = (): XmlComponent => new BuilderElement({ name: "wp:cNvFrPr" });

/**
 * Creates a wp:cNvGrpSpPr element (Non-Visual Group Shape Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvGrpSpPr" type="a:CT_NonVisualGroupDrawingShapeProps"/>
 * ```
 */
export const createWpCNvGrpSpPr = (): XmlComponent => new BuilderElement({ name: "wp:cNvGrpSpPr" });

/**
 * Creates a wp:cNvContentPartPr element (Non-Visual Content Part Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvContentPartPr" type="a:CT_NonVisualContentPartProperties"/>
 * ```
 */
export const createWpCNvContentPartPr = (): XmlComponent =>
  new BuilderElement({ name: "wp:cNvContentPartPr" });

// ── Wordprocessing Shape ──

export interface WpShapeOptions {
  readonly cNvPr?: WpNonVisualDrawingPropsOptions;
  readonly normalEastAsianFlow?: boolean;
  readonly children?: readonly IXmlableObject[];
}

/**
 * Creates a wp:wsp element (Wordprocessing Shape).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="wsp" type="CT_WordprocessingShape"/>
 * ```
 */
export const createWpSp = (options?: WpShapeOptions): XmlComponent =>
  new BuilderElement({
    attributes:
      options?.normalEastAsianFlow !== undefined
        ? {
            normalEastAsianFlow: {
              key: "normalEastAsianFlow",
              value: options.normalEastAsianFlow ? 1 : 0,
            },
          }
        : undefined,
    children: options?.children,
    name: "wp:wsp",
  });

// ── Shape Properties (local wp: namespace) ──

/**
 * Creates a wp:spPr element (Shape Properties within wp namespace).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="spPr" type="a:CT_ShapeProperties"/>
 * ```
 */
export const createWpSpPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "wp:spPr" });

/**
 * Creates a wp:grpSpPr element (Group Shape Properties within wp namespace).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="grpSpPr" type="a:CT_GroupShapeProperties"/>
 * ```
 */
export const createWpGrpSpPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "wp:grpSpPr" });

/**
 * Creates a wp:style element (Shape Style).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="style" type="a:CT_ShapeStyle"/>
 * ```
 */
export const createWpStyle = (): XmlComponent => new BuilderElement({ name: "wp:style" });

/**
 * Creates a wp:xfrm element (2D Transform).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="xfrm" type="a:CT_Transform2D"/>
 * ```
 */
export const createWpXfrm = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "wp:xfrm" });

/**
 * Creates a wp:bodyPr element (Text Body Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="bodyPr" type="a:CT_TextBodyProperties"/>
 * ```
 */
export const createWpBodyPr = (): XmlComponent => new BuilderElement({ name: "wp:bodyPr" });

// ── Textbox ──

export interface WpTextboxOptions {
  readonly id?: number;
  readonly children?: readonly IXmlableObject[];
}

/**
 * Creates a wp:txbx element (Textbox Info).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="txbx" type="CT_TextboxInfo"/>
 * ```
 */
export const createWpTxbx = (options?: WpTextboxOptions): XmlComponent =>
  new BuilderElement({
    attributes: options?.id !== undefined ? { id: { key: "id", value: options.id } } : undefined,
    children: options?.children,
    name: "wp:txbx",
  });

/**
 * Creates a wp:txbxContent element (Textbox Content).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="txbxContent" type="CT_TxbxContent"/>
 * ```
 */
export const createWpTxbxContent = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "wp:txbxContent" });

export interface WpLinkedTextboxOptions {
  readonly id: number;
  readonly seq: number;
}

/**
 * Creates a wp:linkedTxbx element (Linked Textbox Information).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="linkedTxbx" type="CT_LinkedTextboxInformation"/>
 * ```
 */
export const createWpLinkedTxbx = (options: WpLinkedTextboxOptions): XmlComponent =>
  new BuilderElement({
    attributes: {
      id: { key: "id", value: options.id },
      seq: { key: "seq", value: options.seq },
    },
    name: "wp:linkedTxbx",
  });

// ── Extension List ──

/**
 * Creates a wp:extLst element (Extension List).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="extLst" type="a:CT_OfficeArtExtensionList"/>
 * ```
 */
export const createWpExtLst = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "wp:extLst" });

// ── Graphic Frame ──

export interface WpGraphicFrameOptions {
  readonly children?: readonly IXmlableObject[];
}

/**
 * Creates a wp:graphicFrame element (Graphic Frame).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="graphicFrame" type="CT_GraphicFrame"/>
 * ```
 */
export const createWpGraphicFrame = (options?: WpGraphicFrameOptions): XmlComponent =>
  new BuilderElement({ children: options?.children, name: "wp:graphicFrame" });

// ── Content Part ──

export interface WpContentPartOptions {
  readonly bwMode?: string;
  readonly rId?: string;
}

/**
 * Creates a wp:contentPart element (Content Part).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="contentPart" type="CT_WordprocessingContentPart"/>
 * ```
 */
export const createWpContentPart = (options?: WpContentPartOptions): XmlComponent =>
  new BuilderElement({
    attributes: {
      ...(options?.bwMode ? { bwMode: { key: "bwMode", value: options.bwMode } } : {}),
      ...(options?.rId ? { "r:id": { key: "r:id", value: options.rId } } : {}),
    },
    name: "wp:contentPart",
  });

// ── Non-Visual Content Part ──

/**
 * Creates a wp:nvContentPartPr element (Non-Visual Content Part Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="nvContentPartPr" type="CT_WordprocessingContentPartNonVisual"/>
 * ```
 */
export const createWpNvContentPartPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "wp:nvContentPartPr" });

// ── Group Shape ──

export interface WpGroupOptions {
  readonly children?: readonly IXmlableObject[];
}

/**
 * Creates a wp:wgp element (Wordprocessing Group).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="wgp" type="CT_WordprocessingGroup"/>
 * ```
 */
export const createWpGrpSp = (options?: WpGroupOptions): XmlComponent =>
  new BuilderElement({ children: options?.children, name: "wp:wgp" });

/**
 * Creates a wp:grpSp element (Nested Group Shape within Wordprocessing Group).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="grpSp" type="CT_WordprocessingGroup"/>
 * ```
 */
export const createWpNestedGrpSp = (options?: WpGroupOptions): XmlComponent =>
  new BuilderElement({ children: options?.children, name: "wp:grpSp" });

// ── Canvas ──

export interface WpCanvasOptions {
  readonly children?: readonly IXmlableObject[];
}

/**
 * Creates a wp:wpc element (Wordprocessing Canvas).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="wpc" type="CT_WordprocessingCanvas"/>
 * ```
 */
export const createWpCanvas = (options?: WpCanvasOptions): XmlComponent =>
  new BuilderElement({ children: options?.children, name: "wp:wpc" });

// ── Canvas formatting ──

/**
 * Creates a wp:bg element (Background Formatting).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="bg" type="a:CT_BackgroundFormatting"/>
 * ```
 */
export const createWpBg = (): XmlComponent => new BuilderElement({ name: "wp:bg" });

/**
 * Creates a wp:whole element (Whole E2O Formatting).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="whole" type="a:CT_WholeE2oFormatting"/>
 * ```
 */
export const createWpWhole = (): XmlComponent => new BuilderElement({ name: "wp:whole" });
