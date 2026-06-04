/**
 * DrawingML ChartDrawing elements (cdr: namespace).
 *
 * These elements define the drawing layer within chart documents,
 * including shapes, connectors, pictures, group shapes, graphic frames,
 * and their anchoring (absolute or relative size).
 *
 * Reference: ISO/IEC 29500-4, dml-chartDrawing.xsd
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";
import type { XmlComponent, IXmlableObject } from "../../xml-components";

// ── Marker Coordinates ──

/**
 * Creates a cdr:x element (Marker X Coordinate).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="x" type="ST_MarkerCoordinate"/>
 * ```
 */
export const createCdrX = (value: number): XmlComponent =>
  new BuilderElement({
    children: [String(value)],
    name: "cdr:x",
  });

/**
 * Creates a cdr:y element (Marker Y Coordinate).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="y" type="ST_MarkerCoordinate"/>
 * ```
 */
export const createCdrY = (value: number): XmlComponent =>
  new BuilderElement({
    children: [String(value)],
    name: "cdr:y",
  });

/**
 * Creates a cdr:from element (Marker - top-left corner).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="from" type="CT_Marker"/>
 * ```
 */
export const createCdrFrom = (x: number, y: number): XmlComponent =>
  new BuilderElement({
    children: [createCdrX(x), createCdrY(y)],
    name: "cdr:from",
  });

/**
 * Creates a cdr:to element (Marker - bottom-right corner).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="to" type="CT_Marker"/>
 * ```
 */
export const createCdrTo = (x: number, y: number): XmlComponent =>
  new BuilderElement({
    children: [createCdrX(x), createCdrY(y)],
    name: "cdr:to",
  });

/**
 * Creates a cdr:ext element (Positive Size 2D).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="ext" type="a:CT_PositiveSize2D"/>
 * ```
 */
export const createCdrExt = (options: { readonly cx: number; readonly cy: number }): XmlComponent =>
  new BuilderElement({
    attributes: {
      cx: { key: "cx", value: options.cx },
      cy: { key: "cy", value: options.cy },
    },
    name: "cdr:ext",
  });

// ── Non-Visual Properties ──

/**
 * Creates a cdr:cNvPr element (Non-Visual Drawing Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvPr" type="a:CT_NonVisualDrawingProps"/>
 * ```
 */
export const createCdrCNvPr = (options: {
  readonly id?: number;
  readonly name?: string;
}): XmlComponent =>
  new BuilderElement({
    attributes: {
      ...(options.id !== undefined ? { id: { key: "id", value: options.id } } : {}),
      ...(options.name ? { name: { key: "name", value: options.name } } : {}),
    },
    name: "cdr:cNvPr",
  });

/**
 * Creates a cdr:cNvSpPr element (Non-Visual Drawing Shape Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvSpPr" type="a:CT_NonVisualDrawingShapeProps"/>
 * ```
 */
export const createCdrCNvSpPr = (): XmlComponent => new BuilderElement({ name: "cdr:cNvSpPr" });

/**
 * Creates a cdr:cNvCxnSpPr element (Non-Visual Connector Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvCxnSpPr" type="a:CT_NonVisualConnectorProperties"/>
 * ```
 */
export const createCdrCNvCxnSpPr = (): XmlComponent =>
  new BuilderElement({ name: "cdr:cNvCxnSpPr" });

/**
 * Creates a cdr:cNvPicPr element (Non-Visual Picture Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvPicPr" type="a:CT_NonVisualPictureProperties"/>
 * ```
 */
export const createCdrCNvPicPr = (): XmlComponent => new BuilderElement({ name: "cdr:cNvPicPr" });

/**
 * Creates a cdr:cNvGrpSpPr element (Non-Visual Group Shape Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvGrpSpPr" type="a:CT_NonVisualGroupDrawingShapeProps"/>
 * ```
 */
export const createCdrCNvGrpSpPr = (): XmlComponent =>
  new BuilderElement({ name: "cdr:cNvGrpSpPr" });

/**
 * Creates a cdr:cNvGraphicFramePr element (Non-Visual Graphic Frame Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvGraphicFramePr" type="a:CT_NonVisualGraphicFrameProperties"/>
 * ```
 */
export const createCdrCNvGraphicFramePr = (): XmlComponent =>
  new BuilderElement({ name: "cdr:cNvGraphicFramePr" });

// ── Non-Visual Wrappers ──

/**
 * Creates a cdr:nvSpPr element (Shape Non-Visual Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="nvSpPr" type="CT_ShapeNonVisual"/>
 * ```
 */
export const createCdrNvSpPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:nvSpPr" });

/**
 * Creates a cdr:nvCxnSpPr element (Connector Non-Visual Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="nvCxnSpPr" type="CT_ConnectorNonVisual"/>
 * ```
 */
export const createCdrNvCxnSpPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:nvCxnSpPr" });

/**
 * Creates a cdr:nvPicPr element (Picture Non-Visual Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="nvPicPr" type="CT_PictureNonVisual"/>
 * ```
 */
export const createCdrNvPicPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:nvPicPr" });

/**
 * Creates a cdr:nvGrpSpPr element (Group Shape Non-Visual Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="nvGrpSpPr" type="CT_GroupShapeNonVisual"/>
 * ```
 */
export const createCdrNvGrpSpPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:nvGrpSpPr" });

/**
 * Creates a cdr:nvGraphicFramePr element (Graphic Frame Non-Visual Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="nvGraphicFramePr" type="CT_GraphicFrameNonVisual"/>
 * ```
 */
export const createCdrNvGraphicFramePr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:nvGraphicFramePr" });

// ── Shape / Connector / Picture / Group / GraphicFrame ──

/**
 * Creates a cdr:sp element (Shape).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="sp" type="CT_Shape"/>
 * ```
 */
export const createCdrSp = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:sp" });

/**
 * Creates a cdr:cxnSp element (Connector Shape).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cxnSp" type="CT_Connector"/>
 * ```
 */
export const createCdrCxnSp = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:cxnSp" });

/**
 * Creates a cdr:pic element (Picture).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="pic" type="CT_Picture"/>
 * ```
 */
export const createCdrPic = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:pic" });

/**
 * Creates a cdr:grpSp element (Group Shape).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="grpSp" type="CT_GroupShape"/>
 * ```
 */
export const createCdrGrpSp = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:grpSp" });

/**
 * Creates a cdr:graphicFrame element (Graphic Frame).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="graphicFrame" type="CT_GraphicFrame"/>
 * ```
 */
export const createCdrGraphicFrame = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:graphicFrame" });

// ── Shape Properties (cdr: namespace) ──

/**
 * Creates a cdr:spPr element (Shape Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="spPr" type="a:CT_ShapeProperties"/>
 * ```
 */
export const createCdrSpPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:spPr" });

/**
 * Creates a cdr:grpSpPr element (Group Shape Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="grpSpPr" type="a:CT_GroupShapeProperties"/>
 * ```
 */
export const createCdrGrpSpPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:grpSpPr" });

/**
 * Creates a cdr:style element (Shape Style).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="style" type="a:CT_ShapeStyle"/>
 * ```
 */
export const createCdrStyle = (): XmlComponent => new BuilderElement({ name: "cdr:style" });

/**
 * Creates a cdr:txBody element (Text Body).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="txBody" type="a:CT_TextBody"/>
 * ```
 */
export const createCdrTxBody = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:txBody" });

/**
 * Creates a cdr:blipFill element (Blip Fill).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="blipFill" type="a:CT_BlipFillProperties"/>
 * ```
 */
export const createCdrBlipFill = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:blipFill" });

/**
 * Creates a cdr:xfrm element (2D Transform).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="xfrm" type="a:CT_Transform2D"/>
 * ```
 */
export const createCdrXfrm = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:xfrm" });

// ── Anchors ──

/**
 * Creates a cdr:relSizeAnchor element (Relative Size Anchor).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="relSizeAnchor" type="CT_RelSizeAnchor"/>
 * ```
 */
export const createCdrRelSizeAnchor = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:relSizeAnchor" });

/**
 * Creates a cdr:absSizeAnchor element (Absolute Size Anchor).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="absSizeAnchor" type="CT_AbsSizeAnchor"/>
 * ```
 */
export const createCdrAbsSizeAnchor = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "cdr:absSizeAnchor" });
