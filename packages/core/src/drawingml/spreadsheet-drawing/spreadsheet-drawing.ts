/**
 * DrawingML SpreadsheetDrawing elements (xdr: namespace).
 *
 * These elements define anchoring of images, charts, shapes, and other
 * graphical objects to worksheet cells in SpreadsheetML documents.
 *
 * Reference: ISO/IEC 29500-4, dml-spreadsheetDrawing.xsd
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";
import type { XmlComponent, IXmlableObject } from "../../xml-components";

// ── Cell Marker ──

export interface XdrMarkerOptions {
  col: number;
  colOff: number;
  row: number;
  rowOff: number;
}

/**
 * Creates an xdr:col element (Column ID).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="col" type="ST_ColID"/>
 * ```
 */
export const createXdrCol = (value: number): XmlComponent =>
  new BuilderElement({
    children: [String(value)],
    name: "xdr:col",
  });

/**
 * Creates an xdr:colOff element (Column Offset).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="colOff" type="a:ST_Coordinate"/>
 * ```
 */
export const createXdrColOff = (value: number): XmlComponent =>
  new BuilderElement({
    children: [String(value)],
    name: "xdr:colOff",
  });

/**
 * Creates an xdr:row element (Row ID).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="row" type="ST_RowID"/>
 * ```
 */
export const createXdrRow = (value: number): XmlComponent =>
  new BuilderElement({
    children: [String(value)],
    name: "xdr:row",
  });

/**
 * Creates an xdr:rowOff element (Row Offset).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="rowOff" type="a:ST_Coordinate"/>
 * ```
 */
export const createXdrRowOff = (value: number): XmlComponent =>
  new BuilderElement({
    children: [String(value)],
    name: "xdr:rowOff",
  });

/**
 * Creates an xdr:from element (Marker - top-left corner).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="from" type="CT_Marker"/>
 * ```
 */
export const createXdrFrom = (options: XdrMarkerOptions): XmlComponent =>
  new BuilderElement({
    children: [
      createXdrCol(options.col),
      createXdrColOff(options.colOff),
      createXdrRow(options.row),
      createXdrRowOff(options.rowOff),
    ],
    name: "xdr:from",
  });

/**
 * Creates an xdr:to element (Marker - bottom-right corner).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="to" type="CT_Marker"/>
 * ```
 */
export const createXdrTo = (options: XdrMarkerOptions): XmlComponent =>
  new BuilderElement({
    children: [
      createXdrCol(options.col),
      createXdrColOff(options.colOff),
      createXdrRow(options.row),
      createXdrRowOff(options.rowOff),
    ],
    name: "xdr:to",
  });

// ── Non-Visual Properties ──

/**
 * Creates an xdr:cNvPr element (Non-Visual Drawing Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvPr" type="a:CT_NonVisualDrawingProps"/>
 * ```
 */
export const createXdrCNvPr = (options: { id?: number; name?: string }): XmlComponent =>
  new BuilderElement({
    attributes: {
      ...(options.id !== undefined ? { id: { key: "id", value: options.id } } : {}),
      ...(options.name ? { name: { key: "name", value: options.name } } : {}),
    },
    name: "xdr:cNvPr",
  });

/**
 * Creates an xdr:cNvSpPr element (Non-Visual Drawing Shape Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvSpPr" type="a:CT_NonVisualDrawingShapeProps"/>
 * ```
 */
export const createXdrCNvSpPr = (): XmlComponent => new BuilderElement({ name: "xdr:cNvSpPr" });

/**
 * Creates an xdr:cNvCxnSpPr element (Non-Visual Connector Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvCxnSpPr" type="a:CT_NonVisualConnectorProperties"/>
 * ```
 */
export const createXdrCNvCxnSpPr = (): XmlComponent =>
  new BuilderElement({ name: "xdr:cNvCxnSpPr" });

/**
 * Creates an xdr:cNvPicPr element (Non-Visual Picture Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvPicPr" type="a:CT_NonVisualPictureProperties"/>
 * ```
 */
export const createXdrCNvPicPr = (options?: { preferRelativeResize?: boolean }): XmlComponent =>
  new BuilderElement({
    attributes:
      options?.preferRelativeResize !== undefined
        ? {
            preferRelativeResize: {
              key: "preferRelativeResize",
              value: options.preferRelativeResize ? 1 : 0,
            },
          }
        : undefined,
    name: "xdr:cNvPicPr",
  });

/**
 * Creates an xdr:cNvGrpSpPr element (Non-Visual Group Shape Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cNvGrpSpPr" type="a:CT_NonVisualGroupDrawingShapeProps"/>
 * ```
 */
export const createXdrCNvGrpSpPr = (): XmlComponent =>
  new BuilderElement({ name: "xdr:cNvGrpSpPr" });

// ── Non-Visual Wrappers ──

/**
 * Creates an xdr:nvSpPr element (Shape Non-Visual Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="nvSpPr" type="CT_ShapeNonVisual"/>
 * ```
 */
export const createXdrNvSpPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "xdr:nvSpPr" });

/**
 * Creates an xdr:nvCxnSpPr element (Connector Non-Visual Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="nvCxnSpPr" type="CT_ConnectorNonVisual"/>
 * ```
 */
export const createXdrNvCxnSpPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "xdr:nvCxnSpPr" });

/**
 * Creates an xdr:nvPicPr element (Picture Non-Visual Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="nvPicPr" type="CT_PictureNonVisual"/>
 * ```
 */
export const createXdrNvPicPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "xdr:nvPicPr" });

/**
 * Creates an xdr:nvGrpSpPr element (Group Shape Non-Visual Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="nvGrpSpPr" type="CT_GroupShapeNonVisual"/>
 * ```
 */
export const createXdrNvGrpSpPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "xdr:nvGrpSpPr" });

// ── Shape / Connector / Picture / Group ──

export interface XdrShapeAttributes {
  macro?: string;
  textlink?: string;
  fLocksText?: boolean;
  fPublished?: boolean;
}

/**
 * Creates an xdr:sp element (Shape).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="sp" type="CT_Shape"/>
 * ```
 */
export const createXdrSp = (options?: {
  attributes?: XdrShapeAttributes;
  children?: readonly IXmlableObject[];
}): XmlComponent =>
  new BuilderElement({
    attributes: options?.attributes
      ? {
          ...(options.attributes.macro
            ? { macro: { key: "macro", value: options.attributes.macro } }
            : {}),
          ...(options.attributes.textlink
            ? { textlink: { key: "textlink", value: options.attributes.textlink } }
            : {}),
          ...(options.attributes.fLocksText !== undefined
            ? {
                fLocksText: {
                  key: "fLocksText",
                  value: options.attributes.fLocksText ? 1 : 0,
                },
              }
            : {}),
          ...(options.attributes.fPublished !== undefined
            ? {
                fPublished: {
                  key: "fPublished",
                  value: options.attributes.fPublished ? 1 : 0,
                },
              }
            : {}),
        }
      : undefined,
    children: options?.children,
    name: "xdr:sp",
  });

/**
 * Creates an xdr:cxnSp element (Connector Shape).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="cxnSp" type="CT_Connector"/>
 * ```
 */
export const createXdrCxnSp = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "xdr:cxnSp" });

/**
 * Creates an xdr:pic element (Picture).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="pic" type="CT_Picture"/>
 * ```
 */
export const createXdrPic = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "xdr:pic" });

/**
 * Creates an xdr:grpSp element (Group Shape).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="grpSp" type="CT_GroupShape"/>
 * ```
 */
export const createXdrGrpSp = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "xdr:grpSp" });

// ── Shape Properties (xdr: namespace) ──

/**
 * Creates an xdr:spPr element (Shape Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="spPr" type="a:CT_ShapeProperties"/>
 * ```
 */
export const createXdrSpPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "xdr:spPr" });

/**
 * Creates an xdr:grpSpPr element (Group Shape Properties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="grpSpPr" type="a:CT_GroupShapeProperties"/>
 * ```
 */
export const createXdrGrpSpPr = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "xdr:grpSpPr" });

/**
 * Creates an xdr:style element (Shape Style).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="style" type="a:CT_ShapeStyle"/>
 * ```
 */
export const createXdrStyle = (): XmlComponent => new BuilderElement({ name: "xdr:style" });

/**
 * Creates an xdr:txBody element (Text Body).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="txBody" type="a:CT_TextBody"/>
 * ```
 */
export const createXdrTxBody = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "xdr:txBody" });

/**
 * Creates an xdr:blipFill element (Blip Fill).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="blipFill" type="a:CT_BlipFillProperties"/>
 * ```
 */
export const createXdrBlipFill = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "xdr:blipFill" });

/**
 * Creates an xdr:contentPart element (Content Part).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="contentPart" type="CT_Rel"/>
 * ```
 */
export const createXdrContentPart = (rId?: string): XmlComponent =>
  new BuilderElement({
    attributes: rId ? { "r:id": { key: "r:id", value: rId } } : undefined,
    name: "xdr:contentPart",
  });

// ── Anchors ──

export interface XdrTwoCellAnchorOptions {
  editAs?: "twoCell" | "oneCell" | "absolute";
  children?: readonly IXmlableObject[];
}

/**
 * Creates an xdr:twoCellAnchor element (Two Cell Anchor).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="twoCellAnchor" type="CT_TwoCellAnchor"/>
 * ```
 */
export const createXdrTwoCellAnchor = (options?: XdrTwoCellAnchorOptions): XmlComponent =>
  new BuilderElement({
    attributes: options?.editAs ? { editAs: { key: "editAs", value: options.editAs } } : undefined,
    children: options?.children,
    name: "xdr:twoCellAnchor",
  });

/**
 * Creates an xdr:oneCellAnchor element (One Cell Anchor).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="oneCellAnchor" type="CT_OneCellAnchor"/>
 * ```
 */
export const createXdrOneCellAnchor = (children?: readonly IXmlableObject[]): XmlComponent =>
  new BuilderElement({ children, name: "xdr:oneCellAnchor" });
