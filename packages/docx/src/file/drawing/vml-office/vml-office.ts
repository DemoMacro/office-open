/**
 * VML Office Drawing elements (o: namespace).
 *
 * Factory functions for Office-specific VML elements including shape layout,
 * defaults, locks, extrusion, callout, skew, diagram, and OLE helper elements.
 *
 * Reference: ISO/IEC 29500-4, vml-officeDrawing.xsd
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

// ── Shape Layout & Defaults ──

/**
 * Creates an o:shapelayout element (Shape Layout).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="shapelayout" type="CT_ShapeLayout"/>
 * ```
 */
export const createOShapelayout = (): XmlComponent => new BuilderElement({ name: "o:shapelayout" });

/**
 * Creates an o:idmap element (ID Map).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="idmap" type="CT_IdMap"/>
 * ```
 */
export const createOIdmap = (): XmlComponent => new BuilderElement({ name: "o:idmap" });

/**
 * Creates an o:shapedefaults element (Shape Defaults).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="shapedefaults" type="CT_ShapeDefaults"/>
 * ```
 */
export const createOShapedefaults = (): XmlComponent =>
  new BuilderElement({ name: "o:shapedefaults" });

// ── Color ──

/**
 * Creates an o:colormru element (Color Most Recently Used).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="colormru" type="CT_ColorMru"/>
 * ```
 */
export const createOColormru = (): XmlComponent => new BuilderElement({ name: "o:colormru" });

/**
 * Creates an o:colormenu element (Color Menu).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="colormenu" type="CT_ColorMenu"/>
 * ```
 */
export const createOColormenu = (): XmlComponent => new BuilderElement({ name: "o:colormenu" });

// ── Lock ──

/**
 * Creates an o:lock element (Shape Lock).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="lock" type="CT_Lock"/>
 * ```
 */
export const createOLock = (): XmlComponent => new BuilderElement({ name: "o:lock" });

// ── 3D Effects ──

/**
 * Creates an o:extrusion element (Extrusion).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="extrusion" type="CT_Extrusion"/>
 * ```
 */
export const createOExtrusion = (): XmlComponent => new BuilderElement({ name: "o:extrusion" });

// ── Callout ──

/**
 * Creates an o:callout element (Callout).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="callout" type="CT_Callout"/>
 * ```
 */
export const createOCallout = (): XmlComponent => new BuilderElement({ name: "o:callout" });

// ── Skew ──

/**
 * Creates an o:skew element (Skew Transform).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="skew" type="CT_Skew"/>
 * ```
 */
export const createOSkew = (): XmlComponent => new BuilderElement({ name: "o:skew" });

// ── Diagram ──

/**
 * Creates an o:diagram element (Diagram).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="diagram" type="CT_Diagram"/>
 * ```
 */
export const createODiagram = (): XmlComponent => new BuilderElement({ name: "o:diagram" });

// ── Ink ──

/**
 * Creates an o:ink element (Ink).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="ink" type="CT_Ink"/>
 * ```
 */
export const createOInk = (): XmlComponent => new BuilderElement({ name: "o:ink" });

// ── Equation XML ──

/**
 * Creates an o:equationxml element (Equation XML).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="equationxml" type="CT_EquationXml"/>
 * ```
 */
export const createOEquationxml = (): XmlComponent => new BuilderElement({ name: "o:equationxml" });

// ── Signature Line ──

/**
 * Creates an o:signatureline element (Signature Line).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="signatureline" type="CT_SignatureLine"/>
 * ```
 */
export const createOSignatureline = (): XmlComponent =>
  new BuilderElement({ name: "o:signatureline" });

// ── Complex ──

/**
 * Creates an o:complex element (Complex).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="complex" type="CT_Complex"/>
 * ```
 */
export const createOComplex = (): XmlComponent => new BuilderElement({ name: "o:complex" });

// ── Fill ──

/**
 * Creates an o:fill element (Office Fill).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="fill" type="CT_Fill"/>
 * ```
 */
export const createOFill = (): XmlComponent => new BuilderElement({ name: "o:fill" });

// ── Clip Path ──

/**
 * Creates an o:clippath element (Clip Path).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="clippath" type="CT_ClipPath"/>
 * ```
 */
export const createOClippath = (): XmlComponent => new BuilderElement({ name: "o:clippath" });

// ── Stroke Children (left, top, right, bottom, column) ──

/**
 * Creates an o:left element (Left Stroke).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="left" type="CT_StrokeChild"/>
 * ```
 */
export const createOLeft = (): XmlComponent => new BuilderElement({ name: "o:left" });

/**
 * Creates an o:top element (Top Stroke).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="top" type="CT_StrokeChild"/>
 * ```
 */
export const createOTop = (): XmlComponent => new BuilderElement({ name: "o:top" });

/**
 * Creates an o:right element (Right Stroke).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="right" type="CT_StrokeChild"/>
 * ```
 */
export const createORight = (): XmlComponent => new BuilderElement({ name: "o:right" });

/**
 * Creates an o:bottom element (Bottom Stroke).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="bottom" type="CT_StrokeChild"/>
 * ```
 */
export const createOBottom = (): XmlComponent => new BuilderElement({ name: "o:bottom" });

/**
 * Creates an o:column element (Column Stroke).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="column" type="CT_StrokeChild"/>
 * ```
 */
export const createOColumn = (): XmlComponent => new BuilderElement({ name: "o:column" });

// ── Regroup Table ──

/**
 * Creates an o:regrouptable element (Regroup Table).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="regrouptable" type="CT_RegroupTable"/>
 * ```
 */
export const createORegrouptable = (): XmlComponent =>
  new BuilderElement({ name: "o:regrouptable" });

/**
 * Creates an o:entry element (Regroup Table Entry).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="entry" type="CT_Entry"/>
 * ```
 */
export const createOEntry = (): XmlComponent => new BuilderElement({ name: "o:entry" });

// ── Rules ──

/**
 * Creates an o:rules element (Rules).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="rules" type="CT_Rules"/>
 * ```
 */
export const createORules = (): XmlComponent => new BuilderElement({ name: "o:rules" });

/**
 * Creates an o:r element (Rule).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="r" type="CT_R"/>
 * ```
 */
export const createOR = (): XmlComponent => new BuilderElement({ name: "o:r" });

/**
 * Creates an o:proxy element (Rule Proxy).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="proxy" type="CT_Proxy"/>
 * ```
 */
export const createOProxy = (): XmlComponent => new BuilderElement({ name: "o:proxy" });

// ── Relation Table ──

/**
 * Creates an o:relationtable element (Relation Table).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="relationtable" type="CT_RelationTable"/>
 * ```
 */
export const createORelationtable = (): XmlComponent =>
  new BuilderElement({ name: "o:relationtable" });

/**
 * Creates an o:rel element (Relation).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="rel" type="CT_Relation"/>
 * ```
 */
export const createORel = (): XmlComponent => new BuilderElement({ name: "o:rel" });

// ── OLE Object Children ──

/**
 * Creates an o:LinkType element (OLE Link Type).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="LinkType" type="ST_OLELinkType"/>
 * ```
 */
export const createOLinkType = (): XmlComponent => new BuilderElement({ name: "o:LinkType" });

/**
 * Creates an o:LockedField element (OLE Locked Field).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="LockedField" type="s:ST_TrueFalseBlank"/>
 * ```
 */
export const createOLockedField = (): XmlComponent => new BuilderElement({ name: "o:LockedField" });

/**
 * Creates an o:FieldCodes element (OLE Field Codes).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="FieldCodes" type="xsd:string"/>
 * ```
 */
export const createOFieldCodes = (): XmlComponent => new BuilderElement({ name: "o:FieldCodes" });
