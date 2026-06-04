/**
 * VML Main elements (v: namespace).
 *
 * Factory functions for VML shape primitives, formulas, handles,
 * image data, and text path elements.
 *
 * Reference: ISO/IEC 29500-4, vml-main.xsd
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

// ── Shape Primitives ──

/**
 * Creates a v:rect element (Rectangle).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="rect" type="CT_Rect"/>
 * ```
 */
export const createVmlRect = (): XmlComponent => new BuilderElement({ name: "v:rect" });

/**
 * Creates a v:roundrect element (Rounded Rectangle).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="roundrect" type="CT_RoundRect"/>
 * ```
 */
export const createVmlRoundRect = (): XmlComponent => new BuilderElement({ name: "v:roundrect" });

/**
 * Creates a v:oval element (Oval / Ellipse).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="oval" type="CT_Oval"/>
 * ```
 */
export const createVmlOval = (): XmlComponent => new BuilderElement({ name: "v:oval" });

/**
 * Creates a v:line element (Line).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="line" type="CT_Line"/>
 * ```
 */
export const createVmlLine = (): XmlComponent => new BuilderElement({ name: "v:line" });

/**
 * Creates a v:arc element (Arc).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="arc" type="CT_Arc"/>
 * ```
 */
export const createVmlArc = (): XmlComponent => new BuilderElement({ name: "v:arc" });

/**
 * Creates a v:curve element (Bezier Curve).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="curve" type="CT_Curve"/>
 * ```
 */
export const createVmlCurve = (): XmlComponent => new BuilderElement({ name: "v:curve" });

/**
 * Creates a v:polyline element (Polyline).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="polyline" type="CT_PolyLine"/>
 * ```
 */
export const createVmlPolyline = (): XmlComponent => new BuilderElement({ name: "v:polyline" });

/**
 * Creates a v:image element (Image Shape).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="image" type="CT_Image"/>
 * ```
 */
export const createVmlImage = (): XmlComponent => new BuilderElement({ name: "v:image" });

/**
 * Creates a v:group element (Shape Group).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="group" type="CT_Group"/>
 * ```
 */
export const createVmlGroup = (): XmlComponent => new BuilderElement({ name: "v:group" });

// ── Image Data ──

/**
 * Creates a v:imagedata element (Image Data).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="imagedata" type="CT_ImageData"/>
 * ```
 */
export const createVmlImagedata = (): XmlComponent => new BuilderElement({ name: "v:imagedata" });

// ── Formulas & Handles ──

/**
 * Creates a v:formulas element (Formulas Container).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="formulas" type="CT_Formulas"/>
 * ```
 */
export const createVmlFormulas = (): XmlComponent => new BuilderElement({ name: "v:formulas" });

/**
 * Creates a v:f element (Formula).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="f" type="CT_F"/>
 * ```
 */
export const createVmlF = (): XmlComponent => new BuilderElement({ name: "v:f" });

/**
 * Creates a v:handles element (Handles Container).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="handles" type="CT_Handles"/>
 * ```
 */
export const createVmlHandles = (): XmlComponent => new BuilderElement({ name: "v:handles" });

/**
 * Creates a v:h element (Handle).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="h" type="CT_H"/>
 * ```
 */
export const createVmlH = (): XmlComponent => new BuilderElement({ name: "v:h" });

// ── Text Path ──

/**
 * Creates a v:textpath element (Text Path).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="textpath" type="CT_TextPath"/>
 * ```
 */
export const createVmlTextpath = (): XmlComponent => new BuilderElement({ name: "v:textpath" });
