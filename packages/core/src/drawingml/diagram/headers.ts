/**
 * Diagram definition header elements.
 *
 * Reference: ISO/IEC 29500-4, dml-diagram.xsd
 * CT_ColorTransformHeader, CT_DiagramDefinitionHeader, CT_StyleDefinitionHeader
 * and their list containers.
 *
 * @module
 */
import { BuilderElement, type XmlComponent } from "../../xml-components";

// ---------------------------------------------------------------------------
// Shared name/description/category types
// ---------------------------------------------------------------------------

export interface DiagramNameOptions {
  lang?: string;
  val: string;
}

export interface DiagramDescriptionOptions {
  lang?: string;
  val: string;
}

export interface DiagramCategoryOptions {
  type: string;
  pri: number;
}

const createNameEl = (tag: string, options: DiagramNameOptions): XmlComponent => {
  const attrs: Record<string, string> = { val: options.val };
  if (options.lang) attrs.lang = options.lang;
  return new BuilderElement({
    name: tag,
    attributes: {
      val: { key: "val", value: attrs.val },
      ...(attrs.lang && { lang: { key: "lang", value: attrs.lang } }),
    },
  });
};

const createDescEl = (tag: string, options: DiagramDescriptionOptions): XmlComponent => {
  const attrs: Record<string, string> = { val: options.val };
  if (options.lang) attrs.lang = options.lang;
  return new BuilderElement({
    name: tag,
    attributes: {
      val: { key: "val", value: attrs.val },
      ...(attrs.lang && { lang: { key: "lang", value: attrs.lang } }),
    },
  });
};

const createCatEl = (options: DiagramCategoryOptions): XmlComponent =>
  new BuilderElement({
    name: "dgm:cat",
    attributes: {
      type: { key: "type", value: options.type },
      pri: { key: "pri", value: options.pri },
    },
  });

const createCatLst = (categories?: readonly DiagramCategoryOptions[]): XmlComponent => {
  const children: XmlComponent[] = [];
  if (categories) {
    for (const cat of categories) {
      children.push(createCatEl(cat));
    }
  }
  return new BuilderElement({ name: "dgm:catLst", children });
};

// ---------------------------------------------------------------------------
// dgm:colorsDefHdr — color definition header (CT_ColorTransformHeader)
// ---------------------------------------------------------------------------

export interface ColorsDefHdrOptions {
  uniqueId: string;
  minVer?: string;
  resId?: number;
  title: readonly DiagramNameOptions[];
  desc: readonly DiagramDescriptionOptions[];
  categories?: readonly DiagramCategoryOptions[];
}

/**
 * Creates a dgm:colorsDefHdr element (CT_ColorTransformHeader).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_ColorTransformHeader">
 *   <xsd:sequence>
 *     <xsd:element name="title" type="CT_CTName" minOccurs="1" maxOccurs="unbounded"/>
 *     <xsd:element name="desc" type="CT_CTDescription" minOccurs="1" maxOccurs="unbounded"/>
 *     <xsd:element name="catLst" type="CT_CTCategories" minOccurs="0"/>
 *     <xsd:element name="extLst" type="a:CT_OfficeArtExtensionList" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="uniqueId" type="xsd:string" use="required"/>
 *   <xsd:attribute name="minVer" type="xsd:string" use="optional"/>
 *   <xsd:attribute name="resId" type="xsd:int" use="optional" default="0"/>
 * </xsd:complexType>
 * ```
 */
export const createColorsDefHdr = (options: ColorsDefHdrOptions): XmlComponent => {
  const children: XmlComponent[] = [];
  for (const t of options.title) children.push(createNameEl("dgm:title", t));
  for (const d of options.desc) children.push(createDescEl("dgm:desc", d));
  if (options.categories?.length) children.push(createCatLst(options.categories));

  return new BuilderElement({
    name: "dgm:colorsDefHdr",
    attributes: {
      uniqueId: { key: "uniqueId", value: options.uniqueId },
      ...(options.minVer !== undefined && { minVer: { key: "minVer", value: options.minVer } }),
      ...(options.resId !== undefined && { resId: { key: "resId", value: options.resId } }),
    },
    children,
  });
};

// ---------------------------------------------------------------------------
// dgm:colorsDefHdrLst — list of color definition headers (CT_ColorTransformHeaderLst)
// ---------------------------------------------------------------------------

export interface ColorsDefHdrLstOptions {
  headers?: readonly ColorsDefHdrOptions[];
}

/** Creates a dgm:colorsDefHdrLst element. */
export const createColorsDefHdrLst = (options?: ColorsDefHdrLstOptions): XmlComponent => {
  const children: XmlComponent[] = [];
  if (options?.headers) {
    for (const hdr of options.headers) {
      children.push(createColorsDefHdr(hdr));
    }
  }
  return new BuilderElement({ name: "dgm:colorsDefHdrLst", children });
};

// ---------------------------------------------------------------------------
// dgm:layoutDefHdr — layout definition header (CT_DiagramDefinitionHeader)
// ---------------------------------------------------------------------------

export interface LayoutDefHdrOptions {
  uniqueId: string;
  minVer?: string;
  defStyle?: string;
  resId?: number;
  title: readonly DiagramNameOptions[];
  desc: readonly DiagramDescriptionOptions[];
  categories?: readonly DiagramCategoryOptions[];
}

/**
 * Creates a dgm:layoutDefHdr element (CT_DiagramDefinitionHeader).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_DiagramDefinitionHeader">
 *   <xsd:sequence>
 *     <xsd:element name="title" type="CT_Name" minOccurs="1" maxOccurs="unbounded"/>
 *     <xsd:element name="desc" type="CT_Description" minOccurs="1" maxOccurs="unbounded"/>
 *     <xsd:element name="catLst" type="CT_Categories" minOccurs="0"/>
 *     <xsd:element name="extLst" type="a:CT_OfficeArtExtensionList" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="uniqueId" type="xsd:string" use="required"/>
 *   <xsd:attribute name="minVer" type="xsd:string" use="optional"/>
 *   <xsd:attribute name="defStyle" type="xsd:string" use="optional" default=""/>
 *   <xsd:attribute name="resId" type="xsd:int" use="optional" default="0"/>
 * </xsd:complexType>
 * ```
 */
export const createLayoutDefHdr = (options: LayoutDefHdrOptions): XmlComponent => {
  const children: XmlComponent[] = [];
  for (const t of options.title) children.push(createNameEl("dgm:title", t));
  for (const d of options.desc) children.push(createDescEl("dgm:desc", d));
  if (options.categories?.length) children.push(createCatLst(options.categories));

  return new BuilderElement({
    name: "dgm:layoutDefHdr",
    attributes: {
      uniqueId: { key: "uniqueId", value: options.uniqueId },
      ...(options.minVer !== undefined && { minVer: { key: "minVer", value: options.minVer } }),
      ...(options.defStyle !== undefined && {
        defStyle: { key: "defStyle", value: options.defStyle },
      }),
      ...(options.resId !== undefined && { resId: { key: "resId", value: options.resId } }),
    },
    children,
  });
};

// ---------------------------------------------------------------------------
// dgm:layoutDefHdrLst — list of layout definition headers (CT_DiagramDefinitionHeaderLst)
// ---------------------------------------------------------------------------

export interface LayoutDefHdrLstOptions {
  headers?: readonly LayoutDefHdrOptions[];
}

/** Creates a dgm:layoutDefHdrLst element. */
export const createLayoutDefHdrLst = (options?: LayoutDefHdrLstOptions): XmlComponent => {
  const children: XmlComponent[] = [];
  if (options?.headers) {
    for (const hdr of options.headers) {
      children.push(createLayoutDefHdr(hdr));
    }
  }
  return new BuilderElement({ name: "dgm:layoutDefHdrLst", children });
};

// ---------------------------------------------------------------------------
// dgm:styleDefHdr — style definition header (CT_StyleDefinitionHeader)
// ---------------------------------------------------------------------------

export interface StyleDefHdrOptions {
  uniqueId: string;
  minVer?: string;
  resId?: number;
  title: readonly DiagramNameOptions[];
  desc: readonly DiagramDescriptionOptions[];
  categories?: readonly DiagramCategoryOptions[];
}

/**
 * Creates a dgm:styleDefHdr element (CT_StyleDefinitionHeader).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_StyleDefinitionHeader">
 *   <xsd:sequence>
 *     <xsd:element name="title" type="CT_SDName" minOccurs="1" maxOccurs="unbounded"/>
 *     <xsd:element name="desc" type="CT_SDDescription" minOccurs="1" maxOccurs="unbounded"/>
 *     <xsd:element name="catLst" type="CT_SDCategories" minOccurs="0"/>
 *     <xsd:element name="extLst" type="a:CT_OfficeArtExtensionList" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="uniqueId" type="xsd:string" use="required"/>
 *   <xsd:attribute name="minVer" type="xsd:string" use="optional"/>
 *   <xsd:attribute name="resId" type="xsd:int" use="optional" default="0"/>
 * </xsd:complexType>
 * ```
 */
export const createStyleDefHdr = (options: StyleDefHdrOptions): XmlComponent => {
  const children: XmlComponent[] = [];
  for (const t of options.title) children.push(createNameEl("dgm:title", t));
  for (const d of options.desc) children.push(createDescEl("dgm:desc", d));
  if (options.categories?.length) children.push(createCatLst(options.categories));

  return new BuilderElement({
    name: "dgm:styleDefHdr",
    attributes: {
      uniqueId: { key: "uniqueId", value: options.uniqueId },
      ...(options.minVer !== undefined && { minVer: { key: "minVer", value: options.minVer } }),
      ...(options.resId !== undefined && { resId: { key: "resId", value: options.resId } }),
    },
    children,
  });
};

// ---------------------------------------------------------------------------
// dgm:styleDefHdrLst — list of style definition headers (CT_StyleDefinitionHeaderLst)
// ---------------------------------------------------------------------------

export interface StyleDefHdrLstOptions {
  headers?: readonly StyleDefHdrOptions[];
}

/** Creates a dgm:styleDefHdrLst element. */
export const createStyleDefHdrLst = (options?: StyleDefHdrLstOptions): XmlComponent => {
  const children: XmlComponent[] = [];
  if (options?.headers) {
    for (const hdr of options.headers) {
      children.push(createStyleDefHdr(hdr));
    }
  }
  return new BuilderElement({ name: "dgm:styleDefHdrLst", children });
};
