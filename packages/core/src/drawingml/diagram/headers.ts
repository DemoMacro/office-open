/**
 * Diagram definition header elements.
 *
 * Reference: ISO/IEC 29500-4, dml-diagram.xsd
 * CT_ColorTransformHeader, CT_DiagramDefinitionHeader, CT_StyleDefinitionHeader
 * and their list containers.
 *
 * @module
 */
import { element } from "@office-open/xml";

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

const createNameEl = (tag: string, options: DiagramNameOptions): string => {
  const attrs: Record<string, string> = { val: options.val };
  if (options.lang) attrs.lang = options.lang;
  return element(tag, attrs);
};

const createDescEl = (tag: string, options: DiagramDescriptionOptions): string => {
  const attrs: Record<string, string> = { val: options.val };
  if (options.lang) attrs.lang = options.lang;
  return element(tag, attrs);
};

const createCatEl = (options: DiagramCategoryOptions): string =>
  `<dgm:cat type="${options.type}" pri="${options.pri}"/>`;

const createCatLst = (categories?: readonly DiagramCategoryOptions[]): string => {
  const children: string[] = [];
  if (categories) {
    for (const cat of categories) {
      children.push(createCatEl(cat));
    }
  }
  return element("dgm:catLst", undefined, children);
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
export const createColorsDefHdr = (options: ColorsDefHdrOptions): string => {
  const children: string[] = [];
  for (const t of options.title) children.push(createNameEl("dgm:title", t));
  for (const d of options.desc) children.push(createDescEl("dgm:desc", d));
  if (options.categories?.length) children.push(createCatLst(options.categories));

  const attrs: Record<string, string | number> = { uniqueId: options.uniqueId };
  if (options.minVer !== undefined) attrs.minVer = options.minVer;
  if (options.resId !== undefined) attrs.resId = options.resId;

  return element("dgm:colorsDefHdr", attrs, children);
};

// ---------------------------------------------------------------------------
// dgm:colorsDefHdrLst — list of color definition headers (CT_ColorTransformHeaderLst)
// ---------------------------------------------------------------------------

export interface ColorsDefHdrLstOptions {
  headers?: readonly ColorsDefHdrOptions[];
}

/** Creates a dgm:colorsDefHdrLst element. */
export const createColorsDefHdrLst = (options?: ColorsDefHdrLstOptions): string => {
  const children: string[] = [];
  if (options?.headers) {
    for (const hdr of options.headers) {
      children.push(createColorsDefHdr(hdr));
    }
  }
  return element("dgm:colorsDefHdrLst", undefined, children);
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
export const createLayoutDefHdr = (options: LayoutDefHdrOptions): string => {
  const children: string[] = [];
  for (const t of options.title) children.push(createNameEl("dgm:title", t));
  for (const d of options.desc) children.push(createDescEl("dgm:desc", d));
  if (options.categories?.length) children.push(createCatLst(options.categories));

  const attrs: Record<string, string | number> = { uniqueId: options.uniqueId };
  if (options.minVer !== undefined) attrs.minVer = options.minVer;
  if (options.defStyle !== undefined) attrs.defStyle = options.defStyle;
  if (options.resId !== undefined) attrs.resId = options.resId;

  return element("dgm:layoutDefHdr", attrs, children);
};

// ---------------------------------------------------------------------------
// dgm:layoutDefHdrLst — list of layout definition headers (CT_DiagramDefinitionHeaderLst)
// ---------------------------------------------------------------------------

export interface LayoutDefHdrLstOptions {
  headers?: readonly LayoutDefHdrOptions[];
}

/** Creates a dgm:layoutDefHdrLst element. */
export const createLayoutDefHdrLst = (options?: LayoutDefHdrLstOptions): string => {
  const children: string[] = [];
  if (options?.headers) {
    for (const hdr of options.headers) {
      children.push(createLayoutDefHdr(hdr));
    }
  }
  return element("dgm:layoutDefHdrLst", undefined, children);
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
export const createStyleDefHdr = (options: StyleDefHdrOptions): string => {
  const children: string[] = [];
  for (const t of options.title) children.push(createNameEl("dgm:title", t));
  for (const d of options.desc) children.push(createDescEl("dgm:desc", d));
  if (options.categories?.length) children.push(createCatLst(options.categories));

  const attrs: Record<string, string | number> = { uniqueId: options.uniqueId };
  if (options.minVer !== undefined) attrs.minVer = options.minVer;
  if (options.resId !== undefined) attrs.resId = options.resId;

  return element("dgm:styleDefHdr", attrs, children);
};

// ---------------------------------------------------------------------------
// dgm:styleDefHdrLst — list of style definition headers (CT_StyleDefinitionHeaderLst)
// ---------------------------------------------------------------------------

export interface StyleDefHdrLstOptions {
  headers?: readonly StyleDefHdrOptions[];
}

/** Creates a dgm:styleDefHdrLst element. */
export const createStyleDefHdrLst = (options?: StyleDefHdrLstOptions): string => {
  const children: string[] = [];
  if (options?.headers) {
    for (const hdr of options.headers) {
      children.push(createStyleDefHdr(hdr));
    }
  }
  return element("dgm:styleDefHdrLst", undefined, children);
};
