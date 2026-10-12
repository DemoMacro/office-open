/**
 * Diagram definition header elements.
 *
 * Reference: ISO/IEC 29500-4, dml-diagram.xsd
 * CT_ColorTransformHeader, CT_DiagramDefinitionHeader, CT_StyleDefinitionHeader
 * and their list containers.
 *
 * @module
 */
import { children as childElements, element, escapeXml } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import { parseDiagramExtensions, stringifyDiagramExtensions } from "./diagram-props";
import type { DiagramExtensionOptions } from "./diagram-props";

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
  const attrs: Record<string, string> = { val: escapeXml(options.val) };
  if (options.lang) attrs.lang = escapeXml(options.lang);
  return element(tag, attrs);
};

const createDescEl = (tag: string, options: DiagramDescriptionOptions): string => {
  const attrs: Record<string, string> = { val: escapeXml(options.val) };
  if (options.lang) attrs.lang = escapeXml(options.lang);
  return element(tag, attrs);
};

const createCatEl = (options: DiagramCategoryOptions): string =>
  `<dgm:cat type="${escapeXml(options.type)}" pri="${options.pri}"/>`;

const createCatLst = (categories?: readonly DiagramCategoryOptions[]): string => {
  const catChildren: string[] = [];
  if (categories) {
    for (const cat of categories) {
      catChildren.push(createCatEl(cat));
    }
  }
  return element("dgm:catLst", undefined, catChildren);
};

const parseNames = (el: Element, tag: string): DiagramNameOptions[] =>
  childElements(el, tag).map((child) => ({
    ...(child.attributes?.lang !== undefined ? { lang: String(child.attributes.lang) } : {}),
    val: String(child.attributes?.val ?? ""),
  }));

const parseDescriptions = (el: Element, tag: string): DiagramDescriptionOptions[] =>
  childElements(el, tag).map((child) => ({
    ...(child.attributes?.lang !== undefined ? { lang: String(child.attributes.lang) } : {}),
    val: String(child.attributes?.val ?? ""),
  }));

const parseCategories = (el: Element): DiagramCategoryOptions[] =>
  childElements(findCategoryList(el), "dgm:cat").map((child) => ({
    type: String(child.attributes?.type ?? ""),
    pri: Number(child.attributes?.pri ?? 0),
  }));

const findCategoryList = (el: Element): Element | undefined => childElements(el, "dgm:catLst")[0];

const parseHeaderExtensions = (el: Element): DiagramExtensionOptions[] | undefined => {
  const extLst = childElements(el, "dgm:extLst")[0];
  return extLst ? parseDiagramExtensions(extLst) : undefined;
};

const parseHeaderAttributes = (
  el: Element,
): Pick<LayoutDefinitionHeaderOptions, "uniqueId" | "minVer" | "resId"> => {
  const result: Pick<LayoutDefinitionHeaderOptions, "uniqueId" | "minVer" | "resId"> = {
    uniqueId: String(el.attributes?.uniqueId ?? ""),
  };
  if (el.attributes?.minVer !== undefined) result.minVer = String(el.attributes.minVer);
  if (el.attributes?.resId !== undefined) result.resId = Number(el.attributes.resId);
  return result;
};

const stringifyHeaderExtensions = (extensions?: readonly DiagramExtensionOptions[]): string => {
  if (!extensions?.length) return "";
  return stringifyDiagramExtensions(extensions);
};

const parseHeaderList = <T>(
  el: Element,
  headerTag: string,
  parseHeader: (header: Element) => T,
): readonly T[] => childElements(el, headerTag).map(parseHeader);

const stringifyHeaderList = <T>(
  tag: string,
  headers: readonly T[] | undefined,
  createHeader: (header: T) => string,
): string => {
  const listChildren = headers?.map(createHeader) ?? [];
  return element(tag, undefined, listChildren);
};

// ---------------------------------------------------------------------------
// dgm:colorsDefHdr — color definition header (CT_ColorTransformHeader)
// ---------------------------------------------------------------------------

export interface ColorsDefinitionHeaderOptions {
  uniqueId: string;
  minVer?: string;
  resId?: number;
  title: readonly DiagramNameOptions[];
  desc: readonly DiagramDescriptionOptions[];
  categories?: readonly DiagramCategoryOptions[];
  extensions?: readonly DiagramExtensionOptions[];
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
export const createColorsDefinitionHeader = (options: ColorsDefinitionHeaderOptions): string => {
  const children: string[] = [];
  for (const t of options.title) children.push(createNameEl("dgm:title", t));
  for (const d of options.desc) children.push(createDescEl("dgm:desc", d));
  if (options.categories?.length) children.push(createCatLst(options.categories));
  children.push(stringifyHeaderExtensions(options.extensions));

  const attrs: Record<string, string | number> = { uniqueId: options.uniqueId };
  if (options.minVer !== undefined) attrs.minVer = options.minVer;
  if (options.resId !== undefined) attrs.resId = options.resId;

  return element("dgm:colorsDefHdr", attrs, children);
};

export const parseColorsDefinitionHeader = (el: Element): ColorsDefinitionHeaderOptions => {
  const result: ColorsDefinitionHeaderOptions = {
    ...parseHeaderAttributes(el),
    title: parseNames(el, "dgm:title"),
    desc: parseDescriptions(el, "dgm:desc"),
  };
  const categories = parseCategories(el);
  if (categories.length) result.categories = categories;
  const extensions = parseHeaderExtensions(el);
  if (extensions?.length) result.extensions = extensions;
  return result;
};

export const colorsDefinitionHeaderDesc: CustomDescriptor<ColorsDefinitionHeaderOptions> = {
  kind: "custom",
  stringify: (options) => createColorsDefinitionHeader(options),
  parse: (el) => parseColorsDefinitionHeader(el),
};

// ---------------------------------------------------------------------------
// dgm:colorsDefHdrLst — list of color definition headers (CT_ColorTransformHeaderLst)
// ---------------------------------------------------------------------------

export interface ColorsDefinitionHeaderListOptions {
  headers?: readonly ColorsDefinitionHeaderOptions[];
}

/** Creates a dgm:colorsDefHdrLst element. */
export const createColorsDefinitionHeaderList = (
  options?: ColorsDefinitionHeaderListOptions,
): string => {
  return stringifyHeaderList("dgm:colorsDefHdrLst", options?.headers, createColorsDefinitionHeader);
};

export const parseColorsDefinitionHeaderList = (
  el: Element,
): ColorsDefinitionHeaderListOptions => ({
  headers: parseHeaderList(el, "dgm:colorsDefHdr", parseColorsDefinitionHeader),
});

export const colorsDefinitionHeaderListDesc: CustomDescriptor<ColorsDefinitionHeaderListOptions> = {
  kind: "custom",
  stringify: (options) => createColorsDefinitionHeaderList(options),
  parse: (el) => parseColorsDefinitionHeaderList(el),
};

// ---------------------------------------------------------------------------
// dgm:layoutDefHdr — layout definition header (CT_DiagramDefinitionHeader)
// ---------------------------------------------------------------------------

export interface LayoutDefinitionHeaderOptions {
  uniqueId: string;
  minVer?: string;
  defStyle?: string;
  resId?: number;
  title: readonly DiagramNameOptions[];
  desc: readonly DiagramDescriptionOptions[];
  categories?: readonly DiagramCategoryOptions[];
  extensions?: readonly DiagramExtensionOptions[];
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
export const createLayoutDefinitionHeader = (options: LayoutDefinitionHeaderOptions): string => {
  const children: string[] = [];
  for (const t of options.title) children.push(createNameEl("dgm:title", t));
  for (const d of options.desc) children.push(createDescEl("dgm:desc", d));
  if (options.categories?.length) children.push(createCatLst(options.categories));
  children.push(stringifyHeaderExtensions(options.extensions));

  const attrs: Record<string, string | number> = { uniqueId: options.uniqueId };
  if (options.minVer !== undefined) attrs.minVer = options.minVer;
  if (options.defStyle !== undefined) attrs.defStyle = options.defStyle;
  if (options.resId !== undefined) attrs.resId = options.resId;

  return element("dgm:layoutDefHdr", attrs, children);
};

export const parseLayoutDefinitionHeader = (el: Element): LayoutDefinitionHeaderOptions => {
  const result: LayoutDefinitionHeaderOptions = {
    ...parseHeaderAttributes(el),
    title: parseNames(el, "dgm:title"),
    desc: parseDescriptions(el, "dgm:desc"),
  };
  if (el.attributes?.defStyle !== undefined) result.defStyle = String(el.attributes.defStyle);
  const categories = parseCategories(el);
  if (categories.length) result.categories = categories;
  const extensions = parseHeaderExtensions(el);
  if (extensions?.length) result.extensions = extensions;
  return result;
};

export const layoutDefinitionHeaderDesc: CustomDescriptor<LayoutDefinitionHeaderOptions> = {
  kind: "custom",
  stringify: (options) => createLayoutDefinitionHeader(options),
  parse: (el) => parseLayoutDefinitionHeader(el),
};

// ---------------------------------------------------------------------------
// dgm:layoutDefHdrLst — list of layout definition headers (CT_DiagramDefinitionHeaderLst)
// ---------------------------------------------------------------------------

export interface LayoutDefinitionHeaderListOptions {
  headers?: readonly LayoutDefinitionHeaderOptions[];
}

/** Creates a dgm:layoutDefHdrLst element. */
export const createLayoutDefinitionHeaderList = (
  options?: LayoutDefinitionHeaderListOptions,
): string => {
  return stringifyHeaderList("dgm:layoutDefHdrLst", options?.headers, createLayoutDefinitionHeader);
};

export const parseLayoutDefinitionHeaderList = (
  el: Element,
): LayoutDefinitionHeaderListOptions => ({
  headers: parseHeaderList(el, "dgm:layoutDefHdr", parseLayoutDefinitionHeader),
});

export const layoutDefinitionHeaderListDesc: CustomDescriptor<LayoutDefinitionHeaderListOptions> = {
  kind: "custom",
  stringify: (options) => createLayoutDefinitionHeaderList(options),
  parse: (el) => parseLayoutDefinitionHeaderList(el),
};

// ---------------------------------------------------------------------------
// dgm:styleDefHdr — style definition header (CT_StyleDefinitionHeader)
// ---------------------------------------------------------------------------

export interface StyleDefinitionHeaderOptions {
  uniqueId: string;
  minVer?: string;
  resId?: number;
  title: readonly DiagramNameOptions[];
  desc: readonly DiagramDescriptionOptions[];
  categories?: readonly DiagramCategoryOptions[];
  extensions?: readonly DiagramExtensionOptions[];
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
export const createStyleDefinitionHeader = (options: StyleDefinitionHeaderOptions): string => {
  const children: string[] = [];
  for (const t of options.title) children.push(createNameEl("dgm:title", t));
  for (const d of options.desc) children.push(createDescEl("dgm:desc", d));
  if (options.categories?.length) children.push(createCatLst(options.categories));
  children.push(stringifyHeaderExtensions(options.extensions));

  const attrs: Record<string, string | number> = { uniqueId: options.uniqueId };
  if (options.minVer !== undefined) attrs.minVer = options.minVer;
  if (options.resId !== undefined) attrs.resId = options.resId;

  return element("dgm:styleDefHdr", attrs, children);
};

export const parseStyleDefinitionHeader = (el: Element): StyleDefinitionHeaderOptions => {
  const result: StyleDefinitionHeaderOptions = {
    ...parseHeaderAttributes(el),
    title: parseNames(el, "dgm:title"),
    desc: parseDescriptions(el, "dgm:desc"),
  };
  const categories = parseCategories(el);
  if (categories.length) result.categories = categories;
  const extensions = parseHeaderExtensions(el);
  if (extensions?.length) result.extensions = extensions;
  return result;
};

export const styleDefinitionHeaderDesc: CustomDescriptor<StyleDefinitionHeaderOptions> = {
  kind: "custom",
  stringify: (options) => createStyleDefinitionHeader(options),
  parse: (el) => parseStyleDefinitionHeader(el),
};

// ---------------------------------------------------------------------------
// dgm:styleDefHdrLst — list of style definition headers (CT_StyleDefinitionHeaderLst)
// ---------------------------------------------------------------------------

export interface StyleDefinitionHeaderListOptions {
  headers?: readonly StyleDefinitionHeaderOptions[];
}

/** Creates a dgm:styleDefHdrLst element. */
export const createStyleDefinitionHeaderList = (
  options?: StyleDefinitionHeaderListOptions,
): string => {
  return stringifyHeaderList("dgm:styleDefHdrLst", options?.headers, createStyleDefinitionHeader);
};

export const parseStyleDefinitionHeaderList = (el: Element): StyleDefinitionHeaderListOptions => ({
  headers: parseHeaderList(el, "dgm:styleDefHdr", parseStyleDefinitionHeader),
});

export const styleDefinitionHeaderListDesc: CustomDescriptor<StyleDefinitionHeaderListOptions> = {
  kind: "custom",
  stringify: (options) => createStyleDefinitionHeaderList(options),
  parse: (el) => parseStyleDefinitionHeaderList(el),
};
