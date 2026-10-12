/**
 * Diagram extension list, 3D shape, and text properties.
 *
 * Reference: ISO/IEC 29500-4, dml-diagram.xsd, dml-main.xsd
 *
 * @module
 */
import { element, escapeXml, stringifyElement } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { Shape3DOptions } from "../three-d/shape-3d";
import { createShape3D } from "../three-d/shape-3d";

// ---------------------------------------------------------------------------
// dgm:extLst — extension list (a:CT_OfficeArtExtensionList)
// ---------------------------------------------------------------------------

export interface DiagramExtensionListOptions {
  /** Extension URIs */
  extensions?: readonly DiagramExtensionOptions[];
}

export interface DiagramExtensionOptions {
  uri: string;
  /** Foreign-namespace child XML inside a:ext; omit for an empty extension. */
  content?: string;
}

export function stringifyDiagramExtensions(
  extensions?: readonly DiagramExtensionOptions[],
): string {
  if (!extensions?.length) return "";
  const children = extensions
    .map((extension) =>
      extension.content === undefined
        ? `<a:ext uri="${escapeXml(extension.uri)}"/>`
        : `<a:ext uri="${escapeXml(extension.uri)}">${extension.content}</a:ext>`,
    )
    .join("");
  return `<dgm:extLst>${children}</dgm:extLst>`;
}

export function parseDiagramExtensions(extLst: Element): DiagramExtensionOptions[] {
  return (extLst.elements ?? [])
    .filter((child) => child.name === "a:ext")
    .map((extension) => ({
      uri: String(extension.attributes?.["uri"] ?? ""),
      ...(extension.elements?.length
        ? { content: (extension.elements ?? []).map(stringifyElement).join("") }
        : {}),
    }));
}

/**
 * Creates a dgm:extLst element (a:CT_OfficeArtExtensionList).
 *
 * Generic extension list pattern used across OOXML.
 */
export const createDiagramExtensionList = (options?: DiagramExtensionListOptions): string => {
  if (!options?.extensions?.length) return "";
  const children: string[] = [];
  if (options?.extensions) {
    for (const ext of options.extensions) {
      children.push(
        ext.content === undefined
          ? `<a:ext uri="${escapeXml(ext.uri)}"/>`
          : `<a:ext uri="${escapeXml(ext.uri)}">${ext.content}</a:ext>`,
      );
    }
  }
  return element("dgm:extLst", undefined, children);
};

// ---------------------------------------------------------------------------
// dgm:sp3d — 3D shape properties (a:CT_Shape3D)
// ---------------------------------------------------------------------------

/**
 * Creates a dgm:sp3d element wrapping a:CT_Shape3D.
 *
 * Delegates to the shared createShape3D factory but wraps in dgm: namespace context.
 */
export const createDiagramShape3D = (options: Shape3DOptions): string => createShape3D(options);

// ---------------------------------------------------------------------------
// dgm:txPr — text properties (CT_TextProps)
// ---------------------------------------------------------------------------

export interface DiagramTextPropertiesOptions {
  /** Flat text (a:flatTx, the no-3D member of EG_Text3D) — empty marker element */
  flat?: boolean;
}

/**
 * Creates a dgm:txPr element (CT_TextProps).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_TextProps">
 *   <xsd:sequence>
 *     <xsd:group ref="a:EG_Text3D" minOccurs="0" maxOccurs="1"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createDiagramTextProperties = (options?: DiagramTextPropertiesOptions): string =>
  options?.flat ? "<dgm:txPr><a:flatTx/></dgm:txPr>" : "<dgm:txPr/>";
