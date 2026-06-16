/**
 * Diagram extension list, 3D shape, and text properties.
 *
 * Reference: ISO/IEC 29500-4, dml-diagram.xsd, dml-main.xsd
 *
 * @module
 */
import { element } from "@office-open/xml";

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
  content?: string;
}

/**
 * Creates a dgm:extLst element (a:CT_OfficeArtExtensionList).
 *
 * Generic extension list pattern used across OOXML.
 */
export const createDiagramExtensionList = (options?: DiagramExtensionListOptions): string => {
  const children: string[] = [];
  if (options?.extensions) {
    for (const ext of options.extensions) {
      children.push(`<a:ext uri="${ext.uri}"/>`);
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
  /** 3D text properties (flat text, no 3D) — empty element */
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
export const createDiagramTextProperties = (_options?: DiagramTextPropertiesOptions): string =>
  "<dgm:txPr/>";
