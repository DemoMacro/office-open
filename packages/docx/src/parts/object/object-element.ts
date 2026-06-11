/**
 * Object element for WordprocessingML documents — w:object.
 *
 * Embeds an OLE object in a paragraph using VML shape + objectEmbed/objectLink.
 *
 * Reference: OOXML transitional, wml.xsd, CT_Object / CT_ObjectEmbed / CT_ObjectLink
 *
 * @module
 */
import type { VmlShapeStyle } from "@parts/textbox/shape/shape";

// ── Options ──

export interface ObjectEmbedOptions {
  /** Relationship ID for the embedded object */
  rId: string;
  /** OLE program ID (e.g., "Excel.Sheet.12") */
  progId?: string;
  /** Draw aspect ("content" | "icon") */
  drawAspect?: string;
  /** Shape ID */
  shapeId?: string;
  /** Field codes */
  fieldCodes?: string;
}

export interface ObjectLinkOptions extends ObjectEmbedOptions {
  /** Update mode ("always" | "onCall") */
  updateMode: string;
  /** Locked field */
  lockedField?: boolean;
}

export interface ObjectElementOptions {
  /** VML shape style */
  style?: VmlShapeStyle;
  /** Shape ID */
  shapeId?: string;
  /** Original width in twips */
  dxaOrig?: number;
  /** Original height in twips */
  dyaOrig?: number;
  /** Embedded OLE object */
  embed?: ObjectEmbedOptions;
  /** Linked OLE object */
  link?: ObjectLinkOptions;
  /** ActiveX control reference */
  control?: { name?: string; shapeid?: string; rId?: string };
  /** Movie reference */
  movie?: string;
}
