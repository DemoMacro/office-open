/**
 * Header/Footer entry module for WordprocessingML documents.
 *
 * Replaces the former HeaderWrapper/FooterWrapper/Header/Footer/HeaderFooterBase
 * class hierarchy with a simple data structure + pure serialization function.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_HdrFtr
 *
 * @module
 */

import type { Relationships } from "@office-open/core";
import { escapeXml } from "@office-open/xml";
import type { SectionChild } from "@shared/section";

import { stringifyBodyChild } from "../body";
import type { BodyContext } from "../context";
import { DocumentAttributeNamespaces } from "./document/document-attributes";
import type { DocumentAttributeNamespace } from "./document/document-attributes";

/**
 * Simple data structure for a header or footer entry.
 *
 * Replaces HeaderWrapper/FooterWrapper — holds children, relationships,
 * and the reference ID needed for section property references.
 *
 * Children are raw SectionChild objects (plain JSON or class instances).
 */
export interface HeaderFooterEntry {
  children: SectionChild[];
  relationships: Relationships;
  referenceId: number;
}

/**
 * Namespace keys used by header elements.
 * @internal
 */
export const HEADER_NAMESPACES: DocumentAttributeNamespace[] = [
  "cx",
  "cx1",
  "cx2",
  "cx3",
  "cx4",
  "cx5",
  "cx6",
  "cx7",
  "cx8",
  "m",
  "mc",
  "o",
  "r",
  "v",
  "w",
  "w10",
  "w14",
  "w15",
  "w16cid",
  "w16se",
  "wne",
  "wp",
  "wp14",
  "wpc",
  "wpg",
  "wpi",
  "wps",
];

/**
 * Namespace keys used by footer elements.
 * @internal
 */
export const FOOTER_NAMESPACES: DocumentAttributeNamespace[] = [
  "m",
  "mc",
  "o",
  "r",
  "v",
  "w",
  "w10",
  "w14",
  "w15",
  "wne",
  "wp",
  "wp14",
  "wpc",
  "wpg",
  "wpi",
  "wps",
];

/**
 * Serialize a header or footer to XML.
 *
 * Builds the `<w:hdr>` or `<w:ftr>` element with namespace declarations,
 * then serializes each child element via `stringifyBodyChild()`.
 *
 * @param tag - Element tag name ("w:hdr" or "w:ftr")
 * @param namespaces - Namespace keys to declare on the root element
 * @param children - Block-level child elements (raw SectionChild objects)
 * @param ctx - Body context for stringification
 */
export function stringifyHeaderFooter(
  tag: string,
  namespaces: DocumentAttributeNamespace[],
  children: SectionChild[],
  ctx: BodyContext,
): string {
  const attrParts: string[] = [];
  for (const ns of namespaces) {
    attrParts.push(`xmlns:${ns}="${escapeXml(DocumentAttributeNamespaces[ns])}"`);
  }
  // mc:Ignorable must declare the ignorable namespaces (w14/w15/wp14) that
  // header/footer content uses (e.g. w14:paraId). Without it, Word in
  // compatibility mode 14 rejects the part as unreadable content.
  attrParts.push('mc:Ignorable="w14 w15 wp14"');
  const attrStr = attrParts.join(" ");

  const childParts: string[] = [];
  for (const child of children) {
    childParts.push(stringifyBodyChild(child, ctx));
  }

  const body = childParts.join("");
  return body.length === 0 ? `<${tag} ${attrStr}/>` : `<${tag} ${attrStr}>${body}</${tag}>`;
}
