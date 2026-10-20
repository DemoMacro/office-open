import { stringifyElement } from "@office-open/xml";
import type { Element } from "@office-open/xml";

/** ODF slide node without a PPTX canonical equivalent. */
export interface OdpOpaqueSlideNodeOptions {
  /** Zero-based position among the page's visible drawing children. */
  position: number;
  /** Source ODF element serialized for ODP-only round-trip. */
  xml: string;
}

const OPAQUE_SLIDE_NODES = new Set(["dr3d:scene", "draw:measure", "draw:page-thumbnail"]);

export function isOpaqueSlideNode(element: Element): boolean {
  return OPAQUE_SLIDE_NODES.has(element.name ?? "");
}

export function opaqueSlideNode(element: Element, position: number): OdpOpaqueSlideNodeOptions {
  return { position, xml: stringifyElement(element) };
}
