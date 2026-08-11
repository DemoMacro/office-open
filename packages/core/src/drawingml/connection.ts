/**
 * Connector endpoint connection (a:stCxn / a:endCxn / CT_Connection).
 *
 * Glues a connector endpoint to a specific connection site on a target shape,
 * so the endpoint follows the shape when it moves. Used inside cNvCxnSpPr.
 *
 * @module
 */

import type { Element as XmlElement } from "@office-open/xml";

/** a:stCxn / a:endCxn — a connector endpoint glued to a shape connection site. */
export interface EndpointConnectionOptions {
  /** Target shape cNvPr id (ST_DrawingElementId). */
  id: number;
  /** Connection site index on the target shape. */
  index: number;
}

/** Stringify an a:stCxn or a:endCxn element. */
export function stringifyEndpointConnection(
  tag: "stCxn" | "endCxn",
  conn: EndpointConnectionOptions,
): string {
  return `<a:${tag} id="${conn.id}" idx="${conn.index}"/>`;
}

/** Parse an a:stCxn or a:endCxn element. */
export function parseEndpointConnection(el: XmlElement): EndpointConnectionOptions | undefined {
  const id = el.attributes?.["id"];
  const idx = el.attributes?.["idx"];
  if (id === undefined || idx === undefined) return undefined;
  return { id: Number(id), index: Number(idx) };
}
