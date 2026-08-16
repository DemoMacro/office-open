/**
 * v:handles element — CT_Handles (list of v:h adjustment handles).
 *
 * Reference: ISO/IEC 29500-4, vml-main.xsd, CT_Handles / CT_H.
 *
 * @module
 */
import type { Element as XmlElement } from "@office-open/xml";
import { escapeXml } from "@office-open/xml";

import { stringifyVmlTrueFalse, parseVmlTrueFalse, type VmlTrueFalse } from "../attributes";

/** v:h options (CT_H). `switch` matches the XML attribute name. */
export interface VmlHandleOptions {
  position?: string;
  polar?: string;
  map?: string;
  invx?: VmlTrueFalse;
  invy?: VmlTrueFalse;
  switch?: VmlTrueFalse;
  xrange?: string;
  yrange?: string;
  radiusrange?: string;
}

/** v:handles options (CT_Handles). */
export interface VmlHandlesOptions {
  handles?: VmlHandleOptions[];
}

const HANDLE_STRING_FIELDS = [
  "position",
  "polar",
  "map",
  "xrange",
  "yrange",
  "radiusrange",
] as const;
const HANDLE_BOOLEAN_FIELDS = ["invx", "invy", "switch"] as const;

/** Serialize v:handles. */
export function stringifyVmlHandles(opts: VmlHandlesOptions): string {
  const children = (opts.handles ?? []).map((handle) => {
    const attrs: string[] = [];
    for (const field of HANDLE_STRING_FIELDS) {
      const value = handle[field];
      if (value !== undefined) attrs.push(`${field}="${escapeXml(value)}"`);
    }
    for (const field of HANDLE_BOOLEAN_FIELDS) {
      const value = handle[field];
      if (value !== undefined) attrs.push(`${field}="${stringifyVmlTrueFalse(value)}"`);
    }
    const attrStr = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
    return `<v:h${attrStr}/>`;
  });
  return `<v:handles>${children.join("")}</v:handles>`;
}

/** Parse a v:handles element. */
export function parseVmlHandles(el: XmlElement): VmlHandlesOptions {
  const handles: VmlHandleOptions[] = [];
  for (const child of el.elements ?? []) {
    if (child.type !== "element" || child.name !== "v:h") continue;
    const handle: VmlHandleOptions = {};
    const attrs = child.attributes ?? {};
    for (const field of HANDLE_STRING_FIELDS) {
      const raw = attrs[field];
      if (raw !== undefined) handle[field] = String(raw);
    }
    for (const field of HANDLE_BOOLEAN_FIELDS) {
      const raw = attrs[field];
      if (raw !== undefined) handle[field] = parseVmlTrueFalse(String(raw));
    }
    handles.push(handle);
  }
  return handles.length > 0 ? { handles } : {};
}
