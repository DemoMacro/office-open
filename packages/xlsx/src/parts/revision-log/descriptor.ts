/**
 * Revision log descriptors for shared-workbook change tracking.
 *
 * Three parts:
 *   - xl/revisionHeaders.xml        (CT_RevisionHeaders) — index of revision logs
 *   - xl/revisions/revisionN.xml    (CT_Revisions)        — one log per header entry
 *   - xl/users.xml                  (CT_Users)            — shared-workbook users
 *
 * Reference: OOXML transitional, sml.xsd — CT_RevisionHeaders (:1860),
 * CT_RevisionHeader (:1898), CT_Revisions (:1877, 12 revision elements),
 * CT_Users (:2100), AG_RevData (:1893).
 *
 * Nested CT_Cell (oc/nc) and CT_Dxf (odxf/ndxf/dxf) carry through verbatim as
 * `*Xml` strings — their full content model is large and revision round-trip only
 * needs lossless preservation, not a re-authored abstraction.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import type { Element as XmlElement } from "@office-open/xml";

import { parseEntry } from "./parse";
import { S_NS, stringifyEntry } from "./stringify";
import type { RevisionLogOptions } from "./types";

export const revisionLogDesc: CustomDescriptor<RevisionLogOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    if (opts.revisions.length === 0) return undefined;
    return `<revisions xmlns="${S_NS}">${opts.revisions.map(stringifyEntry).filter(Boolean).join("")}</revisions>`;
  },

  parse(el, _ctx) {
    const revisions: RevisionLogOptions["revisions"] = [];
    for (const child of el.elements ?? []) {
      if (child.type !== "element") continue;
      const entry = parseEntry(child as XmlElement);
      if (entry) revisions.push(entry);
    }
    return { revisions } as RevisionLogOptions;
  },
};
