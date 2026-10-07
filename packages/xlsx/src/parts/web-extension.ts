/**
 * Web extension part types and descriptor for SpreadsheetML documents.

 *
 * The `we:` namespace is a Microsoft extension not covered by the ISO
 * transitional schemas; the descriptor preserves the full typed surface so
 * the part never falls back to raw passthrough.
 *
 * Reference: MS-ODRAWXML, CT_WebExtension
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, escapeXml, findChild, stringifyElement } from "@office-open/xml";

const WE_NS = "http://schemas.microsoft.com/office/webextensions/webextension/2010/11";

/** Web extension store reference (CT_WebExtensionRef). */
export interface WebExtensionReferenceOptions {
  /** Extension identifier inside the store (required). */
  id: string;
  /** Extension version (required). */
  version: string;
  /** Store marketplace identifier; empty string = default Office store. */
  store?: string;
  /** Store type URI (e.g. EXCatalog); omitted = unspecified. */
  storeType?: string;
}

/** Named extension property (CT_WebExtensionProperty). */
export interface WebExtensionPropertyOptions {
  name: string;
  value: string;
}

/** Content binding (CT_WebExtensionBinding). */
export interface WebExtensionBindingOptions {
  /** Binding identifier (required). */
  id: string;
  /** Binding kind: "table" | "text" | "matrix". */
  type: "table" | "text" | "matrix";
  /** Bound range reference. */
  ref: string;
}

/** Options for xl/webextensions/webextension{n}.xml (CT_WebExtension). */
export interface WebExtensionPartOptions {
  /** Source package path; round-trip only. */
  sourcePath?: string;
  /** Extension instance identifier (GUID string, required). */
  id: string;
  /** Store reference (we:reference). */
  reference?: WebExtensionReferenceOptions;
  /** Alternate store references (we:alternateReferences container). */
  alternateReferences?: WebExtensionReferenceOptions[];
  /** Named property bag (we:properties). */
  properties?: WebExtensionPropertyOptions[];
  /** Content bindings (we:bindings). */
  bindings?: WebExtensionBindingOptions[];
  /** Frozen snapshot image relationship id (we:snapshot/@r:embed). */
  snapshotRId?: string;
  /** True when the source declared we:snapshot (even without r:embed). */
  snapshotPresent?: boolean;
  /** Extension list preserved verbatim (we:extLst). */
  extLst?: string;
}

// ── Descriptor ──

export const webExtensionPartDesc: CustomDescriptor<WebExtensionPartOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const children: string[] = [];
    if (opts.reference) {
      const r = opts.reference;
      children.push(
        `<we:reference${[
          ` id="${escapeXml(r.id)}"`,
          ` version="${escapeXml(r.version)}"`,
          ...(r.store !== undefined ? [` store="${escapeXml(r.store)}"`] : []),
          ...(r.storeType !== undefined ? [` storeType="${escapeXml(r.storeType)}"`] : []),
        ].join("")}/>`,
      );
    }
    if (opts.alternateReferences !== undefined) {
      children.push(
        opts.alternateReferences.length
          ? `<we:alternateReferences>${opts.alternateReferences
              .map(
                (r) =>
                  `<we:reference${[
                    ` id="${escapeXml(r.id)}"`,
                    ` version="${escapeXml(r.version)}"`,
                    ...(r.store !== undefined ? [` store="${escapeXml(r.store)}"`] : []),
                    ...(r.storeType !== undefined
                      ? [` storeType="${escapeXml(r.storeType)}"`]
                      : []),
                  ].join("")}/>`,
              )
              .join("")}</we:alternateReferences>`
          : "<we:alternateReferences/>",
      );
    }
    if (opts.properties?.length) {
      children.push(
        `<we:properties>${opts.properties
          .map((p) => `<we:property name="${escapeXml(p.name)}" value="${escapeXml(p.value)}"/>`)
          .join("")}</we:properties>`,
      );
    } else if (opts.properties !== undefined) {
      children.push("<we:properties/>");
    }
    if (opts.bindings?.length) {
      children.push(
        `<we:bindings>${opts.bindings
          .map(
            (b) =>
              `<we:binding id="${escapeXml(b.id)}" type="${b.type}" ref="${escapeXml(b.ref)}"/>`,
          )
          .join("")}</we:bindings>`,
      );
    } else if (opts.bindings !== undefined) {
      children.push("<we:bindings/>");
    }
    if (opts.snapshotPresent) {
      children.push(
        `<we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
          (opts.snapshotRId !== undefined ? ` r:embed="${escapeXml(opts.snapshotRId)}"` : "") +
          "/>",
      );
    }
    if (opts.extLst) children.push(opts.extLst);
    return `<we:webextension xmlns:we="${WE_NS}" id="${escapeXml(opts.id)}">${children.join("")}</we:webextension>`;
  },

  parse(el, _ctx) {
    const result: WebExtensionPartOptions = { id: "" };
    const id = attr(el, "id");
    if (id !== undefined) result.id = id;
    const refEl = findChild(el, "we:reference");
    if (refEl) {
      const refId = attr(refEl, "id");
      const refVersion = attr(refEl, "version");
      if (refId !== undefined && refVersion !== undefined) {
        result.reference = {
          id: refId,
          version: refVersion,
          ...(attr(refEl, "store") !== undefined ? { store: attr(refEl, "store") } : {}),
          ...(attr(refEl, "storeType") !== undefined
            ? { storeType: attr(refEl, "storeType") }
            : {}),
        };
      }
    }
    const propsEl = findChild(el, "we:properties");
    if (propsEl) {
      const properties: WebExtensionPropertyOptions[] = [];
      for (const p of propsEl.elements ?? []) {
        if (p.name !== "we:property") continue;
        const name = attr(p, "name");
        const value = attr(p, "value");
        if (name !== undefined && value !== undefined) properties.push({ name, value });
      }
      result.properties = properties;
    }
    const bindingsEl = findChild(el, "we:bindings");
    if (bindingsEl) {
      const bindings: WebExtensionBindingOptions[] = [];
      for (const b of bindingsEl.elements ?? []) {
        if (b.name !== "we:binding") continue;
        const bid = attr(b, "id");
        const type = attr(b, "type");
        const ref = attr(b, "ref");
        if (bid !== undefined && type !== undefined && ref !== undefined) {
          bindings.push({
            id: bid,
            type: type as WebExtensionBindingOptions["type"],
            ref,
          });
        }
      }
      result.bindings = bindings;
    }
    const altRefsEl = findChild(el, "we:alternateReferences");
    if (altRefsEl) {
      const altRefs: WebExtensionReferenceOptions[] = [];
      for (const r of altRefsEl.elements ?? []) {
        if (r.name !== "we:reference") continue;
        const rid = attr(r, "id");
        const rversion = attr(r, "version");
        if (rid !== undefined && rversion !== undefined) {
          altRefs.push({
            id: rid,
            version: rversion,
            ...(attr(r, "store") !== undefined ? { store: attr(r, "store") } : {}),
            ...(attr(r, "storeType") !== undefined ? { storeType: attr(r, "storeType") } : {}),
          });
        }
      }
      result.alternateReferences = altRefs;
    }
    const snapshotEl = findChild(el, "we:snapshot");
    if (snapshotEl) {
      result.snapshotPresent = true;
      const embed = snapshotEl.attributes?.["r:embed"] as string | undefined;
      if (embed !== undefined) result.snapshotRId = embed;
    }
    const extLstEl = findChild(el, "we:extLst");
    if (extLstEl) result.extLst = stringifyElement(extLstEl);
    return result;
  },
};
