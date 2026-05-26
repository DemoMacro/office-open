import { textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import { readXmlFromZip } from "./archive";
import type { IXmlableObject } from "./xml-components";

export interface CoreProperties {
  title?: string;
  subject?: string;
  creator?: string;
  keywords?: string;
  description?: string;
  lastModifiedBy?: string;
  revision?: string;
  created?: string;
  modified?: string;
}

const FIELD_MAP: Array<{ name: string; key: keyof CoreProperties }> = [
  { name: "dc:title", key: "title" },
  { name: "dc:subject", key: "subject" },
  { name: "dc:creator", key: "creator" },
  { name: "dc:description", key: "description" },
  { name: "cp:keywords", key: "keywords" },
  { name: "cp:lastModifiedBy", key: "lastModifiedBy" },
];

export function parseCoreProperties(zip: Map<string, Uint8Array>): CoreProperties {
  const xml = readXmlFromZip(zip, "docProps/core.xml");
  if (!xml) return {};

  return parseCorePropsElement(xml);
}

/**
 * Parse core properties from an already-parsed XML element.
 * Shared by docx and pptx to extract Dublin Core metadata.
 */
export function parseCorePropsElement(el: Element | undefined): CoreProperties {
  if (!el) return {};

  const props: CoreProperties = {};

  for (const field of FIELD_MAP) {
    const child = el.elements?.find((e) => e.name === field.name);
    const value = textOf(child) || undefined;
    if (value) (props as Record<string, unknown>)[field.key] = value;
  }

  const revEl = el.elements?.find((e) => e.name === "cp:revision");
  if (revEl) {
    const rev = textOf(revEl);
    if (rev) {
      const n = Number(rev);
      if (!isNaN(n)) props.revision = rev;
    }
  }

  return props;
}

const CORE_PROPS_NS: IXmlableObject = Object.freeze({
  _attr: Object.freeze({
    "xmlns:cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "xmlns:dc": "http://purl.org/dc/elements/1.1/",
    "xmlns:dcmitype": "http://purl.org/dc/dcmitype/",
    "xmlns:dcterms": "http://purl.org/dc/terms/",
    "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
  }),
});

const W3CDTF_ATTR: IXmlableObject = Object.freeze({
  _attr: Object.freeze({ "xsi:type": "dcterms:W3CDTF" }),
});

/**
 * Build a cp:coreProperties XML object from metadata.
 *
 * Shared by docx and pptx to avoid duplicating namespace declarations
 * and Dublin Core property construction.
 */
export function buildCorePropertiesXml(opts: {
  readonly title?: string;
  readonly subject?: string;
  readonly creator?: string;
  readonly keywords?: string;
  readonly description?: string;
  readonly lastModifiedBy?: string;
  readonly revision?: number;
}): IXmlableObject {
  const children: IXmlableObject[] = [CORE_PROPS_NS];

  if (opts.title) children.push({ "dc:title": [opts.title] });
  if (opts.subject) children.push({ "dc:subject": [opts.subject] });
  if (opts.creator) children.push({ "dc:creator": [opts.creator] });
  if (opts.keywords) children.push({ "cp:keywords": [opts.keywords] });
  if (opts.description) children.push({ "dc:description": [opts.description] });
  if (opts.lastModifiedBy) children.push({ "cp:lastModifiedBy": [opts.lastModifiedBy] });
  if (opts.revision) children.push({ "cp:revision": [String(opts.revision)] });

  const now = new Date().toISOString();
  children.push({ "dcterms:created": [W3CDTF_ATTR, now] });
  children.push({ "dcterms:modified": [W3CDTF_ATTR, now] });

  return { "cp:coreProperties": children };
}
