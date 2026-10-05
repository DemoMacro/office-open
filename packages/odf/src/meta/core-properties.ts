import type { CorePropertiesOptions } from "@office-open/core";

import { readXml, type OdfFiles } from "../container/package";
import { childNamed, escapeText, textOf } from "../runtime/xml";

export const ODF_NAMESPACES = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"',
  'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"',
].join(" ");

export function metaXml(options: CorePropertiesOptions): string {
  const fields = [
    options.title && `<dc:title>${escapeText(options.title)}</dc:title>`,
    options.subject && `<dc:subject>${escapeText(options.subject)}</dc:subject>`,
    options.creator && `<dc:creator>${escapeText(options.creator)}</dc:creator>`,
    options.keywords && `<meta:keyword>${escapeText(options.keywords)}</meta:keyword>`,
    options.description && `<dc:description>${escapeText(options.description)}</dc:description>`,
    options.lastModifiedBy &&
      `<meta:initial-creator>${escapeText(options.lastModifiedBy)}</meta:initial-creator>`,
    options.created && `<meta:creation-date>${escapeText(options.created)}</meta:creation-date>`,
    options.modified && `<dc:date>${escapeText(options.modified)}</dc:date>`,
    options.category && `<office:category>${escapeText(options.category)}</office:category>`,
    options.language && `<dc:language>${escapeText(options.language)}</dc:language>`,
    options.identifier && `<dc:identifier>${escapeText(options.identifier)}</dc:identifier>`,
    options.lastPrinted && `<meta:print-date>${escapeText(options.lastPrinted)}</meta:print-date>`,
    options.revision !== undefined &&
      `<meta:editing-cycles>${options.revision}</meta:editing-cycles>`,
  ].filter(Boolean);
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-meta ${ODF_NAMESPACES} office:version="1.3"><office:meta>${fields.join(
    "",
  )}</office:meta></office:document-meta>`;
}

export function parseMeta(files: OdfFiles): CorePropertiesOptions {
  if (!files["meta.xml"]) return {};
  const meta = childNamed(readXml(files, "meta.xml"), "office:meta");
  const value = (name: string): string | undefined => textOf(childNamed(meta, name)) || undefined;
  const revision = value("meta:editing-cycles");
  return {
    title: value("dc:title"),
    subject: value("dc:subject"),
    creator: value("dc:creator"),
    keywords: value("meta:keyword"),
    description: value("dc:description"),
    lastModifiedBy: value("meta:initial-creator"),
    created: value("meta:creation-date"),
    modified: value("dc:date"),
    category: value("office:category"),
    language: value("dc:language"),
    identifier: value("dc:identifier"),
    lastPrinted: value("meta:print-date"),
    revision: revision ? Number(revision) : undefined,
  };
}
