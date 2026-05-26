import type { ParsedDocument } from "@office-open/core";
import { parseArchive, parseCorePropsElement } from "@office-open/core";
import { attr } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import { parseBody } from "./parse/body";
import { ParseContext } from "./parse/context";
import { parseCustomProperties } from "./parse/custom-properties";
import { parseSettings } from "./parse/settings";
import { buildStyleCache, buildNumberingCache } from "./parse/styles";

export { parseArchive };

/**
 * All part paths extracted from the DOCX package.
 * Field names correspond directly to the OOXML directory structure.
 */
export interface DocxPartRefs {
  /** word/headerN.xml keyed by rId */
  headers: Map<string, string>;
  /** word/footerN.xml keyed by rId */
  footers: Map<string, string>;
  /** word/footnotes.xml */
  footnotes?: string;
  /** word/endnotes.xml */
  endnotes?: string;
  /** word/comments.xml */
  comments?: string;
  /** word/charts/chartN.xml keyed by rId */
  charts: Map<string, string>;
  /** word/diagrams/dataN.xml keyed by rId */
  diagramData: Map<string, string>;
  /** word/media/* keyed by rId */
  media: Map<string, string>;
  /** Alternative format chunks (word/afchunkN.*) keyed by rId */
  afChunks: Map<string, string>;
  /** Sub-documents (word/subdocs/subdocN.docx) keyed by rId */
  subDocs: Map<string, string>;
}

export interface DocxDocument {
  doc: ParsedDocument;
  /** word/document.xml → w:body element */
  body: Element;
  /** word/styles.xml */
  styles?: Element;
  /** word/numbering.xml */
  numbering?: Element;
  /** word/settings.xml */
  settings?: Element;
  /** word/fontTable.xml */
  fontTable?: Element;
  /** word/webSettings.xml */
  webSettings?: Element;
  partRefs: DocxPartRefs;
  /** docProps/core.xml */
  coreProps?: string;
  /** docProps/app.xml */
  appProps?: string;
  /** docProps/custom.xml */
  customProps?: string;
}

function resolveRelsPath(target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  if (target.startsWith("../")) return target.replace("../", "");
  return `word/${target}`;
}

function parseDocPartRefs(doc: ParsedDocument): DocxPartRefs {
  const refs: DocxPartRefs = {
    headers: new Map(),
    footers: new Map(),
    charts: new Map(),
    diagramData: new Map(),
    media: new Map(),
    afChunks: new Map(),
    subDocs: new Map(),
  };

  const relsEl = doc.get("word/_rels/document.xml.rels");
  if (!relsEl) return refs;

  for (const child of relsEl.elements ?? []) {
    if (child.name !== "Relationship") continue;
    const type = attr(child, "Type") ?? "";
    const target = attr(child, "Target") ?? "";
    const id = attr(child, "Id") ?? "";
    if (!target) continue;

    const path = resolveRelsPath(target);

    if (type.includes("/header")) {
      refs.headers.set(id, path);
    } else if (type.includes("/footer")) {
      refs.footers.set(id, path);
    } else if (type.includes("/footnotes")) {
      refs.footnotes = path;
    } else if (type.includes("/endnotes")) {
      refs.endnotes = path;
    } else if (type.includes("/comments")) {
      refs.comments = path;
    } else if (type.includes("/chart")) {
      refs.charts.set(id, path);
    } else if (type.includes("/diagramData")) {
      refs.diagramData.set(id, path);
    } else if (type.includes("/image") || type.includes("/media")) {
      refs.media.set(id, path);
    } else if (type.includes("/aFChunk")) {
      refs.afChunks.set(id, path);
    } else if (type.includes("/subDocument")) {
      refs.subDocs.set(id, path);
    }
  }

  return refs;
}

function parseRootRels(doc: ParsedDocument): {
  coreProps?: string;
  appProps?: string;
  customProps?: string;
} {
  const relsEl = doc.get("_rels/.rels");
  if (!relsEl) return {};

  let coreProps: string | undefined;
  let appProps: string | undefined;
  let customProps: string | undefined;

  for (const child of relsEl.elements ?? []) {
    if (child.name !== "Relationship") continue;
    const type = attr(child, "Type") ?? "";
    const target = attr(child, "Target") ?? "";
    if (!target) continue;

    const path = target.startsWith("/") ? target.slice(1) : target;

    if (type.includes("/core-properties")) {
      coreProps = path;
    } else if (type.includes("/extended-properties")) {
      appProps = path;
    } else if (type.includes("/custom-properties")) {
      customProps = path;
    }
  }

  return { coreProps, appProps, customProps };
}

/**
 * Parse a .docx file and convert it into IPropertiesOptions.
 *
 * This is the main public API for parsing DOCX files.
 * The returned options can be passed directly to `new Document(parsed)`
 * to recreate the document.
 *
 * @param data - Raw bytes of a .docx file
 * @returns Document options including sections and metadata
 */
export function parseDocument(
  data: Uint8Array,
): import("@file/core-properties").IPropertiesOptions {
  const docx = parseDocx(data);
  const ctx = new ParseContext(docx, buildStyleCache(docx), buildNumberingCache(docx));
  const sections = parseBody(docx.body, ctx);

  const opts: Record<string, unknown> = { sections };

  // Core properties
  if (docx.coreProps) {
    const corePropsEl = docx.doc.get(docx.coreProps);
    if (corePropsEl) {
      const cp = parseCorePropsElement(corePropsEl);
      if (cp.title) opts.title = cp.title;
      if (cp.subject) opts.subject = cp.subject;
      if (cp.creator) opts.creator = cp.creator;
      if (cp.keywords) opts.keywords = cp.keywords;
      if (cp.description) opts.description = cp.description;
      if (cp.lastModifiedBy) opts.lastModifiedBy = cp.lastModifiedBy;
      if (cp.revision) opts.revision = parseInt(cp.revision, 10);
    }
  }

  // Settings
  if (docx.settings) {
    Object.assign(opts, parseSettings(docx.settings));
  }

  // Custom properties
  if (docx.customProps) {
    const customPropsEl = docx.doc.get(docx.customProps);
    if (customPropsEl) {
      const cpEntries = parseCustomProperties(customPropsEl);
      if (cpEntries.length > 0) {
        opts.customProperties = cpEntries;
      }
    }
  }

  return opts as unknown as import("@file/core-properties").IPropertiesOptions;
}

export function parseDocx(data: Uint8Array): DocxDocument {
  const doc = parseArchive(data);

  const documentEl = doc.get("word/document.xml");
  if (!documentEl) throw new Error("word/document.xml not found");
  const body = documentEl.elements?.find((e) => e.name === "w:body");
  if (!body) throw new Error("w:body not found in word/document.xml");

  const styles = doc.get("word/styles.xml");
  const numbering = doc.get("word/numbering.xml");
  const settings = doc.get("word/settings.xml");
  const fontTable = doc.get("word/fontTable.xml");
  const webSettings = doc.get("word/webSettings.xml");

  const partRefs = parseDocPartRefs(doc);
  const { coreProps, appProps, customProps } = parseRootRels(doc);

  return {
    doc,
    body,
    styles,
    numbering,
    settings,
    fontTable,
    webSettings,
    partRefs,
    coreProps,
    appProps,
    customProps,
  };
}
