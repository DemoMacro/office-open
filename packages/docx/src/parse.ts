import type { ParsedArchive } from "@office-open/core";
import { parseArchive } from "@office-open/core";
import type { DataType } from "@office-open/core";
import { toUint8Array } from "@office-open/core";
import { attr } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { appPropertiesDesc } from "@parts/app-properties";
import { bibliographyDesc } from "@parts/bibliography";
import { setBodyParseChild } from "@parts/bodychildren";
import { commentsDesc } from "@parts/comments";
import { contentTypesDesc } from "@parts/contenttypes";
import { corePropertiesDesc } from "@parts/core-properties";
import type { DocumentOptions } from "@parts/core-properties";
import { customPropertiesDesc } from "@parts/custom-properties";
import { endnotesDesc } from "@parts/endnotes/descriptor";
import { fontTableDesc } from "@parts/fonts/descriptor";
import type { EmbeddedFontOptionsWithKey } from "@parts/fonts/font-wrapper";
import { footnotesDesc } from "@parts/footnotes/descriptor";
import { glossaryDesc } from "@parts/glossary-document";
import { parseNumberingDefinitions } from "@parts/numbering/numbering";
import { settingsDesc } from "@parts/settings/descriptor";
import { buildStyleCache, buildNumberingCache, parseStyleDefinitions } from "@parts/styles/styles";
import { setTableParseChild } from "@parts/table/descriptor";
import { webSettingsDesc } from "@parts/web-settings";

import { parseParagraphProperties } from "./body";
import { DocxReadContext } from "./context";
import { parseBody, parseSectionChild } from "./parse/body";
import { replaceRelsWithPlaceholders } from "./util/replace-media-placeholders";
import { stringifyElement } from "./util/stringify-element";

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
  /** Hyperlink targets keyed by rId (external URLs) */
  hyperlinks: Map<string, string>;
  /** word/charts/chartN.xml keyed by rId */
  charts: Map<string, string>;
  /** word/diagrams/dataN.xml keyed by rId */
  diagramData: Map<string, string>;
  /** word/media/* keyed by rId (from document.xml.rels) */
  media: Map<string, string>;
  /**
   * Per-part image/media relationships. Each part (document, headers, footers,
   * footnotes, …) has its own .rels with independent rId numbering, so drawings
   * inside a part must resolve images against that part's rels. Maps
   * partPath → (rId → mediaPath).
   */
  partMedia: Map<string, Map<string, string>>;
  /** Alternative format chunks (word/afchunkN.*) keyed by rId */
  afChunks: Map<string, string>;
  /** Sub-documents (word/subdocs/subdocN.docx) keyed by rId */
  subDocs: Map<string, string>;
  /** word/bibliography.xml */
  bibliography?: string;
  /** word/glossary/document.xml */
  glossary?: string;
}

export interface DocxDocument {
  doc: ParsedArchive;
  /** word/document.xml → w:body element */
  body: Element;
  /** word/document.xml → w:background element */
  background?: Element;
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
  /** [Content_Types].xml */
  contentTypes?: Element;
}

function resolveRelsPath(target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  if (target.startsWith("../")) return target.replace("../", "");
  return `word/${target}`;
}

/**
 * Resolve each embedded font's .odttf bytes through fontTable.xml.rels.
 * Reads the binary verbatim and flags it raw so the compiler copies it as-is
 * instead of re-obfuscating (the fontKey already matches the bytes).
 */
function resolveEmbeddedFontData(fonts: EmbeddedFontOptionsWithKey[], doc: ParsedArchive): void {
  const relsEl = doc.get("word/_rels/fontTable.xml.rels");
  if (!relsEl) return;
  const ridToPath = new Map<string, string>();
  for (const child of relsEl.elements ?? []) {
    if (child.name !== "Relationship") continue;
    const type = attr(child, "Type") ?? "";
    if (!type.includes("/font")) continue;
    const id = attr(child, "Id") ?? "";
    const target = attr(child, "Target") ?? "";
    if (id && target) ridToPath.set(id, resolveRelsPath(target));
  }
  for (const font of fonts) {
    if (!font.embedRid) continue;
    const odttfPath = ridToPath.get(font.embedRid);
    if (!odttfPath) continue;
    const bytes = doc.getRaw(odttfPath);
    if (bytes) {
      font.data = Buffer.from(bytes);
      font.rawOdttf = true;
      font.odttfPath = odttfPath;
    }
  }
}

function parseDocPartRefs(doc: ParsedArchive): DocxPartRefs {
  const refs: DocxPartRefs = {
    headers: new Map(),
    footers: new Map(),
    hyperlinks: new Map(),
    charts: new Map(),
    diagramData: new Map(),
    media: new Map(),
    partMedia: new Map(),
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
    } else if (type.includes("/bibliography")) {
      refs.bibliography = path;
    } else if (type.includes("/glossaryDocument")) {
      refs.glossary = path;
    } else if (type.includes("/hyperlink")) {
      refs.hyperlinks.set(id, target);
    }
  }

  // Per-part image relationships. Each part carries its own .rels with
  // independent rId numbering (document rId1 ≠ header rId1), so collect them
  // keyed by part path; drawings inside a part resolve images through its
  // own rels. Covers document, headers, footers, footnotes, endnotes, comments.
  for (const relsPath of doc.keys("word/_rels/")) {
    if (!relsPath.endsWith(".rels")) continue;
    const relsEl = doc.get(relsPath);
    if (!relsEl) continue;
    const partPath = "word/" + relsPath.slice("word/_rels/".length, -".rels".length);
    for (const rel of relsEl.elements ?? []) {
      if (rel.name !== "Relationship") continue;
      const type = attr(rel, "Type") ?? "";
      if (!type.includes("/image") && !type.includes("/media")) continue;
      const id = attr(rel, "Id") ?? "";
      const target = attr(rel, "Target") ?? "";
      if (!id || !target) continue;
      let partMap = refs.partMedia.get(partPath);
      if (!partMap) {
        partMap = new Map();
        refs.partMedia.set(partPath, partMap);
      }
      partMap.set(id, resolveRelsPath(target));
    }
  }

  return refs;
}

function parseRootRels(doc: ParsedArchive): {
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
 * Parse a .docx file and convert it into DocumentOptions.
 *
 * This is the main public API for parsing DOCX files.
 * The returned options can be passed directly to `new Document(parsed)`
 * to recreate the document.
 *
 * @param data - Raw bytes of a .docx file
 * @returns Document options including sections and metadata
 */
export function parseDocument(data: DataType): DocumentOptions {
  const docx = parseDocx(data);
  const ctx = new DocxReadContext(
    docx,
    buildStyleCache(docx.styles),
    buildNumberingCache(docx.numbering),
  );

  // Register the child parser for table and body child descriptors
  setTableParseChild(parseSectionChild);
  setBodyParseChild(parseSectionChild);

  const sections = parseBody(docx.body, ctx);

  const opts: Record<string, unknown> = { sections };

  // Background (w:background in document.xml)
  if (docx.background) {
    const hasChildren = (docx.background.elements ?? []).some((e) => e.type === "element");
    if (hasChildren) {
      // VML/structured background (e.g. v:background/v:fill pattern with a
      // texture image) that doesn't fit the color/theme model: carry the
      // element verbatim, rewriting relationship refs to {fileName} placeholders
      // so the media round-trips via the compiler's placeholder pass.
      const { rawXml, rawMedia } = replaceRelsWithPlaceholders(
        stringifyElement(docx.background),
        ctx,
        "background",
      );
      opts.background = rawMedia.length > 0 ? { rawXml, rawMedia } : { rawXml };
    } else {
      const bg: Record<string, unknown> = {};
      const color = attr(docx.background, "w:color");
      if (color) bg.color = color;
      const themeColor = attr(docx.background, "w:themeColor");
      if (themeColor) bg.themeColor = themeColor;
      const themeShade = attr(docx.background, "w:themeShade");
      if (themeShade) bg.themeShade = themeShade;
      const themeTint = attr(docx.background, "w:themeTint");
      if (themeTint) bg.themeTint = themeTint;
      if (Object.keys(bg).length > 0) opts.background = bg;
    }
  }

  // Core properties
  if (docx.coreProps) {
    const corePropsEl = docx.doc.get(docx.coreProps);
    if (corePropsEl) {
      const cp = corePropertiesDesc.parse(corePropsEl, ctx);
      if (cp.title) opts.title = cp.title;
      if (cp.subject) opts.subject = cp.subject;
      if (cp.creator) opts.creator = cp.creator;
      if (cp.keywords) opts.keywords = cp.keywords;
      if (cp.description) opts.description = cp.description;
      if (cp.lastModifiedBy) opts.lastModifiedBy = cp.lastModifiedBy;
      if (cp.revision) opts.revision = cp.revision;
      if (cp.lastPrinted) opts.lastPrinted = cp.lastPrinted;
      if (cp.created) opts.created = cp.created;
      if (cp.modified) opts.modified = cp.modified;
    }
  }

  // App (extended) properties
  if (docx.appProps) {
    const appPropsEl = docx.doc.get(docx.appProps);
    if (appPropsEl) {
      const ap = appPropertiesDesc.parse(appPropsEl, ctx);
      if (Object.keys(ap).length > 0) opts.appProperties = ap;
    }
  }

  // Settings
  if (docx.settings) {
    const settingsOpts = settingsDesc.parse(docx.settings, ctx);
    Object.assign(opts, settingsOpts);
    // Surface the verbatim capture (rootAttributes + rawXml) onto
    // DocumentOptions.settings so context.ts spreads it into _settingsOptions
    // (the descriptor's stringify input).
    if (settingsOpts.rawXml !== undefined) {
      opts.settings = {
        rawXml: settingsOpts.rawXml,
        rootAttributes: settingsOpts.rootAttributes,
      };
    }
  }

  // Web settings
  if (docx.webSettings) {
    const wsOpts = webSettingsDesc.parse(docx.webSettings, ctx);
    if (Object.keys(wsOpts).length > 0) opts.webSettings = wsOpts;
  }

  // Custom properties
  if (docx.customProps) {
    const customPropsEl = docx.doc.get(docx.customProps);
    if (customPropsEl) {
      const cpResult = customPropertiesDesc.parse(customPropsEl, ctx);
      if (cpResult.properties && cpResult.properties.length > 0) {
        opts.customProperties = cpResult.properties;
      }
    }
  }

  // Comments content
  if (docx.partRefs.comments) {
    const commentsEl = docx.doc.get(docx.partRefs.comments);
    if (commentsEl) {
      const commentsResult = ctx.withPart(docx.partRefs.comments, () =>
        commentsDesc.parse(commentsEl, ctx),
      );
      const children = commentsResult.children;
      if (children && children.length > 0) {
        opts.comments = { children } as unknown as DocumentOptions["comments"];
      }
    }
  }

  // Footnotes content
  if (docx.partRefs.footnotes) {
    const footnotesEl = docx.doc.get(docx.partRefs.footnotes);
    if (footnotesEl) {
      const fnResult = ctx.withPart(docx.partRefs.footnotes, () =>
        footnotesDesc.parse(footnotesEl, ctx),
      );
      const footnotesMap: Record<string, { children: unknown[] }> = {};
      for (const [id, paragraphs] of fnResult.notes) {
        footnotesMap[String(id)] = { children: paragraphs };
      }
      // Preserve round-tripped separators so the generated ids stay consistent
      // with settings.footnotePr (which references them).
      if (
        Object.keys(footnotesMap).length > 0 ||
        fnResult.separator ||
        fnResult.continuationSeparator
      ) {
        const fnOpts = footnotesMap as NonNullable<DocumentOptions["footnotes"]>;
        if (fnResult.separator) fnOpts.separator = fnResult.separator;
        if (fnResult.continuationSeparator)
          fnOpts.continuationSeparator = fnResult.continuationSeparator;
        opts.footnotes = fnOpts;
      }
    }
  }

  // Endnotes content
  if (docx.partRefs.endnotes) {
    const endnotesEl = docx.doc.get(docx.partRefs.endnotes);
    if (endnotesEl) {
      const enResult = ctx.withPart(docx.partRefs.endnotes, () =>
        endnotesDesc.parse(endnotesEl, ctx),
      );
      const endnotesMap: Record<string, { children: unknown[] }> = {};
      for (const [id, paragraphs] of enResult.notes) {
        endnotesMap[String(id)] = { children: paragraphs };
      }
      if (
        Object.keys(endnotesMap).length > 0 ||
        enResult.separator ||
        enResult.continuationSeparator
      ) {
        const enOpts = endnotesMap as NonNullable<DocumentOptions["endnotes"]>;
        if (enResult.separator) enOpts.separator = enResult.separator;
        if (enResult.continuationSeparator)
          enOpts.continuationSeparator = enResult.continuationSeparator;
        opts.endnotes = enOpts;
      }
    }
  }

  // Styles definitions
  if (docx.styles) {
    const styleOpts = parseStyleDefinitions(docx.styles, parseParagraphProperties, ctx);
    if (styleOpts) opts.styles = styleOpts;
  }

  // Numbering definitions
  if (docx.numbering) {
    const numOpts = parseNumberingDefinitions(docx.numbering, parseParagraphProperties, ctx);
    if (numOpts) opts.numbering = numOpts;
  }

  // Font table
  if (docx.fontTable) {
    const ftResult = fontTableDesc.parse(docx.fontTable, ctx);
    if (ftResult.fonts && ftResult.fonts.length > 0) {
      resolveEmbeddedFontData(ftResult.fonts, docx.doc);
      opts.fonts = ftResult.fonts;
    }
  }

  // Bibliography
  if (docx.partRefs.bibliography) {
    const bibEl = docx.doc.get(docx.partRefs.bibliography);
    if (bibEl) {
      const bibResult = bibliographyDesc.parse(bibEl, ctx);
      if (bibResult.sources && bibResult.sources.length > 0) opts.bibliography = bibResult;
    }
  }

  // Glossary document
  if (docx.partRefs.glossary) {
    const glossaryEl = docx.doc.get(docx.partRefs.glossary);
    if (glossaryEl) {
      const glossaryResult = ctx.withPart(docx.partRefs.glossary, () =>
        glossaryDesc.parse(glossaryEl, ctx),
      );
      if (glossaryResult.parts && glossaryResult.parts.length > 0) opts.glossary = glossaryResult;
    }
  }

  // Content types
  if (docx.contentTypes) {
    const ctResult = contentTypesDesc.parse(docx.contentTypes, ctx);
    if (ctResult) opts.contentTypes = ctResult;
  }

  // Raw passthrough: parts generate() doesn't rebuild (word/theme/*, customXml/*).
  // Carried verbatim so their [Content_Types] declarations stay valid and the
  // package opens in Word. (Media/fonts/headers/etc. are rebuilt by the compiler
  // and must NOT be passed through — they'd otherwise duplicate under renamed paths.)
  const rawParts: { path: string; data: Uint8Array }[] = [];
  for (const prefix of ["word/theme/", "customXml/"]) {
    for (const p of docx.doc.keys(prefix)) {
      if (p.endsWith("/")) continue;
      const data = docx.doc.getRaw(p);
      if (data) rawParts.push({ path: p, data });
    }
  }
  if (rawParts.length > 0) opts.rawParts = rawParts;

  return opts as unknown as DocumentOptions;
}

export function parseDocx(data: DataType): DocxDocument {
  const uint8 = toUint8Array(data);
  const doc = parseArchive(uint8);

  const documentEl = doc.get("word/document.xml");
  if (!documentEl) throw new Error("word/document.xml not found");
  const body = documentEl.elements?.find((e) => e.name === "w:body");
  if (!body) throw new Error("w:body not found in word/document.xml");
  const background = documentEl.elements?.find((e) => e.name === "w:background");

  const styles = doc.get("word/styles.xml");
  const numbering = doc.get("word/numbering.xml");
  const settings = doc.get("word/settings.xml");
  const fontTable = doc.get("word/fontTable.xml");
  const webSettings = doc.get("word/webSettings.xml");

  const partRefs = parseDocPartRefs(doc);
  const { coreProps, appProps, customProps } = parseRootRels(doc);

  const contentTypes = doc.get("[Content_Types].xml");

  return {
    doc,
    body,
    background,
    styles,
    numbering,
    settings,
    fontTable,
    webSettings,
    partRefs,
    coreProps,
    appProps,
    customProps,
    contentTypes,
  };
}
