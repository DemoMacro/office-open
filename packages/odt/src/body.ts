import type {
  ChartSpaceOptions,
  FormContainerOptions,
  SequenceDeclarationOptions,
  VariableDeclarationOptions,
} from "@office-open/core";
import type {
  BlockContentChild,
  DocumentOptions,
  ParagraphOptions,
  SectionChild,
} from "@office-open/docx";
import {
  attributeString,
  CHART_MIME,
  chartBodyXml,
  childNamed,
  childrenNamed,
  generateOcf,
  hasOcfManifestOverlay,
  metaXml,
  ODF_NAMESPACES,
  OdfSchemaError,
  parseEmbeddedCharts,
  parseGraphicStyles,
  parseMeta,
  readOcf,
  readXml,
  xmlElement,
  type GraphicStyle,
  type OdfPackageFiles,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

import {
  automaticStyleOverlaysXml,
  parseStyles,
  type OdtAutomaticStyleOverlay,
  type StyleMap,
} from "./automatic-styles";
import {
  documentStylesXml,
  fontFaceDecls,
  parseDefaultStyle,
  parseFontFace,
  parseCommonStyles,
} from "./common-styles";
import {
  parseShape,
  parseShapeHyperlink,
  shapeBlock,
  type OdtChart,
  type OdtImage,
} from "./drawing";
import { OdtParseError } from "./error";
import {
  listParagraphLevel,
  listXml,
  parseListNumberings,
  parseListStyles,
  parseOutlineStyle,
  type AbstractNumbering,
} from "./list";
import { parseMasterHeaderFooter, parsePageLayout } from "./master-pages";
import { parseNotesConfiguration } from "./notes";
import {
  bookmarkEndXml,
  declarationsXml,
  decodeStyleName,
  normalizeParagraph,
  paragraphXml,
  parseParagraph,
  parseSequenceDeclarations,
  parseTrackedChanges,
  parseVariableDeclarations,
  withPageBreakBefore,
  type ChangedProperties,
} from "./paragraph";
import {
  bibliographyIndexXml,
  encodeSectionTag,
  indexXml,
  nextBookmarkId,
  parseAnnotation,
  parseBibliography,
  parseForms,
  parseIndex,
  type NoteChildren,
  type NoteEntry,
  type NotesContext,
} from "./section";
import type { OdtDocumentOptions } from "./semantics";
import { parseTable, tableXml } from "./table";

export const MIME = "application/vnd.oasis.opendocument.text";

export const NAMESPACES = `${ODF_NAMESPACES} xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:xforms="http://www.w3.org/2002/xforms"`;

/** Shared lookup state threaded through the ODT parse pipeline. */
export interface ParseContext {
  styles: StyleMap;
  listStyles: Map<string, boolean>;
  graphicStyles: Map<string, GraphicStyle>;
  chartBodies: Map<string, ChartSpaceOptions>;
  objectMediaTypes: Map<string, string | undefined>;
  listDefinitions: AbstractNumbering[];
  outline?: AbstractNumbering;
  binaries: Record<string, Uint8Array>;
  notes: {
    footnotes: NoteEntry[];
    endnotes: NoteEntry[];
    forms: FormContainerOptions[];
  };
  bookmarkIds: Map<string, number>;
  changes: Map<
    string,
    ChangedProperties & { kind: "insertion" | "deletion"; children?: Element[] }
  >;
  declaredSequences: SequenceDeclarationOptions[];
  declaredVariables: VariableDeclarationOptions[];
  pendingSequences: Set<string>;
  pendingVariables: Map<string, "float" | "string">;
  trackRevisions?: boolean;
}

export function generateDocument(options: DocumentOptions): Uint8Array {
  const { packageManifest, styleOverlays } = options as OdtDocumentOptions;
  const styles: string[] = automaticStyleOverlaysXml(styleOverlays);
  const blocks = options.sections.flatMap((section) => section.children);
  const images: OdtImage[] = [];
  const charts: OdtChart[] = [];
  const notes = notesContext(options);
  collectBookmarkNames(blocks, notes.bookmarkNames);
  for (const note of [...(options.footnotes ?? []), ...(options.endnotes ?? [])])
    collectBookmarkNames(
      note.children.map((noteChild) =>
        typeof noteChild === "string" || !("paragraph" in noteChild || "table" in noteChild)
          ? ({ paragraph: noteChild } as SectionChild)
          : (noteChild as SectionChild),
      ),
      notes.bookmarkNames,
    );
  const bodyBlocks = blocksXml(blocks, styles, images, notes, options.numbering, charts);
  const body = declarationsXml(options, notes) + bodyBlocks;
  const files: OdfPackageFiles = {
    "content.xml": contentXml(body, styles, fontFaceDecls(options.fonts)),
    "styles.xml": documentStylesXml(
      options.sections[0],
      options.settings,
      options.numbering,
      options.styles,
    ),
    "meta.xml": metaXml(options),
  };
  for (const image of images) files[image.path] = image.data;
  for (const entry of charts) files[`${entry.path}/content.xml`] = chartBodyXml(entry.chart);
  return generateOcf(
    MIME,
    files,
    Object.fromEntries(charts.map((entry) => [`${entry.path}/`, CHART_MIME])),
    packageManifest,
  );
}

/** Note ids auto-assign 1, 2, … per class, matching the docx model. */
export function notesContext(options: DocumentOptions): NotesContext {
  const footnotes = new Map<number, NoteChildren>();
  let nextFootnoteId = 1;
  for (const note of options.footnotes ?? []) {
    const id = note.id ?? nextFootnoteId;
    footnotes.set(id, note.children);
    nextFootnoteId = id + 1;
  }
  const endnotes = new Map<number, NoteChildren>();
  let nextEndnoteId = 1;
  for (const note of options.endnotes ?? []) {
    const id = note.id ?? nextEndnoteId;
    endnotes.set(id, note.children);
    nextEndnoteId = id + 1;
  }
  return {
    footnotes,
    endnotes,
    bookmarkNames: new Map(),
    sequenceNames: new Set(
      (options.sequenceDeclarations ?? []).map((declaration) => declaration.name),
    ),
    variableTypes: new Map(
      (options.variableDeclarations ?? []).map((declaration) => [
        declaration.name,
        declaration.valueType,
      ]),
    ),
    forms: options.forms ?? [],
    revisions: new Map(),
  };
}

export function collectBookmarkNames(children: SectionChild[], names: Map<number, string>): void {
  for (const child of children) {
    if ("bookmarkStart" in child) names.set(child.bookmarkStart.id, child.bookmarkStart.name);
    if ("paragraph" in child && typeof child.paragraph !== "string")
      collectInlineBookmarkNames(child.paragraph.children ?? [], names);
    if ("table" in child)
      for (const row of child.table.rows)
        if ("cells" in row)
          for (const cell of row.cells)
            if ("children" in cell) collectBookmarkNames(cell.children, names);
    if ("sdt" in child) collectBookmarkNames(child.sdt.children ?? [], names);
    if ("textbox" in child) collectBookmarkNames(child.textbox.children ?? [], names);
    if ("toc" in child) collectBookmarkNames(child.toc.entries ?? [], names);
  }
}

export function collectInlineBookmarkNames(
  children: NonNullable<ParagraphOptions["children"]>,
  names: Map<number, string>,
): void {
  for (const child of children) {
    if (typeof child === "string") continue;
    if ("bookmarkStart" in child) names.set(child.bookmarkStart.id, child.bookmarkStart.name);
    if ("hyperlink" in child) collectInlineBookmarkNames(child.hyperlink.children ?? [], names);
    if ("insertion" in child) collectInlineBookmarkNames(child.insertion.children, names);
    if ("deletion" in child) collectInlineBookmarkNames(child.deletion.children, names);
  }
}

export function blocksXml(
  blocks: SectionChild[],
  styles: string[],
  images: OdtImage[],
  notes: NotesContext,
  numbering: DocumentOptions["numbering"],
  charts: OdtChart[],
): string {
  const parts: string[] = [];
  let index = 0;
  // Consecutive list paragraphs of the same level and kind group into one
  // text:list — the ODF shape for Word's bullet/numbering runs.
  while (index < blocks.length) {
    const listInfo = listParagraphLevel(blocks[index]!);
    if (listInfo === undefined) {
      parts.push(blockXml(blocks[index]!, styles, images, notes, numbering, charts));
      index += 1;
      continue;
    }
    const group: SectionChild[] = [];
    while (index < blocks.length) {
      const info = listParagraphLevel(blocks[index]!);
      if (
        info?.level !== listInfo.level ||
        info?.ordered !== listInfo.ordered ||
        info?.reference !== listInfo.reference
      )
        break;
      group.push(blocks[index]!);
      index += 1;
    }
    parts.push(listXml(group, listInfo, styles, images, notes, numbering, charts));
  }
  return parts.join("");
}

export function parseDocument(data: Uint8Array): DocumentOptions {
  try {
    return parseOdtBody(data);
  } catch (cause) {
    if (cause instanceof OdtParseError) throw cause;
    if (cause instanceof OdfSchemaError) {
      throw new OdtParseError(cause.message, cause.part, cause.path, cause.name, cause.reason, {
        cause,
      });
    }
    throw new OdtParseError(
      cause instanceof Error ? cause.message : "Unable to parse ODT package",
      "mimetype",
      "/mimetype",
      "mimetype",
      "invalid ODT package",
      { cause },
    );
  }
}

function parseOdtBody(data: Uint8Array): OdtDocumentOptions {
  const { files, binaries, manifest } = readOcf(data, MIME);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:text");
  const styleContainer = childNamed(content, "office:automatic-styles");
  const styleMap = parseStyles(styleContainer);
  const graphicStyles = parseGraphicStyles(styleContainer);
  const chartBodies = parseEmbeddedCharts(manifest, files);
  const objectMediaTypes = new Map(
    manifest.entries
      .filter((entry) => entry.fullPath.endsWith("/"))
      .map((entry) => [entry.fullPath.replace(/\/$/, ""), entry.mediaType]),
  );
  const context: ParseContext = {
    styles: styleMap,
    listStyles: parseListStyles(styleContainer),
    graphicStyles,
    chartBodies,
    objectMediaTypes,
    listDefinitions: parseListNumberings(styleContainer),
    outline: parseOutlineStyle(files),
    binaries,
    notes: { footnotes: [], endnotes: [], forms: [] },
    bookmarkIds: new Map(),
    changes: new Map(),
    declaredSequences: [],
    declaredVariables: [],
    pendingSequences: new Set(),
    pendingVariables: new Map(),
  };
  const children = parseBlocks(body?.elements ?? [], context);
  const meta = Object.fromEntries(
    Object.entries(parseMeta(files)).filter(([, value]) => value !== undefined),
  );
  const result: OdtDocumentOptions = {
    ...meta,
    sections: [{ properties: parsePageLayout(files), children }],
  };
  const styleOverlays: OdtAutomaticStyleOverlay[] = [...styleMap]
    .filter(([, style]) =>
      (style.properties ?? []).some((property) => Object.keys(property.attributes).length > 0),
    )
    .map(([name, style]) => ({
      name,
      family: style.family,
      properties: (style.properties ?? []).filter(
        (property) => Object.keys(property.attributes).length > 0,
      ),
    }));
  if (styleOverlays.length > 0) result.styleOverlays = styleOverlays;
  if (hasOcfManifestOverlay(manifest)) result.packageManifest = manifest;
  if (context.notes.forms.length > 0) result.forms = context.notes.forms;
  if (context.declaredSequences.length > 0) result.sequenceDeclarations = context.declaredSequences;
  if (context.declaredVariables.length > 0) result.variableDeclarations = context.declaredVariables;
  if (context.notes.footnotes.length > 0) result.footnotes = context.notes.footnotes;
  if (context.notes.endnotes.length > 0) result.endnotes = context.notes.endnotes;
  const fonts = childrenNamed(childNamed(content, "office:font-face-decls"), "style:font-face").map(
    parseFontFace,
  );
  if (fonts.length > 0) result.fonts = fonts;
  const numberings = context.outline
    ? [context.outline, ...context.listDefinitions]
    : context.listDefinitions;
  if (numberings.length > 0) result.numbering = { abstractNumberings: numberings };
  if (context.trackRevisions !== undefined)
    result.settings = { ...result.settings, trackRevisions: context.trackRevisions };
  const notesConfiguration = parseNotesConfiguration(files);
  if (notesConfiguration.footnoteProperties || notesConfiguration.endnoteProperties)
    result.settings = { ...result.settings, ...notesConfiguration };
  const defaultStyle = parseDefaultStyle(files);
  const commonStyles = parseCommonStyles(files);
  if (defaultStyle || commonStyles)
    result.styles = {
      ...commonStyles,
      ...result.styles,
      default: { ...result.styles?.default, document: defaultStyle },
    };
  const masterHeaderFooter = parseMasterHeaderFooter(files);
  const section = result.sections[0];
  if (section) {
    if (masterHeaderFooter.headers) section.headers = masterHeaderFooter.headers;
    if (masterHeaderFooter.footers) section.footers = masterHeaderFooter.footers;
  }
  return result;
}

export function contentXml(body: string, styles: string[], fontFaces: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3">${fontFaces}<office:automatic-styles>${styles.join(
    "",
  )}</office:automatic-styles><office:body><office:text>${body}</office:text></office:body></office:document-content>`;
}

export function blockXml(
  child: SectionChild,
  styles: string[],
  images: OdtImage[],
  notes: NotesContext,
  numbering: DocumentOptions["numbering"],
  charts: OdtChart[],
): string {
  if ("toc" in child) return indexXml(child.toc);
  if ("sdt" in child && child.sdt.properties.bibliography) return bibliographyIndexXml(child.sdt);
  if ("paragraph" in child)
    return paragraphXml(
      normalizeParagraph(child.paragraph),
      styles,
      images,
      notes,
      numbering,
      charts,
    );
  if ("bookmarkStart" in child)
    return xmlElement("text:bookmark-start", {
      "xml:id": child.bookmarkStart.id ? `bookmark-${child.bookmarkStart.id}` : undefined,
      "text:name": child.bookmarkStart.name,
    });
  if ("bookmarkEnd" in child) return bookmarkEndXml(child.bookmarkEnd.id, notes);
  if ("table" in child)
    return tableXml(child.table, styles, (block) =>
      blockXml(block, styles, images, notes, numbering, charts),
    );
  if ("sdt" in child && !child.sdt.properties.bibliography) {
    const styleName = /^odf:text-section(?:;style=([\s\S]*))?$/.exec(
      child.sdt.properties.tag ?? "",
    )?.[1];
    return xmlElement(
      "text:section",
      {
        "text:name": child.sdt.properties.alias ?? "",
        "text:style-name": styleName ? decodeStyleName(styleName) : undefined,
        "text:protected": child.sdt.properties.lock === "sdtLocked" ? true : undefined,
      },
      [blocksXml(child.sdt.children ?? [], styles, images, notes, numbering, charts)],
    );
  }
  return "";
}

export function parseBlock(element: Element, context: ParseContext): SectionChild {
  if (element.name === "table:table")
    return parseTable(
      element,
      (name) => context.styles.get(name)?.columnWidth,
      (child) => parseBlock(child, context),
    );
  const paragraph = parseParagraph(element, context);
  return { paragraph };
}

export function parseBlocks(
  elements: Element[],
  context: ParseContext,
  listDepth = 0,
  listState?: { name: string; ordered: boolean },
): SectionChild[] {
  const result: SectionChild[] = [];
  let pendingPageBreak = false;
  for (const element of elements) {
    if (element.name === "text:section") {
      const styleName = attributeString(element, "text:style-name");
      result.push({
        sdt: {
          properties: {
            alias: attributeString(element, "text:name") ?? "",
            ...(styleName ? { tag: encodeSectionTag(styleName) } : {}),
            ...(attributeString(element, "text:protected") === "true"
              ? { lock: "sdtLocked" as const }
              : {}),
          },
          children: blockChildren(
            parseBlocks(element.elements ?? [], context, listDepth, listState),
          ),
        },
      });
      continue;
    }
    if (element.name === "text:list") {
      // A typed text:list unwraps to bullet or numbered paragraphs at the
      // nesting depth, judged by its referenced list style; list-header
      // content (rare) keeps the generic fall-through below.
      const name = attributeString(element, "text:style-name") ?? "";
      const state = { name, ordered: context.listStyles.get(name) ?? false };
      for (const item of childrenNamed(element, "text:list-item")) {
        result.push(...parseBlocks(item.elements ?? [], context, listDepth + 1, state));
      }
      continue;
    }
    if (element.name === "text:soft-page-break") {
      pendingPageBreak = true;
      continue;
    }
    if (element.name === "text:bookmark-start") {
      const name = attributeString(element, "text:name") ?? "";
      if (!name || context.bookmarkIds.has(name)) throw unknownOdtElement(element);
      const encodedId = /^(?:bookmark-)?(\d+)$/.exec(
        attributeString(element, "text:id") ?? attributeString(element, "xml:id") ?? "",
      )?.[1];
      const id = encodedId ? Number(encodedId) : nextBookmarkId(context);
      context.bookmarkIds.set(name, id);
      result.push({
        bookmarkStart: { id, name },
      });
      continue;
    }
    if (element.name === "text:bookmark-end") {
      const name = attributeString(element, "text:name") ?? "";
      if (!name || !context.bookmarkIds.has(name)) throw unknownOdtElement(element);
      result.push({
        bookmarkEnd: { id: context.bookmarkIds.get(name) ?? nextBookmarkId(context) },
      });
      continue;
    }
    if (element.name === "text:reference-mark-start") {
      const name = attributeString(element, "text:name") ?? "";
      const id = context.bookmarkIds.get(name) ?? nextBookmarkId(context);
      context.bookmarkIds.set(name, id);
      result.push({
        bookmarkStart: { id, name },
      });
      continue;
    }
    if (element.name === "office:annotation") {
      result.push({
        paragraph: {
          children: [parseAnnotation(element)],
        },
      });
      continue;
    }
    if (element.name === "text:illustration-index") {
      result.push({ toc: parseIndex(element) });
      continue;
    }
    if (element.name === "text:bibliography") {
      result.push({
        sdt: parseBibliography(element),
      });
      continue;
    }
    if (element.name === "draw:ellipse") {
      result.push(shapeBlock(parseShape(element)));
      continue;
    }
    if (element.name === "draw:rect") {
      result.push(shapeBlock(parseShape(element)));
      continue;
    }
    if (element.name === "draw:a") {
      result.push({ paragraph: { children: [parseShapeHyperlink(element)] } });
      continue;
    }
    if (element.name === "office:forms") {
      result.push(...parseForms(element, context));
      continue;
    }
    if (element.name === "text:sequence-decls") {
      parseSequenceDeclarations(element, context);
      continue;
    }
    if (element.name === "text:variable-decls") {
      parseVariableDeclarations(element, context);
      continue;
    }
    if (element.name === "text:tracked-changes") {
      parseTrackedChanges(element, context);
      continue;
    }
    if (
      [
        "office:forms",
        "text:sequence-decls",
        "text:variable-decls",
        "text:tracked-changes",
      ].includes(element.name ?? "")
    ) {
      throw unknownOdtElement(element);
    }
    if (element.name === "text:p" || element.name === "text:h" || element.name === "table:table") {
      const child = parseBlock(element, context);
      if (listDepth > 0 && "paragraph" in child) {
        const paragraph = normalizeParagraph(child.paragraph);
        const mappedChild = {
          paragraph: listState?.ordered
            ? { ...paragraph, numbering: { reference: listState.name, level: listDepth - 1 } }
            : { ...paragraph, bullet: { level: listDepth - 1 } },
        } as SectionChild;
        result.push(pendingPageBreak ? withPageBreakBefore(mappedChild) : mappedChild);
        pendingPageBreak = false;
      } else {
        result.push(pendingPageBreak ? withPageBreakBefore(child) : child);
        pendingPageBreak = false;
      }
    } else if (element.name) {
      throw unknownOdtElement(element);
    }
  }
  if (pendingPageBreak) {
    throw new OdtParseError(
      "content.xml: /office:document-content/office:body/office:text/text:soft-page-break: no following block receives the page break",
      "content.xml",
      "/office:document-content/office:body/office:text/text:soft-page-break",
      "text:soft-page-break",
      "no following block receives the page break",
    );
  }
  return result;
}

export function blockChildren(children: SectionChild[]): BlockContentChild[] {
  return children.filter((child): child is BlockContentChild => !("altChunk" in child));
}

export function unknownOdtElement(element: Element): OdtParseError {
  const name = element.name ?? "";
  const path = `/office:document-content/office:body/office:text/${name}`;
  return new OdtParseError(
    `content.xml: ${path}: ${name}: no canonical DocumentOptions mapping`,
    "content.xml",
    path,
    name,
    "no canonical DocumentOptions mapping",
  );
}

export function assertElementNames(
  element: Element,
  allowedAttributes: string[],
  allowedChildren: string[] = [],
): void {
  const unknown = Object.keys(element.attributes ?? {}).filter(
    (name) => !allowedAttributes.includes(name),
  );
  if (unknown.length) {
    throw new OdtParseError(
      `content.xml: unknown ${element.name} attribute: ${unknown.join(", ")}`,
      "content.xml",
      `/office:document-content/office:body/office:text/${element.name}`,
      element.name,
      "attribute has no canonical mapping",
    );
  }
  for (const child of element.elements ?? []) {
    if (child.type === "element" && !allowedChildren.includes(child.name ?? "")) {
      throw new OdtParseError(
        `content.xml: unknown ${element.name} child: ${child.name}`,
        "content.xml",
        `/office:document-content/office:body/office:text/${element.name}`,
        child.name,
        "child has no canonical mapping",
      );
    }
  }
}

export function malformedOdtElement(name: string, reason: string): OdtParseError {
  const path = `/office:document-content/office:body/office:text/${name}`;
  return new OdtParseError(`content.xml: ${name}: ${reason}`, "content.xml", path, name, reason);
}
