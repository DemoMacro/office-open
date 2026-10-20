import {
  convertOutput,
  toUint8Array,
  toUint8ArrayAsync,
  streamUint8Array,
  type DataType,
  type OutputByType,
  type OutputType,
} from "@office-open/core";
import type {
  ChartSpaceOptions,
  OfficeFormsOptions,
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
  chartSubdocumentFiles,
  childNamed,
  childrenNamed,
  descendantElements,
  generateOcf,
  hasOcfManifestOptions,
  metaXml,
  ODF_NAMESPACES,
  OdfSchemaError,
  parseOfficeForms,
  parseEmbeddedCharts,
  parseGraphicStyles,
  parseMeta,
  readOcfSync,
  readXml,
  unknownAttributeNames,
  unknownChildElement,
  xmlElement,
  type GraphicStyle,
  type OcfManifestOptions,
  type OdfPackageFiles,
  type PackageMemberOptions,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

import {
  automaticStyleResourcesXml,
  parseStyles,
  type OdtAutomaticStyleResource,
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
  isShapeElement,
  parseChartFrame,
  parseGroup,
  parsePictureFrame,
  parseShape,
  parseShapeHyperlink,
  parseTextBoxFrame,
  shapeBlock,
  type OdtChart,
  type OdtImage,
} from "./drawing";
import type { OdtEmbeddedObjectOptions } from "./embedded-objects";
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
import type { OdtDocumentOptions } from "./options";
import {
  bookmarkEndXml,
  declarationsXml,
  decodeStyleName,
  normalizeParagraph,
  paragraphXml,
  formatRevisionBlocks,
  parseParagraph,
  parseRuns,
  parseSequenceDeclarations,
  parseTrackedChanges,
  parseVariableDeclarations,
  parseUserFieldDeclarations,
  parseDdeConnectionDeclarations,
  trackRevisionBlocks,
  withPageBreakBefore,
  type ParsedTrackChange,
} from "./paragraph";
import {
  bibliographyIndexXml,
  bookmarkDisplayName,
  canonicalBookmarkName,
  countBookmarkName,
  encodeSectionTag,
  indexXml,
  nextBookmarkId,
  parseAnnotationEnd,
  parseAnnotationStart,
  parseBibliography,
  parseContentIndex,
  parseIndex,
  type NoteChildren,
  type NoteEntry,
  type NotesContext,
} from "./section";
import { parseTable, tableXml } from "./table";

export const MIME = "application/vnd.oasis.opendocument.text";
export const TEMPLATE_MIME = "application/vnd.oasis.opendocument.text-template";

export const NAMESPACES = `${ODF_NAMESPACES} xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:xforms="http://www.w3.org/2002/xforms"`;

/** Shared lookup state threaded through the ODT parse pipeline. */
export interface ParseContext {
  styles: StyleMap;
  listStyles: Map<string, boolean>;
  graphicStyles: Map<string, GraphicStyle>;
  chartBodies: Map<string, ChartSpaceOptions>;
  embeddedObjects: Map<string, OdtEmbeddedObjectOptions>;
  objectMediaTypes: Map<string, string | undefined>;
  listDefinitions: AbstractNumbering[];
  outline?: AbstractNumbering;
  binaries: Record<string, Uint8Array>;
  usedBinaryPaths: Set<string>;
  notes: {
    footnotes: NoteEntry[];
    endnotes: NoteEntry[];
    forms: OfficeFormsOptions;
  };
  comments: NonNullable<DocumentOptions["comments"]>;
  annotationIds: Map<string, number>;
  bookmarkIds: Map<string, number>;
  indexAutoMarkFile?: import("./options").OdtIndexAutoMarkFileOptions;
  pendingFormatChanges: Map<string, import("./paragraph").ChangedProperties>;
  sectionSources: Map<string, import("./options").OdtSectionSourceOptions>;
  changes: Map<string, ParsedTrackChange>;
  changeAlternatives: Map<string, ParsedTrackChange>;
  pendingInsertions: Set<string>;
  declaredSequences: SequenceDeclarationOptions[];
  declaredVariables: VariableDeclarationOptions[];
  documentVariables: NonNullable<NonNullable<DocumentOptions["settings"]>["docVars"]>;
  pendingSequences: Set<string>;
  pendingVariables: Map<string, "float" | "string">;
  nextEmbeddedShapeId: number;
  trackRevisions?: boolean;
}

function generateOdtPackage(options: OdtDocumentOptions): Uint8Array {
  const { packageManifest, embeddedObjects, packageMembers, styleResources } = options;
  const styles: string[] = automaticStyleResourcesXml(styleResources);
  const blocks = options.sections.flatMap((section) => section.children);
  const images: OdtImage[] = [];
  const charts: OdtChart[] = [];
  const notes = notesContext(options);
  collectBookmarkNames(blocks, notes);
  for (const note of [...(options.footnotes ?? []), ...(options.endnotes ?? [])])
    collectBookmarkNames(
      note.children.map((noteChild) =>
        typeof noteChild === "string" || !("paragraph" in noteChild || "table" in noteChild)
          ? ({ paragraph: noteChild } as SectionChild)
          : (noteChild as SectionChild),
      ),
      notes,
    );
  const bodyBlocks = blocksXml(
    blocks,
    styles,
    images,
    notes,
    options.numbering,
    charts,
    embeddedObjects ?? [],
  );
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
  for (const entry of charts)
    Object.assign(files, chartSubdocumentFiles(`${entry.path}/`, entry.chart));
  const mediaTypes: Record<string, string> = Object.fromEntries(
    charts.map((entry) => [`${entry.path}/`, CHART_MIME]),
  );
  for (const object of embeddedObjects ?? []) {
    validateObjectPath(object.path);
    mediaTypes[`${object.path}/`] = object.mediaType;
  }
  for (const member of packageMembers ?? []) {
    addPackageFile(files, member.path, member.data);
    if (member.mediaType !== undefined) mediaTypes[member.path] = member.mediaType;
  }
  for (const object of embeddedObjects ?? []) {
    for (const member of object.members) {
      if (!member.path.startsWith(`${object.path}/`))
        throw new Error(`Embedded object member is outside ${object.path}: ${member.path}`);
      addPackageFile(files, member.path, member.data);
      if (member.mediaType !== undefined) mediaTypes[member.path] = member.mediaType;
    }
  }
  return generateOcf(MIME, files, mediaTypes, packageManifest);
}

export function generateDocumentSync<T extends OutputType = "nodebuffer">(
  options: OdtDocumentOptions,
  output?: T,
): OutputByType[T] {
  return convertOutput(generateOdtPackage(options), (output ?? "nodebuffer") as T, MIME);
}

export async function generateDocument<T extends OutputType = "nodebuffer">(
  options: OdtDocumentOptions,
  output?: T,
): Promise<OutputByType[T]> {
  return generateDocumentSync(options, output);
}

export function generateDocumentStream(options: OdtDocumentOptions): ReadableStream<Uint8Array> {
  return streamUint8Array(generateOdtPackage(options));
}

function addPackageFile(files: OdfPackageFiles, path: string, data: DataType): void {
  validateMemberPath(path);
  if (path.endsWith("/"))
    throw new Error(`ODT package member path must not be a directory: ${path}`);
  if (path === "META-INF/manifest.xml")
    throw new Error("ODT package members cannot replace META-INF/manifest.xml");
  if (files[path] !== undefined)
    throw new Error(`ODT package member conflicts with modeled content: ${path}`);
  files[path] = typeof data === "string" ? data : toUint8Array(data);
}

function validateMemberPath(path: string): void {
  if (!path || path.startsWith("/") || path.endsWith("/") || path.split("/").includes(".."))
    throw new Error(`Invalid ODT package member path: ${path}`);
}

function validateObjectPath(path: string): void {
  if (!path || path.includes("/") || path.startsWith(".") || path.endsWith("/"))
    throw new Error(`Invalid ODT embedded object path: ${path}`);
}

/** Note ids auto-assign 1, 2, … per class, matching the docx model. */
export function notesContext(options: OdtDocumentOptions): NotesContext {
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
  const sectionSources = new Map(
    (options.sectionSources ?? []).map((source) => [source.sectionName ?? "", source]),
  );
  return {
    footnotes,
    endnotes,
    bookmarkNames: new Map(),
    bookmarkEndNames: new Map(),
    sequenceNames: new Set(
      (options.sequenceDeclarations ?? []).map((declaration) => declaration.name),
    ),
    variableTypes: new Map(
      (options.variableDeclarations ?? []).map((declaration) => [
        declaration.name,
        declaration.valueType,
      ]),
    ),
    forms: {
      automaticFocus: options.forms?.automaticFocus,
      designMode: options.forms?.designMode,
      containers: [...(options.forms?.containers ?? [])],
      models: options.forms?.models,
    },
    revisions: new Map(),
    comments: new Map((options.comments ?? []).map((comment) => [comment.id, comment])),
    annotationNames: new Map(),
    sectionSources,
  };
}

export function collectBookmarkNames(children: SectionChild[], notes: NotesContext): void {
  const pending: (SectionChild | NonNullable<ParagraphOptions["children"]>[number])[] = [
    ...children,
  ].reverse();
  while (pending.length > 0) {
    const child = pending.pop()!;
    if (typeof child === "string") continue;
    if ("bookmarkStart" in child)
      countBookmarkName(child.bookmarkStart.id, child.bookmarkStart.name, notes);
    if ("paragraph" in child && typeof child.paragraph !== "string")
      collectInlineBookmarkNames(child.paragraph.children ?? [], notes);
    if ("table" in child)
      for (const row of child.table.rows)
        if ("cells" in row)
          for (const cell of row.cells)
            if ("children" in cell) pending.push(...[...cell.children].reverse());
    if ("sdt" in child) pending.push(...[...(child.sdt.children ?? [])].reverse());
    if ("textbox" in child) pending.push(...[...(child.textbox.children ?? [])].reverse());
    if ("toc" in child) {
      pending.push(...[...(child.toc.leading ?? [])].reverse());
      pending.push(...[...(child.toc.entries ?? [])].reverse());
      pending.push(...[...(child.toc.trailing ?? [])].reverse());
    }
  }
}

export function collectInlineBookmarkNames(
  children: NonNullable<ParagraphOptions["children"]>,
  notes: NotesContext,
): void {
  for (const child of children) {
    if (typeof child === "string") continue;
    if ("bookmarkStart" in child)
      countBookmarkName(child.bookmarkStart.id, child.bookmarkStart.name, notes);
    if ("hyperlink" in child) namesMapFromInlineChildren(child.hyperlink.children ?? [], notes);
    if ("insertion" in child) namesMapFromInlineChildren(child.insertion.children, notes);
    if ("deletion" in child) namesMapFromInlineChildren(child.deletion.children, notes);
    if ("sdt" in child) namesMapFromInlineChildren(child.sdt.children ?? [], notes);
    if ("formField" in child) {
      for (const member of [...(child.instructionMembers ?? []), ...(child.resultMembers ?? [])]) {
        if ("bookmarkStart" in member)
          countBookmarkName(member.bookmarkStart.id, member.bookmarkStart.name, notes);
      }
    }
  }
}

function namesMapFromInlineChildren(
  children: NonNullable<ParagraphOptions["children"]>,
  notes: NotesContext,
): void {
  const pending: NonNullable<ParagraphOptions["children"]> = [...children].reverse();
  while (pending.length > 0) {
    const child = pending.pop()!;
    if (typeof child === "string") continue;
    if ("bookmarkStart" in child)
      countBookmarkName(child.bookmarkStart.id, child.bookmarkStart.name, notes);
    if ("hyperlink" in child) pending.push(...[...(child.hyperlink.children ?? [])].reverse());
    if ("insertion" in child) pending.push(...[...child.insertion.children].reverse());
    if ("deletion" in child) pending.push(...[...child.deletion.children].reverse());
  }
}

export function blocksXml(
  blocks: SectionChild[],
  styles: string[],
  images: OdtImage[],
  notes: NotesContext,
  numbering: DocumentOptions["numbering"],
  charts: OdtChart[],
  embeddedObjects: OdtEmbeddedObjectOptions[] = [],
): string {
  const parts: string[] = [];
  let index = 0;
  // Consecutive list paragraphs of the same level and kind group into one
  // text:list — the ODF shape for Word's bullet/numbering runs.
  while (index < blocks.length) {
    const block = blocks[index];
    if (!block) break;
    const listInfo = listParagraphLevel(block);
    if (listInfo === undefined) {
      parts.push(blockXml(block, styles, images, notes, numbering, charts, embeddedObjects));
      index += 1;
      continue;
    }
    const group: SectionChild[] = [];
    while (index < blocks.length) {
      const item = blocks[index];
      if (!item) break;
      const info = listParagraphLevel(item);
      if (
        info?.level !== listInfo.level ||
        info?.ordered !== listInfo.ordered ||
        info?.reference !== listInfo.reference
      )
        break;
      group.push(item);
      index += 1;
    }
    parts.push(listXml(group, listInfo, styles, images, notes, numbering, charts, embeddedObjects));
  }
  return parts.join("");
}

export async function parseDocument(data: DataType): Promise<OdtDocumentOptions> {
  return parseDocumentSync(await toUint8ArrayAsync(data));
}

export function parseDocumentSync(data: DataType): OdtDocumentOptions {
  try {
    return parseOdtPackage(toUint8Array(data));
  } catch (cause) {
    if (cause instanceof OdtParseError) throw cause;
    if (cause instanceof OdfSchemaError) {
      throw new OdtParseError(
        cause.message,
        cause.part,
        cause.path,
        cause.elementName,
        cause.reason,
        { cause },
      );
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

function parseOdtPackage(data: Uint8Array): OdtDocumentOptions {
  const { files, binaries, manifest } = readOcfSync(data, [MIME, TEMPLATE_MIME]);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:text");
  const styleContainer = childNamed(content, "office:automatic-styles");
  const styleMap = parseStyles(styleContainer);
  const graphicStyles = parseGraphicStyles(styleContainer);
  const chartBodies = parseEmbeddedCharts(manifest, files, binaries, files["styles.xml"]);
  const objectMediaTypes = new Map(
    manifest.entries
      .filter((entry) => entry.fullPath.endsWith("/"))
      .map((entry) => [entry.fullPath.replace(/\/$/, ""), entry.mediaType]),
  );
  objectMediaTypes.delete("");
  const embeddedObjects = new Map<string, OdtEmbeddedObjectOptions>();
  for (const [path, mediaType] of objectMediaTypes) {
    if (mediaType === undefined || (mediaType === CHART_MIME && chartBodies.has(path))) continue;
    if (files[`${path}/content.xml`] === undefined) continue;
    embeddedObjects.set(path, {
      path,
      mediaType,
      members: objectMembers(path, files, binaries),
    });
  }
  validateEmbeddedObjectBodies(body, binaries, embeddedObjects, chartBodies);
  const context: ParseContext = {
    styles: styleMap,
    listStyles: parseListStyles(styleContainer),
    graphicStyles,
    chartBodies,
    embeddedObjects,
    objectMediaTypes,
    listDefinitions: parseListNumberings(styleContainer),
    outline: parseOutlineStyle(files),
    binaries,
    usedBinaryPaths: new Set(),
    notes: { footnotes: [], endnotes: [], forms: { containers: [] } },
    comments: [],
    annotationIds: new Map(),
    bookmarkIds: new Map(),
    pendingFormatChanges: new Map(),
    sectionSources: new Map(),
    changes: new Map(),
    changeAlternatives: new Map(),
    pendingInsertions: new Set(),
    declaredSequences: [],
    declaredVariables: [],
    documentVariables: [],
    pendingSequences: new Set(),
    pendingVariables: new Map(),
    nextEmbeddedShapeId: 1,
  };
  const children = parseBlocks(body?.elements ?? [], context);
  const meta = Object.fromEntries(
    Object.entries(parseMeta(files)).filter(([, value]) => value !== undefined),
  );
  const result: OdtDocumentOptions = {
    ...meta,
    sections: [{ properties: parsePageLayout(files), children }],
  };
  const packageMembers = sourcePackageMembers(
    manifest,
    files,
    binaries,
    chartBodies,
    embeddedObjects,
    context.usedBinaryPaths,
  );
  if (embeddedObjects.size > 0) result.embeddedObjects = [...embeddedObjects.values()];
  if (packageMembers.length > 0) result.packageMembers = packageMembers;
  const styleResources: OdtAutomaticStyleResource[] = [...styleMap]
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
  if (styleResources.length > 0) result.styleResources = styleResources;
  if (hasOcfManifestOptions(manifest)) result.packageManifest = manifest;
  if (
    context.notes.forms.containers.length > 0 ||
    context.notes.forms.models?.length ||
    context.notes.forms.automaticFocus !== undefined ||
    context.notes.forms.designMode !== undefined
  )
    result.forms = context.notes.forms;
  if (context.comments.length > 0) result.comments = context.comments;
  if (context.declaredSequences.length > 0) result.sequenceDeclarations = context.declaredSequences;
  if (context.declaredVariables.length > 0) result.variableDeclarations = context.declaredVariables;
  if (context.documentVariables.length > 0)
    result.settings = { ...result.settings, docVars: context.documentVariables };
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
  if (context.indexAutoMarkFile) result.indexAutoMarkFile = context.indexAutoMarkFile;
  if (context.sectionSources.size > 0) result.sectionSources = [...context.sectionSources.values()];
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

function validateEmbeddedObjectBodies(
  body: Element | undefined,
  binaries: Record<string, Uint8Array>,
  embeddedObjects: Map<string, OdtEmbeddedObjectOptions>,
  chartBodies: Map<string, ChartSpaceOptions>,
): void {
  const references = descendantElements(body).flatMap((element) =>
    element.name === "draw:object-ole"
      ? [attributeString(element, "xlink:href")?.replace(/^\.\//, "").replace(/^\//, "")]
      : [],
  );
  for (const href of references) {
    if (href && (binaries[href] || embeddedObjects.has(href) || chartBodies.has(href))) continue;
    throw new OdtParseError(
      `content.xml: ${href ?? ""}: embedded object subdocument is missing`,
      "content.xml",
      "/draw:frame/draw:object-ole/@xlink:href",
      "draw:object-ole",
      "embedded-object-subdocument-missing",
    );
  }
}

function objectMembers(
  path: string,
  files: Record<string, string>,
  binaries: Record<string, Uint8Array>,
): PackageMemberOptions[] {
  return [
    ...Object.entries(files)
      .filter(([memberPath]) => memberPath.startsWith(`${path}/`))
      .map(([memberPath, data]) => ({ path: memberPath, data })),
    ...Object.entries(binaries)
      .filter(([memberPath]) => memberPath.startsWith(`${path}/`))
      .map(([memberPath, data]) => ({ path: memberPath, data })),
  ].sort((left, right) => left.path.localeCompare(right.path));
}

function sourcePackageMembers(
  manifest: OcfManifestOptions,
  files: Record<string, string>,
  binaries: Record<string, Uint8Array>,
  chartBodies: Map<string, ChartSpaceOptions>,
  embeddedObjects: Map<string, OdtEmbeddedObjectOptions>,
  usedBinaryPaths: Set<string>,
): PackageMemberOptions[] {
  const objectMemberPaths = new Set(
    [...embeddedObjects.values()].flatMap((object) => object.members.map((member) => member.path)),
  );
  const mediaTypes = new Map(
    manifest.entries.map((entry) => [entry.fullPath, entry.mediaType] as const),
  );
  const modeledPaths = new Set([
    "content.xml",
    "styles.xml",
    "meta.xml",
    "META-INF/manifest.xml",
    ...[...chartBodies.keys()].map((path) => `${path}/content.xml`),
    ...usedBinaryPaths,
  ]);
  return [
    ...Object.entries(files).map(([path, data]) => ({ path, data })),
    ...Object.entries(binaries).map(([path, data]) => ({ path, data })),
  ]
    .filter(({ path }) => !modeledPaths.has(path) && !objectMemberPaths.has(path))
    .sort((left, right) => left.path.localeCompare(right.path))
    .map(({ path, data }) => ({ path, mediaType: mediaTypes.get(path), data }));
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
  embeddedObjects: OdtEmbeddedObjectOptions[] = [],
): string {
  if ("toc" in child)
    return indexXml(child.toc, styles, images, notes, numbering, charts, embeddedObjects);
  if ("sdt" in child && child.sdt.properties.bibliography) return bibliographyIndexXml(child.sdt);
  if ("paragraph" in child)
    return paragraphXml(
      normalizeParagraph(child.paragraph),
      styles,
      images,
      notes,
      numbering,
      charts,
      embeddedObjects,
    );
  if ("bookmarkStart" in child) {
    const name = bookmarkDisplayName(child.bookmarkStart.id, child.bookmarkStart.name, notes);
    const endNames = notes.bookmarkEndNames.get(child.bookmarkStart.id) ?? [];
    endNames.push(name);
    notes.bookmarkEndNames.set(child.bookmarkStart.id, endNames);
    notes.bookmarkNames.set(child.bookmarkStart.id, name);
    return xmlElement("text:bookmark-start", {
      "xml:id": child.bookmarkStart.id ? `bookmark-${child.bookmarkStart.id}` : undefined,
      "text:name": name,
    });
  }
  if ("bookmarkEnd" in child) return bookmarkEndXml(child.bookmarkEnd.id, notes);
  if ("table" in child)
    return tableXml(child.table, styles, (block) =>
      blockXml(block, styles, images, notes, numbering, charts, embeddedObjects),
    );
  if ("sdt" in child && !child.sdt.properties.bibliography) {
    const styleName = /^odf:text-section(?:;style=([\s\S]*))?$/.exec(
      child.sdt.properties.tag ?? "",
    )?.[1];
    const source = notes.sectionSources.get(child.sdt.properties.alias ?? "");
    const sourceXml = source
      ? xmlElement("text:section-source", {
          "xlink:type": source.href ? "simple" : undefined,
          "xlink:href": source.href,
          "xlink:show": source.show,
          "text:filter-name": source.filterName,
          "text:section-name": source.targetSectionName,
        })
      : "";
    return xmlElement(
      "text:section",
      {
        "text:name": child.sdt.properties.alias ?? "",
        "text:style-name": styleName ? decodeStyleName(styleName) : undefined,
        "text:protected": child.sdt.properties.lock === "sdtLocked" ? true : undefined,
      },
      [
        sourceXml,
        blocksXml(
          child.sdt.children ?? [],
          styles,
          images,
          notes,
          numbering,
          charts,
          embeddedObjects,
        ),
      ],
    );
  }
  return "";
}

export function parseBlock(element: Element, context: ParseContext): SectionChild {
  if (element.name === "text:index-title")
    return {
      sdt: {
        properties: { alias: attributeString(element, "text:name") ?? "" },
        children: blockChildren(parseBlocks(element.elements ?? [], context)),
      },
    };
  if (element.name === "table:table")
    return parseTable(
      element,
      (name) => context.styles.get(name)?.columnWidth,
      (child) => parseBlock(child, context),
    );
  const paragraph = parseParagraph(element, context);
  return { paragraph };
}

function hoistBoundaryTrackMarkers(element: Element): Element[] {
  const nodes = element.elements ?? [];
  const isMarker = (node: Element) =>
    node.name === "text:change-start" || node.name === "text:change-end";
  const markerId = (node: Element) => attributeString(node, "text:change-id");
  const markers = nodes.filter(isMarker);
  if (markers.length === 0) return [element];
  if (nodes.length === 1 && isMarker(nodes[0]!)) return [nodes[0]!];

  const hasPairedMarker = (marker: Element) =>
    nodes.some(
      (node) =>
        node !== marker &&
        isMarker(node) &&
        node.name !== marker.name &&
        markerId(node) === markerId(marker),
    );
  const first = nodes[0]!;
  const last = nodes.at(-1)!;
  const boundary = isMarker(first) && !hasPairedMarker(first) ? first : undefined;
  const trailing = !boundary && isMarker(last) && !hasPairedMarker(last) ? last : undefined;
  if (!boundary && !trailing) return [element];
  const marker = (boundary ?? trailing)!;
  const rest = nodes.filter((node) => node !== marker);
  return [marker, { ...element, elements: rest }];
}

function expandBoundaryTrackRanges(elements: Element[]): Element[] {
  const result: Element[] = [];
  let index = 0;
  while (index < elements.length) {
    const element = elements[index]!;
    const nodes = element.elements ?? [];
    const startNodeIndex = nodes.findIndex((node) => node.name === "text:change-start");
    const start = startNodeIndex >= 0 ? nodes[startNodeIndex] : undefined;
    if ((element.name !== "text:p" && element.name !== "text:h") || start === undefined) {
      result.push(...hoistBoundaryTrackMarkers(element));
      index += 1;
      continue;
    }

    const id = attributeString(start, "text:change-id");
    if (
      nodes.some(
        (node) => node.name === "text:change-end" && attributeString(node, "text:change-id") === id,
      )
    ) {
      result.push(...hoistBoundaryTrackMarkers(element));
      index += 1;
      continue;
    }
    let endIndex = -1;
    let endNodeIndex = -1;
    for (const [candidateIndex, candidate] of elements.entries()) {
      if (candidateIndex <= index) continue;
      const markerIndex = (candidate.elements ?? []).findIndex(
        (node) => node.name === "text:change-end" && attributeString(node, "text:change-id") === id,
      );
      if (markerIndex >= 0) {
        endIndex = candidateIndex;
        endNodeIndex = markerIndex;
        break;
      }
    }
    if (endIndex < 0 || endNodeIndex < 0) {
      result.push(...hoistBoundaryTrackMarkers(element));
      index += 1;
      continue;
    }

    const end = elements[endIndex]!;
    const endNodes = end.elements ?? [];
    const changed = [] as Element[];
    if (startNodeIndex > 0) result.push({ ...element, elements: nodes.slice(0, startNodeIndex) });
    if (startNodeIndex + 1 < nodes.length) {
      changed.push({ ...element, elements: nodes.slice(startNodeIndex + 1) });
    }
    changed.push(...elements.slice(index + 1, endIndex));
    const nodesBeforeEnd = endNodes.slice(0, endNodeIndex);
    if (nodesBeforeEnd.length > 0) changed.push({ ...end, elements: nodesBeforeEnd });
    result.push({ ...start, elements: changed });

    const nodesAfterEnd = endNodes.slice(endNodeIndex + 1);
    if (nodesAfterEnd.length > 0) result.push({ ...end, elements: nodesAfterEnd });
    index = endIndex + 1;
  }
  return result;
}

export function parseBlocks(
  elements: Element[],
  context: ParseContext,
  listDepth = 0,
  listState?: { name: string; ordered: boolean },
): SectionChild[] {
  const result: SectionChild[] = [];
  let pendingPageBreak = false;
  let resumeAfter = -1;
  const hoistedElements = expandBoundaryTrackRanges(elements).flatMap((element) => {
    if (element.name !== "text:p" && element.name !== "text:h") return [element];
    return hoistBoundaryTrackMarkers(element);
  });
  const push = (...blocks: SectionChild[]): void => {
    let tracked = blocks;
    for (const id of context.pendingInsertions) {
      const change = context.changeAlternatives.get(`${id}:insertion`) ?? context.changes.get(id);
      if (change?.kind === "insertion") {
        tracked = trackRevisionBlocks(
          "insertion",
          { id: change.id, author: change.author, date: change.date },
          tracked,
        );
      }
    }
    for (const revision of context.pendingFormatChanges.values())
      tracked = formatRevisionBlocks(tracked, revision);
    result.push(...tracked);
  };
  for (const [elementIndex, element] of hoistedElements.entries()) {
    if (elementIndex <= resumeAfter) continue;
    if (element.name === "text:alphabetical-index-auto-mark-file") {
      assertElementNames(element, ["xlink:type", "xlink:href"]);
      const href = attributeString(element, "xlink:href");
      if (!href) throw unknownOdtElement(element);
      if (context.indexAutoMarkFile) throw unknownOdtElement(element);
      context.indexAutoMarkFile = { href };
      continue;
    }
    if (element.name === "text:section") {
      const styleName = attributeString(element, "text:style-name");
      const sourceElement = element.elements?.find(
        (child) => child.name === "text:section-source" || child.name === "text:section-source-dde",
      );
      if (sourceElement?.name === "text:section-source") {
        assertElementNames(sourceElement, [
          "xlink:type",
          "xlink:href",
          "xlink:show",
          "text:filter-name",
          "text:section-name",
        ]);
        const href = attributeString(sourceElement, "xlink:href");
        if (!href) throw unknownOdtElement(sourceElement);
        const show = attributeString(sourceElement, "xlink:show");
        if (show !== undefined && show !== "embed") throw unknownOdtElement(sourceElement);
        const sectionName = attributeString(element, "text:name") ?? "";
        const source = {
          href,
          ...(sectionName ? { sectionName } : {}),
          ...(attributeString(sourceElement, "text:filter-name")
            ? { filterName: attributeString(sourceElement, "text:filter-name") }
            : {}),
          ...(show ? { show: "embed" as const } : {}),
          ...(attributeString(sourceElement, "text:section-name")
            ? { targetSectionName: attributeString(sourceElement, "text:section-name") }
            : {}),
        };
        if (context.sectionSources.has(sectionName)) throw unknownOdtElement(sourceElement);
        context.sectionSources.set(sectionName, source);
      } else if (sourceElement) {
        throw unknownOdtElement(sourceElement);
      }
      push({
        sdt: {
          properties: {
            alias: attributeString(element, "text:name") ?? "",
            ...(styleName ? { tag: encodeSectionTag(styleName) } : {}),
            ...(attributeString(element, "text:protected") === "true"
              ? { lock: "sdtLocked" as const }
              : {}),
          },
          children: blockChildren(
            parseBlocks(
              (element.elements ?? []).filter((child) => child !== sourceElement),
              context,
              listDepth,
              listState,
            ),
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
        push(...parseBlocks(item.elements ?? [], context, listDepth + 1, state));
      }
      continue;
    }
    if (element.name === "text:soft-page-break") {
      pendingPageBreak = true;
      continue;
    }
    if (element.name === "text:index-title") {
      push(parseBlock(element, context));
      continue;
    }
    if (element.name === "text:bookmark-start") {
      const sourceName = attributeString(element, "text:name") ?? "";
      const decoded = canonicalBookmarkName(sourceName);
      const name = decoded.name;
      if (!name) throw unknownOdtElement(element);
      const existingId = context.bookmarkIds.get(name);
      const encodedId = /^(?:bookmark-)?(\d+)$/.exec(
        attributeString(element, "text:id") ??
          attributeString(element, "xml:id") ??
          (decoded.encodedId !== undefined ? String(decoded.encodedId) : ""),
      )?.[1];
      if (existingId !== undefined && sourceName === name && encodedId !== String(existingId))
        throw unknownOdtElement(element);
      const id =
        decoded.encodedId ??
        existingId ??
        (encodedId ? Number(encodedId) : nextBookmarkId(context));
      if (existingId === undefined) context.bookmarkIds.set(name, id);
      if (sourceName !== name) context.bookmarkIds.set(sourceName, id);
      push({
        bookmarkStart: { id, name },
      });
      continue;
    }
    if (element.name === "text:bookmark-end") {
      const sourceName = attributeString(element, "text:name") ?? "";
      const name = canonicalBookmarkName(sourceName).name;
      const id = context.bookmarkIds.get(sourceName) ?? context.bookmarkIds.get(name);
      if (!name || id === undefined) throw unknownOdtElement(element);
      push({
        bookmarkEnd: { id },
      });
      continue;
    }
    if (element.name === "text:reference-mark-start") {
      const name = attributeString(element, "text:name") ?? "";
      const id = context.bookmarkIds.get(name) ?? nextBookmarkId(context);
      context.bookmarkIds.set(name, id);
      push({
        bookmarkStart: { id, name },
      });
      continue;
    }
    if (element.name === "office:annotation") {
      push({ paragraph: { children: parseAnnotationStart(element, context) } });
      continue;
    }
    if (
      [
        "field:fieldmark-start",
        "office:event-listeners",
        "table:calculation-settings",
        "text:alphabetical-index-auto-mark-file",
        "text:alphabetical-index-mark",
        "text:bookmark-end",
        "text:change",
        "text:conditional-text",
        "text:deletion",
        "text:insertion",
        "text:modification-date",
        "text:section-source",
        "text:user-field-decl",
        "text:user-index-mark",
      ].includes(element.name ?? "")
    ) {
      throw invalidOdtElement(element);
    }
    if (element.name === "text:change-start") {
      const id = attributeString(element, "text:change-id") ?? "";
      const change =
        ["insertion", "deletion", "format"]
          .map((kind) => context.changeAlternatives.get(`${id}:${kind}`))
          .find((candidate) => candidate !== undefined) ?? context.changes.get(id);
      if (context.pendingFormatChanges.has(id) || change === undefined)
        throw unknownOdtElement(element);
      let endIndex = (element.elements?.length ?? 0) > 0 ? elementIndex : -1;
      if (endIndex < 0) {
        let depth = 1;
        for (let index = elementIndex + 1; index < elements.length && endIndex < 0; index += 1) {
          const node = elements[index]!;
          if (node.name === "text:change-start" && attributeString(node, "text:change-id") === id)
            depth += 1;
          if (node.name === "text:change-end" && attributeString(node, "text:change-id") === id) {
            depth -= 1;
            if (depth === 0) endIndex = index;
          }
        }
      }
      if (endIndex < 0) endIndex = elements.length;
      const revision = { id: change.id, author: change.author, date: change.date };
      const changed = parseBlocks(
        element.elements ?? hoistedElements.slice(elementIndex + 1, endIndex),
        context,
        listDepth,
        listState,
      );
      push(
        ...(change.kind === "format"
          ? formatRevisionBlocks(changed, revision)
          : trackRevisionBlocks(change.kind, revision, changed)),
      );
      resumeAfter = endIndex;
      continue;
    }
    if (element.name === "text:change-end") {
      const id = attributeString(element, "text:change-id") ?? "";
      if (context.pendingInsertions.has(id) || context.pendingFormatChanges.has(id)) {
        context.pendingInsertions.delete(id);
        context.pendingFormatChanges.delete(id);
        continue;
      }
      throw invalidOdtElement(element);
    }
    if (element.name === "text:change") {
      push({ paragraph: { children: parseRuns(element, context) } });
      continue;
    }
    if (element.name === "office:annotation-end") {
      push({ paragraph: { children: parseAnnotationEnd(element, context) } });
      continue;
    }
    if (element.name === "text:illustration-index") {
      push({ toc: parseIndex(element, context) });
      continue;
    }
    if (
      element.name === "text:table-of-content" ||
      element.name === "text:table-index" ||
      element.name === "text:object-index" ||
      element.name === "text:user-index" ||
      element.name === "text:alphabetical-index"
    ) {
      const sourceName =
        element.name === "text:table-of-content"
          ? "text:table-of-content-source"
          : element.name === "text:table-index"
            ? "text:table-index-source"
            : element.name === "text:object-index"
              ? "text:object-index-source"
              : element.name === "text:user-index"
                ? "text:user-index-source"
                : "text:alphabetical-index-source";
      push({ toc: parseContentIndex(element, sourceName, context) });
      continue;
    }
    if (element.name === "text:bibliography") {
      push({
        sdt: parseBibliography(element),
      });
      continue;
    }
    if (isShapeElement(element.name)) {
      push(shapeBlock(parseShape(element)));
      continue;
    }
    if (element.name === "draw:g") {
      push({ paragraph: { children: [parseGroup(element, context)] } });
      continue;
    }
    if (element.name === "draw:frame") {
      if (childNamed(element, "draw:text-box")) {
        push({ paragraph: { children: [parseTextBoxFrame(element, context)] } });
        continue;
      }
      const child = parseFrame(element, context);
      if (child) push(child);
      continue;
    }
    if (element.name === "draw:a") {
      push({ paragraph: { children: [parseShapeHyperlink(element)] } });
      continue;
    }
    if (element.name === "office:forms") {
      context.notes.forms = parseOfficeForms(element);
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
    if (element.name === "text:user-field-decls") {
      parseUserFieldDeclarations(element, context);
      continue;
    }
    if (element.name === "text:dde-connection-decls") {
      parseDdeConnectionDeclarations(element, context);
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
        "text:user-field-decls",
        "text:dde-connection-decls",
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
        push(pendingPageBreak ? withPageBreakBefore(mappedChild) : mappedChild);
        pendingPageBreak = false;
      } else {
        push(pendingPageBreak ? withPageBreakBefore(child) : child);
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

function parseFrame(element: Element, context: ParseContext): SectionChild | undefined {
  const children = childNamed(element, "draw:object")
    ? parseChartFrame(element, context)
    : parsePictureFrame(element, context);
  if (children.length === 0) return undefined;
  return { paragraph: { children } };
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

export function invalidOdtElement(element: Element): OdtParseError {
  const name = element.name ?? "";
  const path = `/office:document-content/office:body/office:text/${name}`;
  return new OdtParseError(
    `content.xml: ${path}: ${name}: element is not allowed in office:text`,
    "content.xml",
    path,
    name,
    "invalid-element-placement",
  );
}

export function assertElementNames(
  element: Element,
  allowedAttributes: string[],
  allowedChildren: string[] = [],
): void {
  const unknownAttributes = unknownAttributeNames(element, allowedAttributes);
  if (unknownAttributes.length) {
    throw new OdtParseError(
      `content.xml: unknown ${element.name} attribute: ${unknownAttributes.join(", ")}`,
      "content.xml",
      `/office:document-content/office:body/office:text/${element.name}`,
      element.name,
      "attribute has no canonical mapping",
    );
  }
  const unknownChild = unknownChildElement(element, allowedChildren);
  if (unknownChild) {
    throw new OdtParseError(
      `content.xml: unknown ${element.name} child: ${unknownChild.name}`,
      "content.xml",
      `/office:document-content/office:body/office:text/${element.name}`,
      unknownChild.name,
      "child has no canonical mapping",
    );
  }
}

export function malformedOdtElement(name: string, reason: string): OdtParseError {
  const path = `/office:document-content/office:body/office:text/${name}`;
  return new OdtParseError(`content.xml: ${name}: ${reason}`, "content.xml", path, name, reason);
}
