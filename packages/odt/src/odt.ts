import { toUint8Array } from "@office-open/core";
import type { ChartSpaceOptions } from "@office-open/core";
import type { FillOptions, OutlineOptions } from "@office-open/core/drawing";
import type {
  BlockContentChild,
  ChartOptions,
  FormFieldOptions,
  ParagraphChild,
  DocumentOptions,
  ParagraphOptions,
  PictureOptions,
  ShapeOptions,
  RunOptions,
  SectionChild,
  SectionOptions,
} from "@office-open/docx";
import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  escapeText,
  generateOcf,
  lengthToEmu,
  metaXml,
  ODF_NAMESPACES,
  parseMeta,
  readOcf,
  readXml,
  textOf,
  xmlElement,
  type OdfFiles,
  type OdfPackageFiles,
} from "@office-open/ocf";
import {
  CHART_MIME,
  chartBodyXml,
  graphicFill,
  graphicOutline,
  parseEmbeddedCharts,
  parseGraphicStyles,
  OdfSchemaError,
  PRESET_GEOMETRY_DOCX,
  presetGeometryOdf,
  pushShapeStyle,
} from "@office-open/odf-schema";
import type { GraphicStyle } from "@office-open/odf-schema";
import type { Element } from "@office-open/xml";

import { OdtParseError } from "./error";
import type { OdtDocumentOptions } from "./semantics";
import { parseTable, tableXml } from "./table";

const MIME = "application/vnd.oasis.opendocument.text";
const NAMESPACES = `${ODF_NAMESPACES} xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:xforms="http://www.w3.org/2002/xforms"`;

/** Track-change types derived from the canonical paragraph child union. */
type CanonicalTrackChange = Extract<ParagraphChild, { insertion: unknown }>["insertion"];
type ChangedProperties = Omit<CanonicalTrackChange, "children">;
type TrackChangeChild = CanonicalTrackChange["children"][number];
type CanonicalSimpleField = Extract<ParagraphChild, { simpleField: unknown }>["simpleField"];
type PictureType = NonNullable<PictureOptions["type"]>;
type RasterPictureType = Exclude<PictureType, "svg">;

function isRasterPictureType(value: string): value is RasterPictureType {
  return ["png", "jpg", "gif", "bmp", "tif", "ico", "emf", "wmf"].includes(value);
}

function isTrackChangeChild(child: string | ParagraphChild): child is TrackChangeChild {
  if (typeof child === "string") return true;
  return (
    "text" in child ||
    "break" in child ||
    "pageBreak" in child ||
    "columnBreak" in child ||
    "picture" in child ||
    "chart" in child ||
    "wpsShape" in child ||
    "formField" in child
  );
}

/** Section headers/footers render as master-page style:header/style:footer. */
function masterHeaderFooter(
  section: Pick<SectionOptions, "headers" | "footers"> | undefined,
): string {
  return (
    headerFooterXml(section?.headers?.default, "style:header") +
    headerFooterXml(section?.footers?.default, "style:footer")
  );
}

function headerFooterXml(children: SectionChild[] | undefined, element: string): string {
  const blocks = (children ?? [])
    .filter((child) => "paragraph" in child)
    .map((child) => headerParagraphXml(normalizeParagraph(child.paragraph)));
  if (blocks.length === 0) return "";
  return xmlElement(element, undefined, blocks);
}

function headerParagraphXml(options: ParagraphOptions): string {
  if (options.text !== undefined && options.children === undefined)
    return xmlElement("text:p", undefined, [spacesXml(options.text)]);
  return xmlElement(
    "text:p",
    undefined,
    (options.children ?? []).map((child) => {
      if (typeof child === "string") return spacesXml(child);
      if ("simpleField" in child)
        return headerFieldXml(
          (child as { simpleField: { instruction?: string; cachedValue?: string } }).simpleField,
        );
      if ("text" in child) return spacesXml(String((child as { text?: string }).text ?? ""));
      return "";
    }),
  );
}

function headerFieldXml(field: { instruction?: string; cachedValue?: string }): string {
  const value = field.cachedValue ?? "1";
  if (/NUMPAGES/i.test(field.instruction ?? ""))
    return xmlElement("text:page-count", undefined, [value]);
  return xmlElement("text:page-number", { "text:select-page": "current" }, [value]);
}

/** Master-page header/footer paragraphs → docx section header/footer children. */
function parseMasterHeaderFooter(files: OdfFiles): {
  headers?: SectionOptions["headers"];
  footers?: SectionOptions["headers"];
} {
  const master = childNamed(
    childNamed(readXml(files, "styles.xml"), "office:master-styles"),
    "style:master-page",
  );
  const header = childNamed(master, "style:header");
  const footer = childNamed(master, "style:footer");
  const headerParagraphs = header ? parseHeaderParagraphs(header) : [];
  const footerParagraphs = footer ? parseHeaderParagraphs(footer) : [];
  return {
    headers: headerParagraphs.length > 0 ? { default: headerParagraphs } : undefined,
    footers: footerParagraphs.length > 0 ? { default: footerParagraphs } : undefined,
  };
}

function parseHeaderParagraphs(container: Element): SectionChild[] {
  return childrenNamed(container, "text:p").map((paragraph) => {
    const runs = parseHeaderRuns(paragraph);
    const options: ParagraphOptions =
      runs.length === 1 && typeof runs[0] === "string" ? { text: runs[0] } : { children: runs };
    return { paragraph: options } as SectionChild;
  });
}

function parseHeaderRuns(paragraph: Element): NonNullable<ParagraphOptions["children"]> {
  return (paragraph.elements ?? []).flatMap((child): NonNullable<ParagraphOptions["children"]> => {
    if (child.type === "text") return [String(child.text ?? "")];
    if (child.name === "text:page-number")
      return [{ simpleField: { instruction: " PAGE ", cachedValue: textOf(child) ?? "" } }];
    if (child.name === "text:page-count")
      return [{ simpleField: { instruction: " NUMPAGES ", cachedValue: textOf(child) ?? "" } }];
    return [];
  });
}

/** ODF style:tab-stop → the closest typed docx tab stop. */
function parseTabStop(element: Element): TabStop {
  const type = attributeString(element, "style:type");
  const leaderText = attributeString(element, "style:leader-text");
  return {
    type: type === "char" ? "decimal" : ((type ?? "left") as TabStop["type"]),
    position: lengthToTwips(attributeString(element, "style:position")) ?? 0,
    leader:
      leaderText === "."
        ? "dot"
        : leaderText === "-"
          ? "hyphen"
          : leaderText === "_"
            ? "underscore"
            : leaderText === "·"
              ? "middleDot"
              : undefined,
  };
}

/** Runs of two or more spaces emit text:s so XML whitespace folding keeps them. */
function spacesXml(text: string): string {
  const parts: string[] = [];
  let index = 0;
  for (const match of text.matchAll(/ {2,}/g)) {
    const start = match.index ?? 0;
    parts.push(escapeText(text.slice(index, start)));
    parts.push(xmlElement("text:s", { "text:c": match[0].length }));
    index = start + match[0].length;
  }
  parts.push(escapeText(text.slice(index)));
  return parts.join("");
}

/** Binary image collected during generation — emitted as a Pictures/ entry. */
interface OdtImage {
  path: string;
  data: Uint8Array;
}

interface CharacterProperties {
  bold?: boolean;
  italic?: boolean;
  underline?: "single";
  strike?: boolean;
  size?: number;
  color?: string;
  font?: string;
}

type StyleMap = Map<
  string,
  {
    alignment?: string;
    columnWidth?: number;
    pageBreakBefore?: boolean;
    tabStops?: TabStop[];
    character: CharacterProperties;
  }
>;

/** Shared lookup state threaded through the ODT parse pipeline. */
interface ParseContext {
  styles: StyleMap;
  listStyles: Map<string, boolean>;
  graphicStyles: Map<string, GraphicStyle>;
  chartBodies: Map<string, ChartSpaceOptions>;
  listDefinitions: AbstractNumbering[];
  outline?: AbstractNumbering;
  binaries: Record<string, Uint8Array>;
  notes: { footnotes: NoteEntry[]; endnotes: NoteEntry[] };
  bookmarkIds: Map<string, number>;
  changes: Map<
    string,
    ChangedProperties & { kind: "insertion" | "deletion"; children?: Element[] }
  >;
  pendingSequences: Set<string>;
  pendingVariables: Map<string, "float" | "string">;
  trackRevisions?: boolean;
}

/** Footnote/endnote entry, indexed from the shared docx document model. */
type NoteEntry = NonNullable<DocumentOptions["footnotes"]>[number];

type NoteChildren = NoteEntry["children"];

/** Canonical form field used as the OOXML-family equivalent of a simple ODF control. */
type CanonicalFormField = FormFieldOptions;

/** Canonical tracked change metadata and deleted content for ODF emission. */
type CanonicalRevision = ChangedProperties & { children?: TrackChangeChild[] };

/** Docx tab stop, indexed from the shared paragraph model. */
type TabStop = NonNullable<ParagraphOptions["tabStops"]>[number];

/** Docx font-table entry, indexed from the shared document model. */
type FontEntry = NonNullable<DocumentOptions["fonts"]>[number];

type FootnoteProperties = NonNullable<
  NonNullable<DocumentOptions["settings"]>["footnoteProperties"]
>;

type AbstractNumbering = NonNullable<
  NonNullable<DocumentOptions["numbering"]>["abstractNumberings"]
>[number];

type NumberingLevel = AbstractNumbering["levels"][number];

type DocumentDefaults = NonNullable<
  NonNullable<NonNullable<DocumentOptions["styles"]>["default"]>["document"]
>;

type EndnoteProperties = NonNullable<NonNullable<DocumentOptions["settings"]>["endnoteProperties"]>;

/** docx ST_NumberFormat tokens with a direct ODF num-format token. */
const NUM_FORMAT_ODF: Record<string, string> = {
  decimal: "1",
  lowerLetter: "a",
  upperLetter: "A",
  lowerRoman: "i",
  upperRoman: "I",
};

const NUM_FORMAT_DOCX: Record<string, FootnoteProperties["numFmt"]> = {
  "1": "decimal",
  a: "lowerLetter",
  A: "upperLetter",
  i: "lowerRoman",
  I: "upperRoman",
} as Record<string, FootnoteProperties["numFmt"]>;

/** Note bodies keyed by reference id, threaded through ODT emission. */
interface NotesContext {
  footnotes: Map<number, NoteChildren>;
  endnotes: Map<number, NoteChildren>;
  bookmarkNames: Map<number, string>;
  forms: CanonicalFormField[];
  revisions: Map<number, CanonicalRevision>;
  sequenceNames: Set<string>;
  variableTypes: Map<string, "float" | "string">;
}

/** Embedded chart subdocument collected during generation. */
interface OdtChart {
  path: string;
  chart: ChartSpaceOptions;
}

export function generateOdt(options: OdtDocumentOptions): Uint8Array {
  const styles: string[] = [];
  const blocks = options.sections.flatMap((section) => section.children);
  const images: OdtImage[] = [];
  const charts: OdtChart[] = [];
  const notes = notesContext(options);
  collectBookmarkNames(blocks, notes.bookmarkNames);
  const bodyBlocks = blocksXml(blocks, styles, images, notes, options.numbering, charts);
  const body = declarationsXml(notes) + bodyBlocks;
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
  );
}

/** Note ids auto-assign 1, 2, … per class, matching the docx model. */
function notesContext(options: DocumentOptions): NotesContext {
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
    forms: [],
    revisions: new Map(),
    sequenceNames: new Set(),
    variableTypes: new Map(),
  };
}

function collectBookmarkNames(children: SectionChild[], names: Map<number, string>): void {
  for (const child of children) {
    if ("bookmarkStart" in child) names.set(child.bookmarkStart.id, child.bookmarkStart.name);
    if ("sdt" in child) collectBookmarkNames(child.sdt.children ?? [], names);
  }
}

function blocksXml(
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

export function parseOdt(data: Uint8Array): OdtDocumentOptions {
  try {
    return parseOdtDocument(data);
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

function parseOdtDocument(data: Uint8Array): OdtDocumentOptions {
  const { files, binaries, manifest } = readOcf(data, MIME);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:text");
  const styleContainer = childNamed(content, "office:automatic-styles");
  const styleMap = parseStyles(styleContainer);
  const graphicStyles = parseGraphicStyles(styleContainer);
  const chartBodies = parseEmbeddedCharts(manifest, files);
  const context: ParseContext = {
    styles: styleMap,
    listStyles: parseListStyles(styleContainer),
    graphicStyles,
    chartBodies,
    listDefinitions: parseListNumberings(styleContainer),
    outline: parseOutlineStyle(files),
    binaries,
    notes: { footnotes: [], endnotes: [] },
    bookmarkIds: new Map(),
    changes: new Map(),
    pendingSequences: new Set(),
    pendingVariables: new Map(),
  };
  const children = parseBlocks(body?.elements ?? [], context);
  if (context.pendingSequences.size) {
    throw new OdtParseError(
      `content.xml: sequence declaration without a canonical sequence field: ${[...context.pendingSequences].join(", ")}`,
      "content.xml",
      "/office:document-content/office:body/office:text/text:sequence-decls",
      "text:sequence-decl",
      "declaration has no canonical field",
    );
  }
  if (context.pendingVariables.size) {
    throw new OdtParseError(
      `content.xml: variable declaration without a canonical variable field: ${[...context.pendingVariables.keys()].join(", ")}`,
      "content.xml",
      "/office:document-content/office:body/office:text/text:variable-decls",
      "text:variable-decl",
      "declaration has no canonical field",
    );
  }
  const meta = Object.fromEntries(
    Object.entries(parseMeta(files)).filter(([, value]) => value !== undefined),
  );
  const result: DocumentOptions = {
    ...meta,
    sections: [{ properties: parsePageLayout(files), children }],
  };
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
  if (defaultStyle)
    result.styles = {
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

function contentXml(body: string, styles: string[], fontFaces: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3">${fontFaces}<office:automatic-styles>${styles.join(
    "",
  )}</office:automatic-styles><office:body><office:text>${body}</office:text></office:body></office:document-content>`;
}

/** Docx font table → ODF font-face declarations (no embedded data in ODF). */
function fontFaceDecls(fonts: DocumentOptions["fonts"]): string {
  if (!fonts?.length) return "";
  return xmlElement("office:font-face-decls", undefined, fonts.map(fontFaceXml));
}

function fontFaceXml(font: FontEntry): string {
  return xmlElement("style:font-face", {
    "style:name": font.name,
    "style:font-family-generic": font.family,
    "style:font-pitch": font.pitch,
    "svg:panose-1": font.panose1,
  });
}

function parseFontFace(element: Element): FontEntry {
  return {
    name: attributeString(element, "style:name") ?? "",
    family: attributeString(element, "style:font-family-generic"),
    pitch: attributeString(element, "style:font-pitch"),
    panose1: attributeString(element, "svg:panose-1"),
  };
}

function documentStylesXml(
  section: Pick<SectionOptions, "headers" | "footers" | "properties"> | undefined,
  settings: DocumentOptions["settings"],
  numbering: DocumentOptions["numbering"],
  styles: DocumentOptions["styles"],
): string {
  const properties = section?.properties;
  const pageSize = typeof properties?.pageSize === "object" ? properties.pageSize : undefined;
  const pageMargin = typeof properties?.pageMargin === "object" ? properties.pageMargin : undefined;
  const layoutAttributes = [
    pageSize?.width !== undefined && `fo:page-width="${twipsToLength(pageSize.width)}"`,
    pageSize?.height !== undefined && `fo:page-height="${twipsToLength(pageSize.height)}"`,
    pageSize?.orientation && `style:print-orientation="${pageSize.orientation}"`,
    pageMargin?.top !== undefined && `fo:margin-top="${twipsToLength(pageMargin.top)}"`,
    pageMargin?.right !== undefined && `fo:margin-right="${twipsToLength(pageMargin.right)}"`,
    pageMargin?.bottom !== undefined && `fo:margin-bottom="${twipsToLength(pageMargin.bottom)}"`,
    pageMargin?.left !== undefined && `fo:margin-left="${twipsToLength(pageMargin.left)}"`,
  ].filter(Boolean);
  const headerFooter = masterHeaderFooter(section);
  const needsMaster = layoutAttributes.length > 0 || headerFooter !== "";
  const pageLayout = needsMaster
    ? `<style:page-layout style:name="pm1"><style:page-layout-properties ${layoutAttributes.join(
        " ",
      )}/></style:page-layout>`
    : "";
  const masterStyles = needsMaster
    ? `<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1">${headerFooter}</style:master-page></office:master-styles>`
    : "";
  const notes =
    notesConfigurationXml(settings) + outlineStyleXml(numbering) + defaultStyleXml(styles);
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NAMESPACES} office:version="1.3"><office:styles>${notes}</office:styles><office:automatic-styles>${pageLayout}</office:automatic-styles>${masterStyles}</office:document-styles>`;
}

/** Heading-linked abstract numbering renders as text:outline-style. */
function outlineStyleXml(numbering: DocumentOptions["numbering"]): string {
  const abstract = numbering?.abstractNumberings?.find((entry) =>
    entry.levels.some((level) => level.paragraphStyle?.startsWith("Heading")),
  );
  if (!abstract) return "";
  return xmlElement(
    "text:outline-style",
    { "style:name": abstract.reference },
    abstract.levels.map(outlineLevelXml),
  );
}

function outlineLevelXml(level: NumberingLevel): string {
  const template = level.text ?? "";
  const tokens = template.match(/%\d/g) ?? [];
  const prefix = template.split(/%\d/)[0] ?? "";
  const suffix = template.split(/%\d/).at(-1) ?? "";
  return xmlElement("text:outline-level-style", {
    "text:level": level.level + 1,
    "text:style-name": level.paragraphStyle,
    "style:num-format": level.format ? (NUM_FORMAT_ODF[level.format] ?? level.format) : undefined,
    "style:num-prefix": prefix || undefined,
    "style:num-suffix": suffix || undefined,
    "text:display-levels": tokens.length > 0 ? tokens.length : undefined,
    "text:start-value": level.start,
  });
}

/** text:outline-style → a heading-linked abstract numbering definition. */
function parseOutlineStyle(files: OdfFiles): AbstractNumbering | undefined {
  const outline = childNamed(
    childNamed(readXml(files, "styles.xml"), "office:styles"),
    "text:outline-style",
  );
  const levels = childrenNamed(outline, "text:outline-level-style").map((level) => {
    const oneBased = attributeNumber(level, "text:level") ?? 1;
    const display = attributeNumber(level, "text:display-levels") ?? 1;
    const prefix = attributeString(level, "style:num-prefix") ?? "";
    const suffix = attributeString(level, "style:num-suffix") ?? "";
    const tokens = Array.from(
      { length: display },
      (_, index) => `%${oneBased - display + 1 + index}`,
    ).join(".");
    const rawFormat = attributeString(level, "style:num-format");
    const start = attributeNumber(level, "text:start-value");
    return {
      level: oneBased - 1,
      ...(rawFormat
        ? { format: (NUM_FORMAT_DOCX[rawFormat] ?? rawFormat) as NumberingLevel["format"] }
        : {}),
      text: `${prefix}${tokens}${suffix}`,
      ...(start !== undefined ? { start } : {}),
      ...(attributeString(level, "text:style-name")
        ? { paragraphStyle: attributeString(level, "text:style-name") }
        : {}),
    } as NumberingLevel;
  });
  if (levels.length === 0) return undefined;
  return { reference: attributeString(outline, "style:name") ?? "Outline", levels };
}

/** Docx level suffix → ODF text:label-followed-by token. */
const SUFFIX_ODF: Record<string, string> = {
  nothing: "nothing",
  space: "space",
  tab: "listtab",
};

/** ODF text:label-followed-by token → docx level suffix. */
const SUFFIX_DOCX: Record<string, string> = {
  listtab: "tab",
  nothing: "nothing",
  space: "space",
};

/** Non-outline abstract numberings render as reusable text:list-style definitions. */
function listStyleXml(definition: AbstractNumbering): string {
  return xmlElement(
    "text:list-style",
    { "style:name": definition.reference },
    definition.levels.map(listLevelXml),
  );
}

function listLevelXml(level: NumberingLevel): string {
  const isBullet = level.format === "bullet";
  const template = level.text ?? "";
  const tokens = template.match(/%\d/g) ?? [];
  const prefix = template.split(/%\d/)[0] ?? "";
  const suffix = template.split(/%\d/).at(-1) ?? "";
  const indent = level.paragraph?.indent;
  const marginLeft = typeof indent?.left === "number" ? indent.left : undefined;
  const hanging = typeof indent?.hanging === "number" ? indent.hanging : undefined;
  const hasLabelAlignment =
    marginLeft !== undefined || hanging !== undefined || level.suffix !== undefined;
  const properties = hasLabelAlignment
    ? [
        xmlElement(
          "style:list-level-properties",
          { "style:list-level-position-and-space-mode": "label-alignment" },
          [
            xmlElement("style:list-level-label-alignment", {
              "text:label-followed-by": level.suffix ? SUFFIX_ODF[level.suffix] : undefined,
              "fo:margin-left": marginLeft !== undefined ? twipsToLength(marginLeft) : undefined,
              "fo:text-indent": hanging !== undefined ? twipsToLength(-hanging) : undefined,
            }),
          ],
        ),
      ]
    : [];
  return xmlElement(
    isBullet ? "text:list-level-style-bullet" : "text:list-level-style-number",
    {
      "text:level": level.level + 1,
      "style:num-format": isBullet
        ? undefined
        : level.format
          ? (NUM_FORMAT_ODF[level.format] ?? level.format)
          : undefined,
      "style:num-prefix": !isBullet && prefix ? prefix : undefined,
      "style:num-suffix": !isBullet && suffix ? suffix : undefined,
      "text:display-levels": !isBullet && tokens.length > 0 ? tokens.length : undefined,
      "text:bullet-char": isBullet ? template || "•" : undefined,
      "text:start-value": isBullet ? undefined : level.start,
    },
    properties,
  );
}

/** Docx document defaults → ODF style:default-style (paragraph family). */
function defaultStyleXml(styles: DocumentOptions["styles"]): string {
  const document = styles?.default?.document;
  if (!document) return "";
  const paragraph = document.paragraph ?? undefined;
  const run = document.run ?? undefined;
  const runAttributes = run
    ? {
        "fo:font-weight": run.bold ? "bold" : undefined,
        "fo:font-style": run.italic ? "italic" : undefined,
        "style:text-underline-style": run.underline?.type ? "solid" : undefined,
        "style:text-line-through-style": run.strike ? "solid" : undefined,
        "fo:font-size": typeof run.size === "number" ? `${run.size}pt` : undefined,
        "fo:color":
          typeof run.color === "string" && /^[0-9A-Fa-f]{6}$/.test(run.color)
            ? `#${run.color}`
            : undefined,
        "fo:font-family": typeof run.font === "string" ? run.font : undefined,
      }
    : {};
  const paragraphAttributes = paragraph
    ? {
        "fo:text-align": paragraph.alignment,
        "fo:margin-left":
          paragraph.indent?.left !== undefined ? twipsToLength(paragraph.indent.left) : undefined,
        "fo:margin-right":
          paragraph.indent?.right !== undefined ? twipsToLength(paragraph.indent.right) : undefined,
        "fo:margin-top":
          paragraph.spacing?.before !== undefined
            ? twipsToLength(paragraph.spacing.before)
            : undefined,
        "fo:margin-bottom":
          paragraph.spacing?.after !== undefined
            ? twipsToLength(paragraph.spacing.after)
            : undefined,
      }
    : {};
  const children = [
    ...(Object.values(paragraphAttributes).some((value) => value !== undefined)
      ? [xmlElement("style:paragraph-properties", paragraphAttributes)]
      : []),
    ...(Object.values(runAttributes).some((value) => value !== undefined)
      ? [xmlElement("style:text-properties", runAttributes)]
      : []),
  ];
  if (children.length === 0) return "";
  return xmlElement("style:default-style", { "style:family": "paragraph" }, children);
}

/** ODF style:default-style → docx document defaults. */
function parseDefaultStyle(files: OdfFiles): DocumentDefaults | undefined {
  const defaultStyle = childrenNamed(
    childNamed(readXml(files, "styles.xml"), "office:styles"),
    "style:default-style",
  ).find((style) => attributeString(style, "style:family") === "paragraph");
  if (!defaultStyle) return undefined;
  const paragraphProps = childNamed(defaultStyle, "style:paragraph-properties");
  const textProps = childNamed(defaultStyle, "style:text-properties");
  const indentLeft = lengthToTwips(attributeString(paragraphProps, "fo:margin-left"));
  const indentRight = lengthToTwips(attributeString(paragraphProps, "fo:margin-right"));
  const before = lengthToTwips(attributeString(paragraphProps, "fo:margin-top"));
  const after = lengthToTwips(attributeString(paragraphProps, "fo:margin-bottom"));
  const alignment = attributeString(paragraphProps, "fo:text-align");
  const paragraph = {
    ...(alignment ? { alignment } : {}),
    ...(indentLeft !== undefined || indentRight !== undefined
      ? {
          indent: {
            ...(indentLeft !== undefined ? { left: indentLeft } : {}),
            ...(indentRight !== undefined ? { right: indentRight } : {}),
          },
        }
      : {}),
    ...(before !== undefined || after !== undefined
      ? {
          spacing: {
            ...(before !== undefined ? { before } : {}),
            ...(after !== undefined ? { after } : {}),
          },
        }
      : {}),
  } as DocumentDefaults["paragraph"];
  const size = attributeString(textProps, "fo:font-size");
  const color = attributeString(textProps, "fo:color");
  const run = {
    bold: attributeString(textProps, "fo:font-weight") === "bold" || undefined,
    italic: attributeString(textProps, "fo:font-style") === "italic" || undefined,
    underline:
      attributeString(textProps, "style:text-underline-style") === "solid"
        ? { type: "single" as const }
        : undefined,
    strike: attributeString(textProps, "style:text-line-through-style") === "solid" || undefined,
    ...(size?.endsWith("pt") ? { size: Number(size.slice(0, -2)) } : {}),
    color: color?.startsWith("#") ? color.slice(1) : undefined,
    font: attributeString(textProps, "fo:font-family") || undefined,
  };
  const hasRun = Object.values(run).some((value) => value !== undefined);
  const hasParagraph = Object.keys(paragraph ?? {}).length > 0;
  if (!hasParagraph && !hasRun) return undefined;
  return {
    ...(hasParagraph ? { paragraph } : {}),
    ...(hasRun
      ? {
          run: Object.fromEntries(
            Object.entries(run).filter(([, value]) => value !== undefined),
          ) as DocumentDefaults["run"],
        }
      : {}),
  };
}

/** Footnote/endnote numbering config → ODF text:notes-configuration elements. */
function notesConfigurationXml(settings: DocumentOptions["settings"]): string {
  return (
    notesConfigXml("footnote", settings?.footnoteProperties) +
    notesConfigXml("endnote", settings?.endnoteProperties)
  );
}

function notesConfigXml(
  noteClass: "footnote" | "endnote",
  properties: FootnoteProperties | EndnoteProperties | undefined,
): string {
  if (!properties) return "";
  const numFmt =
    properties.format ?? (properties.numFmt ? NUM_FORMAT_ODF[properties.numFmt] : undefined);
  return xmlElement("text:notes-configuration", {
    "text:note-class": noteClass,
    "style:num-format": numFmt,
    "text:start-value": properties.numStart,
    "text:start-numbering-at":
      properties.numRestart === "continuous"
        ? "document"
        : properties.numRestart === "eachSect"
          ? "chapter"
          : properties.numRestart === "eachPage"
            ? "page"
            : undefined,
    "text:footnotes-position":
      properties.pos === "pageBottom"
        ? "page"
        : properties.pos === "beneathText"
          ? "text"
          : properties.pos === "sectEnd"
            ? "section"
            : properties.pos === "docEnd"
              ? "document"
              : undefined,
  });
}

/** ODF notes-configuration → docx footnote/endnote document properties. */
function parseNotesConfiguration(files: OdfFiles): {
  footnoteProperties?: FootnoteProperties;
  endnoteProperties?: EndnoteProperties;
} {
  const styles = readXml(files, "styles.xml");
  const result: { footnoteProperties?: FootnoteProperties; endnoteProperties?: EndnoteProperties } =
    {};
  for (const config of childrenNamed(
    childNamed(styles, "office:styles"),
    "text:notes-configuration",
  )) {
    const noteClass = attributeString(config, "text:note-class");
    if (noteClass !== "footnote" && noteClass !== "endnote") continue;
    const rawNumFormat = attributeString(config, "style:num-format");
    const mappedNumFormat = rawNumFormat ? NUM_FORMAT_DOCX[rawNumFormat] : undefined;
    const base = {
      pos: docxNotePosition(attributeString(config, "text:footnotes-position")),
      numFmt: mappedNumFormat,
      format: mappedNumFormat ? undefined : rawNumFormat,
      numStart: attributeNumber(config, "text:start-value"),
      numRestart: docxNumberRestart(attributeString(config, "text:start-numbering-at")),
    };
    if (noteClass === "footnote") result.footnoteProperties = base as FootnoteProperties;
    else result.endnoteProperties = base as EndnoteProperties;
  }
  return result;
}

function docxNotePosition(value: string | undefined): FootnoteProperties["pos"] {
  if (value === "page") return "pageBottom";
  if (value === "text") return "beneathText";
  if (value === "section") return "sectEnd";
  if (value === "document") return "docEnd";
  return undefined;
}

function docxNumberRestart(value: string | undefined): FootnoteProperties["numRestart"] {
  if (value === "document") return "continuous";
  if (value === "chapter") return "eachSect";
  if (value === "page") return "eachPage";
  return undefined;
}

function parsePageLayout(files: OdfFiles): SectionOptions["properties"] {
  const styles = readXml(files, "styles.xml");
  const layout = childNamed(
    childNamed(childNamed(styles, "office:automatic-styles"), "style:page-layout"),
    "style:page-layout-properties",
  );
  const width = lengthToTwips(attributeString(layout, "fo:page-width"));
  const height = lengthToTwips(attributeString(layout, "fo:page-height"));
  const orientation = attributeString(layout, "style:print-orientation");
  const margin = {
    top: lengthToTwips(attributeString(layout, "fo:margin-top")),
    right: lengthToTwips(attributeString(layout, "fo:margin-right")),
    bottom: lengthToTwips(attributeString(layout, "fo:margin-bottom")),
    left: lengthToTwips(attributeString(layout, "fo:margin-left")),
  };
  const hasMargin = Object.values(margin).some((value) => value !== undefined);
  if (width === undefined && height === undefined && !hasMargin && !orientation) return undefined;
  return {
    pageSize:
      width !== undefined || height !== undefined || orientation
        ? { width, height, orientation: orientation as "portrait" | "landscape" | undefined }
        : false,
    pageMargin: hasMargin ? margin : undefined,
  };
}

function twipsToLength(value: number | string): string {
  return emuToLength(typeof value === "number" ? value * 635 : lengthToEmu(value));
}

function lengthToTwips(value: string | undefined): number | undefined {
  const emu = lengthToEmu(value);
  return emu === undefined ? undefined : Math.round(emu / 635);
}

function blockXml(
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
      "text:id": child.bookmarkStart.id ? `bookmark-${child.bookmarkStart.id}` : undefined,
      "text:name": child.bookmarkStart.name,
    });
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

function decodeStyleName(value: string): string {
  return value.replace(/\\([\\;])/g, "$1");
}

function normalizeParagraph(input: string | ParagraphOptions): ParagraphOptions {
  return typeof input === "string" ? { text: input } : input;
}

function listParagraphLevel(
  child: SectionChild,
): { level: number; ordered: boolean; reference?: string } | undefined {
  if (!("paragraph" in child)) return undefined;
  const options = normalizeParagraph(child.paragraph);
  if (options.heading) return undefined;
  if (options.bullet?.level !== undefined) return { level: options.bullet.level, ordered: false };
  const numbering = options.numbering;
  if (typeof numbering === "object" && "reference" in numbering)
    return { level: numbering.level ?? 0, ordered: true, reference: numbering.reference };
  return undefined;
}

function listXml(
  group: SectionChild[],
  info: { level: number; ordered: boolean; reference?: string },
  styles: string[],
  images: OdtImage[],
  notes: NotesContext,
  numbering: DocumentOptions["numbering"],
  charts: OdtChart[],
): string {
  const styleName = addListStyle(styles, info, numbering);
  // ODF nesting is 1-based: list level 0 renders as a single text:list,
  // level 1 nests one text:list inside the first list-item, and so on.
  const items = group
    .map((child) =>
      xmlElement("text:list-item", undefined, [
        blockXml(child, styles, images, notes, numbering, charts),
      ]),
    )
    .join("");
  let xml = items;
  for (let depth = 0; depth < info.level; depth += 1) {
    xml = xmlElement("text:list-item", undefined, [
      xmlElement("text:list", { "text:style-name": styleName }, [xml]),
    ]);
  }
  return xmlElement("text:list", { "text:style-name": styleName }, [xml]);
}

function addListStyle(
  styles: string[],
  info: { level: number; ordered: boolean; reference?: string },
  numbering: DocumentOptions["numbering"],
): string {
  const definition = info.ordered
    ? numbering?.abstractNumberings?.find((entry) => entry.reference === info.reference)
    : undefined;
  if (definition) {
    if (
      !styles.some((xml) => xml.startsWith(`<text:list-style style:name="${definition.reference}"`))
    )
      styles.push(listStyleXml(definition));
    return definition.reference;
  }
  const name = `L${styles.length + 1}`;
  styles.push(
    xmlElement("text:list-style", { "style:name": name }, [
      info.ordered
        ? xmlElement("text:list-level-style-number", { "text:level": 1, "style:num-format": "1" }, [
            xmlElement("style:list-level-properties", {
              "style:list-level-position-and-space-mode": "label-alignment",
            }),
          ])
        : xmlElement("text:list-level-style-bullet", { "text:level": 1, "text:bullet-char": "•" }, [
            xmlElement("style:list-level-properties", {
              "style:list-level-position-and-space-mode": "label-alignment",
            }),
            xmlElement("style:text-properties", { "fo:font-family": "OpenSymbol" }),
          ]),
    ]),
  );
  return name;
}

/** List styles record whether each name resolves to a numbering run. */
function parseListStyles(container: Element | undefined): Map<string, boolean> {
  const result = new Map<string, boolean>();
  for (const style of childrenNamed(container, "text:list-style")) {
    const name = attributeString(style, "style:name") ?? "";
    result.set(name, childrenNamed(style, "text:list-level-style-number").length > 0);
  }
  return result;
}

/** text:list-style definitions → abstract numberings keyed by style name. */
function parseListNumberings(container: Element | undefined): AbstractNumbering[] {
  const definitions: AbstractNumbering[] = [];
  for (const style of childrenNamed(container, "text:list-style")) {
    const name = attributeString(style, "style:name");
    if (!name) continue;
    const levels = [
      ...childrenNamed(style, "text:list-level-style-number"),
      ...childrenNamed(style, "text:list-level-style-bullet"),
    ].map(parseListLevel);
    if (levels.length > 0) definitions.push({ reference: name, levels });
  }
  return definitions;
}

function parseListLevel(element: Element): NumberingLevel {
  const isBullet = element.name === "text:list-level-style-bullet";
  const oneBased = attributeNumber(element, "text:level") ?? 1;
  const display = attributeNumber(element, "text:display-levels") ?? 1;
  const prefix = attributeString(element, "style:num-prefix") ?? "";
  const suffix = attributeString(element, "style:num-suffix") ?? "";
  const tokens = Array.from(
    { length: display },
    (_, index) => `%${oneBased - display + 1 + index}`,
  ).join(".");
  const rawFormat = attributeString(element, "style:num-format");
  const start = attributeNumber(element, "text:start-value");
  const labelAlignment = childNamed(
    childNamed(element, "style:list-level-properties"),
    "style:list-level-label-alignment",
  );
  const marginLeft = lengthToTwips(attributeString(labelAlignment, "fo:margin-left"));
  const textIndent = lengthToTwips(attributeString(labelAlignment, "fo:text-indent"));
  const followedBy = attributeString(labelAlignment, "text:label-followed-by");
  const indent: { left?: number; hanging?: number } = {};
  if (marginLeft !== undefined) indent.left = marginLeft;
  if (textIndent !== undefined && textIndent < 0) indent.hanging = -textIndent;
  return {
    level: oneBased - 1,
    format: isBullet
      ? "bullet"
      : ((NUM_FORMAT_DOCX[rawFormat ?? ""] ?? rawFormat) as NumberingLevel["format"]),
    text: isBullet
      ? (attributeString(element, "text:bullet-char") ?? "•")
      : `${prefix}${tokens}${suffix}`,
    ...(start !== undefined ? { start } : {}),
    ...(followedBy && SUFFIX_DOCX[followedBy]
      ? { suffix: SUFFIX_DOCX[followedBy] as NumberingLevel["suffix"] }
      : {}),
    ...(Object.keys(indent).length > 0 ? { paragraph: { indent } } : {}),
  } as NumberingLevel;
}

function paragraphXml(
  options: ParagraphOptions,
  styles: string[],
  images: OdtImage[],
  notes: NotesContext,
  numbering: DocumentOptions["numbering"],
  charts: OdtChart[],
): string {
  const children = runXml(options, styles, images, notes, numbering, charts);
  if (
    options.children?.length &&
    options.children.every((child) => typeof child === "object" && "formField" in child)
  ) {
    return "";
  }
  const alignment = typeof options.alignment === "string" ? options.alignment : undefined;
  const styleName =
    alignment || options.pageBreakBefore || options.tabStops?.length
      ? addParagraphStyle(
          { alignment, pageBreakBefore: options.pageBreakBefore, tabStops: options.tabStops },
          styles,
        )
      : undefined;
  const heading = /^Heading([1-9])$/.exec(options.heading ?? "");
  const attributes = {
    "text:style-name": styleName,
    "text:outline-level": heading ? Number(heading[1]) : undefined,
  };
  return xmlElement(heading ? "text:h" : "text:p", attributes, children);
}

function runXml(
  options: ParagraphOptions,
  styles: string[],
  images: OdtImage[],
  notes: NotesContext,
  numbering: DocumentOptions["numbering"],
  charts: OdtChart[],
): string[] {
  if (options.text !== undefined && options.children === undefined) {
    return [spacesXml(options.text)];
  }
  return (options.children ?? []).map((child) => {
    if (typeof child === "string") return `<text:span>${spacesXml(child)}</text:span>`;
    if ("pageBreak" in child) return "<text:soft-page-break/>";
    if ("columnBreak" in child) return "<text:line-break/>";
    if ("break" in child) return lineBreakXml(child as RunOptions);
    if ("bookmark" in child) {
      return xmlElement("text:bookmark", {
        "text:name": (child as { bookmark: { name: string } }).bookmark.name,
      });
    }
    if ("bookmarkEnd" in child) {
      return xmlElement("text:bookmark-end", {
        "text:name": notes.bookmarkNames.get(child.bookmarkEnd.id) ?? "",
      });
    }
    if ("simpleField" in child) {
      return simpleFieldXml(child.simpleField, notes);
    }
    if ("comment" in child) {
      return annotationXml(
        child.comment.author,
        child.comment.date,
        child.comment.children.flatMap((commentChild) =>
          typeof commentChild === "string"
            ? [commentChild]
            : "text" in commentChild && typeof commentChild.text === "string"
              ? [commentChild.text]
              : "paragraph" in commentChild
                ? [
                    typeof commentChild.paragraph === "string"
                      ? commentChild.paragraph
                      : ((commentChild.paragraph as { text?: string } | undefined)?.text ?? ""),
                  ]
                : [],
        ),
      );
    }
    if ("formField" in child) return formFieldXml(child.formField, notes);
    if ("insertion" in child) {
      return insertionXml(child.insertion, notes, styles, images, numbering, charts);
    }
    if ("deletion" in child) {
      const id = child.deletion.id ?? notes.revisions.size + 1;
      notes.revisions.set(id, {
        id,
        author: child.deletion.author,
        date: child.deletion.date,
        children: child.deletion.children,
      });
      return xmlElement("text:change", { "text:change-id": `rev${id}` });
    }
    if ("footnoteReference" in child) {
      return noteXml(
        (child as { footnoteReference: number | { id: number } }).footnoteReference,
        notes,
        "footnote",
        styles,
        images,
        numbering,
        charts,
      );
    }
    if ("endnoteReference" in child) {
      return noteXml(
        (child as { endnoteReference: number | { id: number } }).endnoteReference,
        notes,
        "endnote",
        styles,
        images,
        numbering,
        charts,
      );
    }
    if ("hyperlink" in child) {
      const link = (
        child as {
          hyperlink: {
            url?: string;
            anchor?: string;
            tooltip?: string;
            targetFrame?: string;
            children?: Array<string | { text?: string }>;
          };
          text?: string;
        }
      ).hyperlink;
      const href = link.url ?? (link.anchor !== undefined ? `#${link.anchor}` : undefined);
      if (href === undefined) return "";
      const inner = link.children
        ? link.children
            .map((c) => (typeof c === "string" ? spacesXml(c) : spacesXml(c.text ?? "")))
            .join("")
        : spacesXml((child as { text?: string }).text ?? "");
      return xmlElement(
        "text:a",
        {
          "xlink:type": "simple",
          "xlink:href": href,
          "office:title": link.tooltip,
          "office:target-frame-name": link.targetFrame,
        },
        [inner],
      );
    }
    if ("text" in child) {
      const run = child as RunOptions;
      const styleName = addCharacterStyle(characterProperties(run), styles);
      return (
        lineBreakXml(run) +
        xmlElement("text:span", { "text:style-name": styleName }, [spacesXml(run.text ?? "")])
      );
    }
    if ("tab" in child) return "<text:tab/>";
    if ("picture" in child)
      return pictureFrameXml((child as { picture: PictureOptions }).picture, images);
    if ("wpsShape" in child)
      return wpsShapeFrameXml((child as { wpsShape: ShapeOptions }).wpsShape, styles);
    if ("chart" in child) {
      return chartFrameXml((child as { chart: ChartOptions }).chart, charts);
    }
    return "";
  });
}

function simpleFieldXml(field: CanonicalSimpleField, notes: NotesContext): string {
  const instruction = field.instruction.trim();
  let match = /^REF\s+(?<name>[^\s]+)(?:\s+(?<switch>\\[a-zA-Z]))?$/.exec(instruction);
  if (match) {
    return bookmarkReferenceXml({
      name: decodeFieldArgument(match.groups?.name ?? ""),
      referenceFormat: odfReferenceFormat(match.groups?.switch),
    });
  }
  match = /^SEQ\s+(?<name>[^\s]+)$/.exec(instruction);
  if (match) {
    notes.sequenceNames.add(decodeFieldArgument(match.groups?.name ?? ""));
    return sequenceXml({
      name: decodeFieldArgument(match.groups?.name ?? ""),
      display: field.cachedValue,
    });
  }
  match = /^STYLEREF\s+(?<level>[1-9])\s+\\n$/.exec(instruction);
  if (match) {
    return chapterXml({
      display: field.cachedValue,
      outlineLevel: Number(match.groups?.level),
    });
  }
  match = /^VARIABLE\s+(?<name>[^\s]+?)(?:\s*=\s*(?<formula>[\s\S]+))?$/.exec(instruction);
  if (match) {
    const name = decodeFieldArgument(match.groups?.name ?? "");
    const valueType = Number.isFinite(Number(field.cachedValue)) ? "float" : "string";
    notes.variableTypes.set(name, valueType);
    return variableSetXml({
      name,
      valueType,
      value: field.cachedValue,
      display: field.cachedValue,
      formula: match.groups?.formula,
    });
  }
  match = /^IF\s+(?<condition>[\s\S]+)\s+"(?<content>[^"]*)"\s+""$/.exec(instruction);
  if (match) {
    return hiddenTextXml({
      condition: decodeFieldArgument(match.groups?.condition ?? ""),
      content: match.groups?.content ?? "",
      hidden: field.cachedValue === undefined ? true : undefined,
    });
  }
  if (instruction.startsWith("=")) {
    return expressionXml({
      formula: instruction.slice(1).trim(),
      valueType: Number.isFinite(Number(field.cachedValue)) ? "float" : "string",
      value: field.cachedValue,
      display: field.cachedValue,
    });
  }
  throw new OdtParseError(
    `content.xml: field instruction has no ODT mapping: ${instruction}`,
    "content.xml",
    "/office:document-content/office:body/office:text",
    "text:field",
    "unsupported field instruction",
  );
}

function decodeFieldArgument(value: string): string {
  return value.replace(/\\([\\"])/g, "$1");
}

function odfReferenceFormat(fieldSwitch: string | undefined): string | undefined {
  if (fieldSwitch === "\\p") return "page";
  if (fieldSwitch === "\\r") return "number";
  return undefined;
}

function fieldReferenceSwitch(format: string | undefined): string {
  if (format === "page") return " \\p";
  if (format === "number") return " \\r";
  return "";
}

function formFieldXml(field: FormFieldOptions, notes: NotesContext): string {
  const common = new Set(["name", "enabled", "tabIndex"]);
  const variants = Object.keys(field).filter((key) => !common.has(key));
  const variant = variants.find((key) => ["checkBox", "dropDownList", "textInput"].includes(key));
  const unsupported = variants.filter((key) => key !== variant);
  if (!variant || unsupported.length) {
    throw unsupportedOdtValue(
      "form field",
      `property ${(variant ? unsupported : variants).join(", ")} has no ODT mapping`,
    );
  }
  if (variant === "checkBox") {
    const options = field.checkBox ?? {};
    const allowed = new Set(["checked", "default"]);
    const optionExtras = Object.keys(options).filter((key) => !allowed.has(key));
    if (optionExtras.length) throw unsupportedOdtValue("checkbox", optionExtras.join(", "));
    if (options.default !== undefined && options.default !== options.checked) {
      throw unsupportedOdtValue("checkbox", "default state has no ODT mapping");
    }
    notes.forms.push({ ...field });
    return "";
  }
  if (variant === "dropDownList") {
    const options = field.dropDownList ?? {};
    const allowed = new Set(["entries", "result"]);
    const optionExtras = Object.keys(options).filter((key) => !allowed.has(key));
    if (optionExtras.length) throw unsupportedOdtValue("dropdown", optionExtras.join(", "));
    notes.forms.push({ ...field });
    return "";
  }
  const options = field.textInput ?? {};
  const allowed = new Set(["type", "default", "value", "maxLength"]);
  const optionExtras = Object.keys(options).filter((key) => !allowed.has(key));
  if (optionExtras.length || (options.type !== undefined && options.type !== "regular")) {
    throw unsupportedOdtValue("text input", optionExtras.join(", ") || options.type || "type");
  }
  notes.forms.push({ ...field });
  return "";
}

function insertionXml(
  value: ChangedProperties & { children?: TrackChangeChild[] },
  notes: NotesContext,
  styles: string[],
  images: OdtImage[],
  numbering: DocumentOptions["numbering"],
  charts: OdtChart[],
): string {
  const id = value.id ?? notes.revisions.size + 1;
  notes.revisions.set(id, { id, author: value.author, date: value.date });
  const children = runXml({ children: value.children }, styles, images, notes, numbering, charts);
  return (
    xmlElement("text:change-start", { "text:change-id": `rev${id}` }) +
    children.join("") +
    xmlElement("text:change-end", { "text:change-id": `rev${id}` })
  );
}

function declarationsXml(notes: NotesContext): string {
  const forms = notes.forms.map((field, index) => {
    const id = `field${index + 1}`;
    const common = {
      "form:id": id,
      "form:name": field.name,
      "form:disabled": field.enabled === false ? true : undefined,
      "form:tab-index": field.tabIndex,
    };
    if (field.checkBox) {
      return xmlElement("form:form", { "form:name": field.name ?? id }, [
        xmlElement("form:checkbox", {
          ...common,
          "form:current-state": field.checkBox.checked ? "checked" : "unchecked",
        }),
      ]);
    }
    if (field.dropDownList) {
      return xmlElement("form:form", { "form:name": field.name ?? id }, [
        xmlElement(
          "form:listbox",
          common,
          field.dropDownList.entries.map((label, optionIndex) =>
            xmlElement(
              "form:option",
              {
                "form:selected": field.dropDownList?.result === optionIndex || undefined,
              },
              [escapeText(label)],
            ),
          ),
        ),
      ]);
    }
    if (!field.textInput) throw unsupportedOdtValue("form field", "control type is required");
    return xmlElement("form:form", { "form:name": field.name ?? id }, [
      xmlElement("form:text", {
        ...common,
        "form:current-value": field.textInput.value ?? field.textInput.default,
        "form:max-length": field.textInput.maxLength,
      }),
    ]);
  });
  const sequenceDecls = notes.sequenceNames.size
    ? xmlElement(
        "text:sequence-decls",
        undefined,
        [...notes.sequenceNames].map((name) =>
          xmlElement("text:sequence-decl", { "text:name": name }),
        ),
      )
    : "";
  const variableDecls = notes.variableTypes.size
    ? xmlElement(
        "text:variable-decls",
        undefined,
        [...notes.variableTypes].map(([name, valueType]) =>
          xmlElement("text:variable-decl", { "text:name": name, "office:value-type": valueType }),
        ),
      )
    : "";
  const revisions = notes.revisions.size
    ? xmlElement(
        "text:tracked-changes",
        undefined,
        [...notes.revisions.entries()].map(([, revision]) =>
          xmlElement("text:changed-region", { "text:id": `rev${revision.id}` }, [
            xmlElement(`text:${revision.children ? "deletion" : "insertion"}`, undefined, [
              xmlElement("office:change-info", undefined, [
                xmlElement("dc:creator", undefined, [escapeText(revision.author)]),
                xmlElement("dc:date", undefined, [String(revision.date)]),
              ]),
              ...(revision.children ?? []).map((child) =>
                xmlElement(
                  "text:p",
                  undefined,
                  runXml({ children: [child] }, [], [], notes, undefined, []),
                ),
              ),
            ]),
          ]),
        ),
      )
    : "";
  const formsXml = forms.length ? xmlElement("office:forms", undefined, forms) : "";
  return formsXml + sequenceDecls + variableDecls + revisions;
}

function unsupportedOdtValue(name: string, reason: string): OdtParseError {
  return new OdtParseError(
    `content.xml: ${name} ${reason}`,
    "content.xml",
    "/office:document-content/office:body/office:text",
    name,
    "no canonical ODT mapping",
  );
}

function malformedOdtElement(name: string, reason: string): OdtParseError {
  const path = `/office:document-content/office:body/office:text/${name}`;
  return new OdtParseError(`content.xml: ${name}: ${reason}`, "content.xml", path, name, reason);
}

function variableSetXml(value: {
  name: string;
  valueType: "float" | "string";
  value?: string;
  display?: string;
  formula?: string;
}): string {
  return xmlElement(
    "text:variable-set",
    {
      "text:name": value.name,
      "office:value-type": value.valueType,
      "office:value": value.value,
      "text:display": value.display,
      "text:formula": value.formula,
    },
    value.display === undefined ? [] : [escapeText(value.display)],
  );
}

function bookmarkReferenceXml(value: { name: string; referenceFormat?: string }): string {
  return xmlElement(
    "text:bookmark-ref",
    {
      "text:reference-format": value.referenceFormat,
    },
    [escapeText(value.name)],
  );
}

function chapterXml(value: { display?: string; outlineLevel?: number }): string {
  return xmlElement(
    "text:chapter",
    {
      "text:display": value.display,
      "text:outline-level": value.outlineLevel,
    },
    value.display === undefined ? [] : [escapeText(value.display)],
  );
}

function expressionXml(value: {
  formula?: string;
  valueType?: "float" | "string";
  value?: string;
  display?: string;
}): string {
  return xmlElement(
    "text:expression",
    {
      "text:formula": value.formula,
      "office:value-type": value.valueType,
      "office:value": value.value,
      "text:display": value.display,
    },
    value.display === undefined ? [] : [escapeText(value.display)],
  );
}

function sequenceXml(value: {
  name: string;
  referenceName?: string;
  display?: string;
  formula?: string;
}): string {
  return xmlElement(
    "text:sequence",
    {
      "text:name": value.name,
      "text:ref-name": value.referenceName,
      "text:display": value.display,
      "text:formula": value.formula,
    },
    value.display === undefined ? [] : [escapeText(value.display)],
  );
}

function hiddenTextXml(value: {
  condition: string;
  content: string;
  hidden?: boolean;
  fixed?: boolean;
}): string {
  return xmlElement(
    "text:hidden-text",
    {
      "text:condition": value.condition,
      "text:is-hidden": value.hidden,
      "text:is-fixed": value.fixed,
    },
    [escapeText(value.content)],
  );
}

function annotationXml(
  author: string | undefined,
  date: Date | string | null | undefined,
  paragraphs: string[],
): string {
  return xmlElement(
    "office:annotation",
    {
      "dc:creator": author,
      "dc:date": date instanceof Date ? date.toISOString() : (date ?? undefined),
    },
    paragraphs.map((paragraph) => xmlElement("text:p", undefined, [escapeText(paragraph)])),
  );
}

function indexXml(value: Extract<SectionChild, { toc: unknown }>["toc"]): string {
  return xmlElement(
    "text:illustration-index",
    {
      "text:name": value.alias,
    },
    [
      xmlElement("text:illustration-index-source", {
        "text:index-scope": "document",
        "text:use-caption": value.captionLabelIncludingNumbers !== undefined,
      }),
      xmlElement(
        "text:index-body",
        undefined,
        value.entries
          ?.map(sectionChildText)
          .map((text) => xmlElement("text:p", undefined, [escapeText(text)])),
      ),
    ],
  );
}

function bibliographyIndexXml(value: Extract<SectionChild, { sdt: unknown }>["sdt"]): string {
  return xmlElement(
    "text:bibliography",
    {
      "text:name": value.properties.alias,
    },
    [
      xmlElement("text:bibliography-source", {}, [
        xmlElement("text:bibliography-entry-template", {
          "text:bibliography-type": "article",
          "text:style-name": "Index",
        }),
      ]),
      xmlElement(
        "text:index-body",
        undefined,
        value.children
          ?.map(sectionChildText)
          .map((text) => xmlElement("text:p", undefined, [escapeText(text)])),
      ),
    ],
  );
}

function sectionChildText(child: SectionChild): string {
  if (typeof child === "string") return child;
  if ("paragraph" in child)
    return typeof child.paragraph === "string" ? child.paragraph : (child.paragraph.text ?? "");
  return "";
}

/** Inline chart renders as a draw:frame + draw:object pointing at the subdocument. */
function chartFrameXml(chart: ChartOptions, charts: OdtChart[]): string {
  const path = `Object ${charts.length + 1}`;
  charts.push({ path, chart });
  return xmlElement(
    "draw:frame",
    {
      "text:anchor-type": "as-char",
      "svg:width": emuToLength(chart.transformation.width),
      "svg:height": emuToLength(chart.transformation.height),
    },
    [xmlElement("draw:object", { "xlink:href": `./${path}`, "xlink:type": "simple" })],
  );
}

/** Inline shape renders as a positioned draw:custom-shape with preset geometry. */
function wpsShapeFrameXml(shape: ShapeOptions, styles: string[]): string {
  const geometry = typeof shape.geometry === "string" ? { preset: shape.geometry } : shape.geometry;
  const offset = shape.transformation.offset;
  return xmlElement(
    "draw:custom-shape",
    {
      "text:anchor-type": "as-char",
      "draw:style-name": addShapeStyle(shape, styles),
      "svg:x":
        offset?.left !== undefined
          ? typeof offset.left === "number"
            ? emuToLength(offset.left)
            : offset.left
          : undefined,
      "svg:y":
        offset?.top !== undefined
          ? typeof offset.top === "number"
            ? emuToLength(offset.top)
            : offset.top
          : undefined,
      "svg:width":
        typeof shape.transformation.width === "number"
          ? emuToLength(shape.transformation.width)
          : shape.transformation.width,
      "svg:height":
        typeof shape.transformation.height === "number"
          ? emuToLength(shape.transformation.height)
          : shape.transformation.height,
      "draw:name": shape.altText?.name,
    },
    [
      xmlElement(
        "draw:enhanced-geometry",
        { "draw:type": geometry?.preset ? presetGeometryOdf(geometry.preset) : undefined },
        [],
      ),
    ],
  );
}

/** Shape fill and outline land in a reusable graphic style. */
function addShapeStyle(shape: ShapeOptions, styles: string[]): string | undefined {
  return pushShapeStyle(shape.fill, shape.outline, styles);
}

/** Note reference renders inline as text:note carrying its body paragraphs. */
function noteXml(
  reference: number | { id: number },
  notes: NotesContext,
  noteClass: "footnote" | "endnote",
  styles: string[],
  images: OdtImage[],
  numbering: DocumentOptions["numbering"],
  charts: OdtChart[],
): string {
  const id = typeof reference === "number" ? reference : reference.id;
  const children = (noteClass === "endnote" ? notes.endnotes : notes.footnotes).get(id) ?? [];
  return xmlElement("text:note", { "text:id": `${noteClass}${id}`, "text:note-class": noteClass }, [
    xmlElement("text:note-citation", undefined, [String(id)]),
    xmlElement("text:note-body", undefined, [
      blocksXml(
        children.map((noteChild) =>
          typeof noteChild === "string" || !("paragraph" in noteChild || "table" in noteChild)
            ? ({ paragraph: noteChild } as SectionChild)
            : (noteChild as SectionChild),
        ),
        styles,
        images,
        notes,
        numbering,
        charts,
      ),
    ]),
  ]);
}

/** Run breaks render as text:line-break (ODF has no w:br/@clear equivalent). */
function lineBreakXml(run: RunOptions): string {
  if (!run.break) return "";
  const count = typeof run.break === "number" ? run.break : (run.break.count ?? 1);
  return "<text:line-break/>".repeat(Math.max(0, count));
}

/** Inline picture renders as a character-anchored draw:frame + draw:image. */
function pictureFrameXml(picture: PictureOptions, images: OdtImage[]): string {
  const raster = picture.type === "svg" ? picture.fallback : picture;
  if (raster.data === undefined) return "";
  const data = toUint8Array(raster.data);
  const path = `Pictures/picture${images.length + 1}.${raster.type}`;
  images.push({ path, data });
  return xmlElement(
    "draw:frame",
    {
      "text:anchor-type": "as-char",
      "svg:width": emuToLength(picture.transformation.width),
      "svg:height": emuToLength(picture.transformation.height),
      "draw:name": picture.altText?.name,
    },
    [xmlElement("draw:image", { "xlink:href": path })],
  );
}

function characterProperties(run: RunOptions): CharacterProperties {
  return {
    bold: run.bold,
    italic: run.italic,
    underline: run.underline ? "single" : undefined,
    strike: run.strike,
    size: run.size,
    color:
      typeof run.color === "string" && /^[0-9A-Fa-f]{6}$/.test(run.color) ? run.color : undefined,
    font: typeof run.font === "string" ? run.font : undefined,
  };
}

function addParagraphStyle(
  properties: { alignment?: string; pageBreakBefore?: boolean; tabStops?: TabStop[] },
  styles: string[],
): string {
  const name = `P${styles.length + 1}`;
  styles.push(
    xmlElement(
      "style:style",
      { "style:name": name, "style:family": "paragraph", "style:parent-style-name": "Standard" },
      [
        xmlElement(
          "style:paragraph-properties",
          {
            "fo:text-align": properties.alignment,
            "fo:break-before": properties.pageBreakBefore ? "page" : undefined,
          },
          properties.tabStops?.length
            ? [xmlElement("style:tab-stops", undefined, properties.tabStops.map(tabStopXml))]
            : undefined,
        ),
      ],
    ),
  );
  return name;
}

/** Tab stop → ODF style:tab-stop; decimal maps to char with a dot. */
function tabStopXml(tab: TabStop): string {
  if (
    tab.type !== "left" &&
    tab.type !== "center" &&
    tab.type !== "right" &&
    tab.type !== "decimal"
  )
    return "";
  return xmlElement("style:tab-stop", {
    "style:position": twipsToLength(tab.position),
    "style:type": tab.type === "left" ? undefined : tab.type === "decimal" ? "char" : tab.type,
    "style:char": tab.type === "decimal" ? "." : undefined,
    "style:leader-text":
      tab.leader === "dot"
        ? "."
        : tab.leader === "hyphen"
          ? "-"
          : tab.leader === "underscore"
            ? "_"
            : tab.leader === "middleDot"
              ? "·"
              : undefined,
  });
}

function addCharacterStyle(properties: CharacterProperties, styles: string[]): string | undefined {
  if (Object.values(properties).every((value) => value === undefined)) return undefined;
  const name = `C${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "text" }, [
      xmlElement("style:text-properties", {
        "fo:font-weight": properties.bold ? "bold" : undefined,
        "fo:font-style": properties.italic ? "italic" : undefined,
        "style:text-underline-style": properties.underline ? "solid" : undefined,
        "style:text-line-through-style": properties.strike ? "solid" : undefined,
        "fo:font-size": properties.size ? `${properties.size}pt` : undefined,
        "fo:color": properties.color ? `#${properties.color}` : undefined,
        "fo:font-family": properties.font,
      }),
    ]),
  );
  return name;
}
function parseStyles(container: Element | undefined): StyleMap {
  const result: StyleMap = new Map();
  for (const style of childrenNamed(container, "style:style")) {
    const name = attributeString(style, "style:name") ?? "";
    const paragraph = childNamed(style, "style:paragraph-properties");
    const tabStops = childrenNamed(childNamed(paragraph, "style:tab-stops"), "style:tab-stop").map(
      parseTabStop,
    );
    const character = childNamed(style, "style:text-properties");
    const columnWidth = attributeString(
      childNamed(style, "style:table-column-properties"),
      "style:column-width",
    );
    const color = attributeString(character, "fo:color");
    const size = attributeString(character, "fo:font-size");
    result.set(name, {
      alignment: attributeString(paragraph, "fo:text-align"),
      pageBreakBefore: attributeString(paragraph, "fo:break-before") === "page",
      tabStops: tabStops.length > 0 ? tabStops : undefined,
      columnWidth: columnWidth?.endsWith("cm") ? Number(columnWidth.slice(0, -2)) * 567 : undefined,
      character: {
        bold: attributeString(character, "fo:font-weight") === "bold",
        italic: attributeString(character, "fo:font-style") === "italic",
        underline:
          attributeString(character, "style:text-underline-style") === "solid"
            ? "single"
            : undefined,
        strike: attributeString(character, "style:text-line-through-style") === "solid",
        size: size?.endsWith("pt") ? Number(size.slice(0, -2)) : undefined,
        color: color?.startsWith("#") ? color.slice(1) : undefined,
        font: attributeString(character, "fo:font-family"),
      },
    });
  }
  return result;
}

function parseBlock(element: Element, context: ParseContext): SectionChild {
  if (element.name === "table:table")
    return parseTable(
      element,
      (name) => context.styles.get(name)?.columnWidth,
      (child) => parseBlock(child, context),
    );
  const paragraph = parseParagraph(element, context);
  return { paragraph };
}

function parseBlocks(
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

function blockChildren(children: SectionChild[]): BlockContentChild[] {
  return children.filter((child): child is BlockContentChild => !("altChunk" in child));
}

function unknownOdtElement(element: Element): OdtParseError {
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

function nextBookmarkId(context: ParseContext): number {
  return Math.max(0, ...context.bookmarkIds.values()) + 1;
}

function shapeBlock(shape: ShapeOptions): SectionChild {
  return { paragraph: { children: [{ wpsShape: shape }] } };
}

function withPageBreakBefore(child: SectionChild): SectionChild {
  if (!("paragraph" in child) || typeof child.paragraph === "string") return child;
  return { paragraph: { ...child.paragraph, pageBreakBefore: true } };
}

function encodeSectionTag(styleName: string): string {
  return `odf:text-section;style=${styleName.replace(/([\\;])/g, "\\$1")}`;
}

function parseParagraph(element: Element, context: ParseContext): ParagraphOptions {
  const style = context.styles.get(attributeString(element, "text:style-name") ?? "");
  const headingLevel = attributeNumber(element, "text:outline-level");
  const runs = parseRuns(element, context);
  const result: ParagraphOptions = {};
  if (style?.alignment) result.alignment = style.alignment as ParagraphOptions["alignment"];
  if (style?.pageBreakBefore) result.pageBreakBefore = true;
  if (style?.tabStops) result.tabStops = style.tabStops;
  if (headingLevel && headingLevel <= 6) {
    result.heading = `Heading${headingLevel}` as ParagraphOptions["heading"];
  }
  if (result.heading && context.outline) {
    const outlineLevel = Number(/Heading([1-9])$/.exec(result.heading)?.[1]);
    if (outlineLevel && context.outline.levels.some((level) => level.level === outlineLevel - 1))
      result.numbering = { reference: context.outline.reference, level: outlineLevel - 1 };
  }
  if (runs.length === 1 && typeof runs[0] === "string") result.text = runs[0];
  else if (runs.length > 0) result.children = runs as ParagraphOptions["children"];
  return result;
}

function parseRuns(
  element: Element,
  context: ParseContext,
): (string | RunOptions | ParagraphChild)[] {
  const nodes = element.elements ?? [];
  let resumeAfter = -1;
  const runs = (element.elements ?? []).flatMap(
    (child, index): (string | RunOptions | ParagraphChild)[] => {
      if (index <= resumeAfter) return [];
      if (child.type === "text") return [String(child.text ?? "")];
      if (child.name === "text:line-break") return [{ break: 1 }];
      if (child.name === "text:soft-page-break") return [{ pageBreak: true }];
      if (child.name === "text:s") {
        const count = attributeNumber(child, "text:c") ?? 1;
        return [" ".repeat(count)];
      }
      if (child.name === "text:a") {
        const href = attributeString(child, "xlink:href") ?? "";
        return [
          {
            hyperlink: {
              ...(href.startsWith("#") ? { anchor: href.slice(1) } : { url: href }),
              tooltip: attributeString(child, "office:title"),
              targetFrame: attributeString(child, "office:target-frame-name"),
              children: parseRuns(child, context),
            },
          },
        ];
      }
      if (child.name === "text:note") return parseNote(child, context);
      if (child.name === "text:bookmark") {
        return [{ bookmark: { name: attributeString(child, "text:name") ?? "" } }];
      }
      if (child.name === "text:bookmark-start" || child.name === "text:reference-mark-start") {
        const name = attributeString(child, "text:name") ?? "";
        if (!name || context.bookmarkIds.has(name)) throw unknownOdtElement(child);
        return [{ bookmarkStart: { id: bookmarkRunId(context, name), name } }];
      }
      if (child.name === "text:bookmark-end" || child.name === "text:reference-mark-end") {
        const name = attributeString(child, "text:name") ?? "";
        const id = context.bookmarkIds.get(name);
        if (!name || id === undefined) throw unknownOdtElement(child);
        return [{ bookmarkEnd: { id } }];
      }
      if (child.name === "text:change-start") {
        const id = attributeString(child, "text:change-id") ?? "";
        const endIndex = nodes.findIndex(
          (node, nodeIndex) =>
            nodeIndex > index &&
            node.name === "text:change-end" &&
            attributeString(node, "text:change-id") === id,
        );
        const change = context.changes.get(id);
        if (endIndex < 0 || change?.kind !== "insertion") throw unknownOdtElement(child);
        const insertionBody: Element = {
          ...nodes[endIndex]!,
          elements: nodes.slice(index + 1, endIndex),
        };
        const parsedChildren = parseRuns(insertionBody, context);
        const first = parsedChildren[0];
        const children = (
          parsedChildren.length === 1 &&
          typeof first === "object" &&
          "text" in first &&
          first.text !== undefined
            ? [first.text]
            : parsedChildren
        ).filter(isTrackChangeChild);
        if (children.length !== parsedChildren.length) throw unknownOdtElement(child);
        resumeAfter = endIndex;
        return [
          {
            insertion: {
              id: change.id,
              author: change.author,
              date: change.date,
              children,
            },
          },
        ];
      }
      if (child.name === "text:change-end") return [];
      if (child.name === "text:change") {
        const id = attributeString(child, "text:change-id") ?? "";
        const change = context.changes.get(id);
        if (change?.kind !== "deletion") throw unknownOdtElement(child);
        const parsedChildren = change.children?.length
          ? parseRuns(change.children[0]!, context)
          : [];
        const first = parsedChildren[0];
        const children = (
          parsedChildren.length === 1 &&
          typeof first === "object" &&
          "text" in first &&
          first.text !== undefined
            ? [first.text]
            : parsedChildren
        ).filter(isTrackChangeChild);
        if (children.length !== parsedChildren.length) throw unknownOdtElement(child);
        return [
          {
            deletion: {
              id: change.id,
              author: change.author,
              date: change.date,
              children,
            },
          },
        ];
      }
      if (child.name === "text:format-change") throw unknownOdtElement(child);
      if (child.name === "text:bookmark-ref" || child.name === "text:reference-ref") {
        const format = fieldReferenceSwitch(attributeString(child, "text:reference-format"));
        return [
          { simpleField: { instruction: `REF ${encodeFieldArgument(textOf(child))}${format}` } },
        ];
      }
      if (child.name === "text:chapter") return [chapterField(parseChapter(child))];
      if (child.name === "text:expression") return [expressionField(parseExpression(child))];
      if (child.name === "text:variable-set") {
        const value = parseVariableSet(child);
        const expected = context.pendingVariables.get(value.name);
        if (expected !== undefined && expected !== value.valueType) throw unknownOdtElement(child);
        context.pendingVariables.delete(value.name);
        return [variableField(value)];
      }
      if (child.name === "text:sequence") {
        const value = parseSequence(child);
        if (!context.pendingSequences.delete(value.name)) throw unknownOdtElement(child);
        return [sequenceField(value)];
      }
      if (child.name === "text:hidden-text") return [hiddenTextField(parseHiddenText(child))];
      if (child.name === "text:bibliography-mark") throw unknownOdtElement(child);
      if (child.name === "office:annotation") return [parseAnnotation(child)];
      if (child.name === "draw:a") return [parseShapeHyperlink(child)];
      if (child.name === "draw:ellipse" || child.name === "draw:rect")
        return [{ wpsShape: parseShape(child) }];
      if (child.name === "text:span") {
        const properties = context.styles.get(
          attributeString(child, "text:style-name") ?? "",
        )?.character;
        return [
          {
            text: textOf(child),
            ...properties,
            underline: properties?.underline ? { type: "single" } : undefined,
          },
        ];
      }
      if (child.name === "text:tab") return [{ text: "", children: [{ tab: true }] }];
      if (child.name === "draw:frame") {
        return childNamed(child, "draw:object")
          ? parseChartFrame(child, context)
          : parsePictureFrame(child, context);
      }
      if (child.name === "draw:object") return parseChartFrame(child, context);
      if (child.name === "draw:custom-shape") {
        const shape = parseCustomShape(child, context);
        return shape ? [shape] : [];
      }
      if (child.name) throw unknownOdtElement(child);
      return [];
    },
  );
  return runs.reduce<(string | RunOptions | ParagraphChild)[]>((merged, run) => {
    const last = merged.at(-1);
    if (typeof last === "string" && typeof run === "string") merged[merged.length - 1] = last + run;
    else merged.push(run);
    return merged;
  }, []);
}

function bookmarkRunId(context: ParseContext, name: string): number {
  const existing = context.bookmarkIds.get(name);
  if (existing !== undefined) return existing;
  const id = nextBookmarkId(context);
  context.bookmarkIds.set(name, id);
  return id;
}

function assertElementNames(
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

function parseForms(element: Element, _context: ParseContext): SectionChild[] {
  const forms = childrenNamed(element, "form:form");
  if (forms.length !== 1) {
    throw new OdtParseError(
      "content.xml: ODF form grouping has no single canonical form-field equivalent",
      "content.xml",
      "/office:document-content/office:body/office:text/office:forms",
      "office:forms",
      "unsupported form grouping",
    );
  }
  const form = forms[0]!;
  assertElementNames(form, ["form:name"], ["form:text", "form:checkbox", "form:listbox"]);
  const controls = form.elements?.filter((child) => child.type === "element") ?? [];
  if (controls.length !== 1) {
    throw new OdtParseError(
      "content.xml: ODF form grouping has no single canonical form-field equivalent",
      "content.xml",
      "/office:document-content/office:body/office:text/office:forms/form:form",
      "form:form",
      "unsupported form grouping",
    );
  }
  const control = controls[0]!;
  const id = attributeString(control, "form:id");
  const name = attributeString(control, "form:name");
  if (id && !/^field\d+$/.test(id)) {
    if (name) throw unknownOdtElement(control);
    if (!/^[A-Za-z_:][\w.:-]*$/.test(id)) throw unknownOdtElement(control);
  }
  const common = {
    ...((name ?? id) ? { name: name ?? id } : {}),
    ...(attributeString(control, "form:disabled") === "true" ? { enabled: false } : {}),
    ...(attributeNumber(control, "form:tab-index") !== undefined
      ? { tabIndex: attributeNumber(control, "form:tab-index") }
      : {}),
  };
  const field: ParagraphChild = (() => {
    if (control.name === "form:text") {
      assertElementNames(
        control,
        [
          "form:id",
          "form:name",
          "form:disabled",
          "form:tab-index",
          "form:current-value",
          "form:max-length",
        ],
        [],
      );
      const value = attributeString(control, "form:current-value");
      return {
        formField: {
          ...common,
          textInput: {
            ...(value ? { value } : {}),
            ...(attributeNumber(control, "form:max-length") !== undefined
              ? { maxLength: attributeNumber(control, "form:max-length") }
              : {}),
          },
        },
      };
    }
    if (control.name === "form:checkbox") {
      assertElementNames(
        control,
        ["form:id", "form:name", "form:disabled", "form:tab-index", "form:current-state"],
        [],
      );
      const state = attributeString(control, "form:current-state");
      if (state !== "checked" && state !== "unchecked") throw unknownOdtElement(control);
      return { formField: { ...common, checkBox: { checked: state === "checked" } } };
    }
    assertElementNames(
      control,
      ["form:id", "form:name", "form:disabled", "form:tab-index"],
      ["form:option", "form:item"],
    );
    const entries = (control.elements ?? [])
      .filter((child) => ["form:option", "form:item"].includes(child.name ?? ""))
      .map((child) => {
        assertElementNames(child, ["form:selected"]);
        return textOf(child);
      });
    const selectedIndex = (control.elements ?? []).findIndex(
      (child) => attributeString(child, "form:selected") === "true",
    );
    return {
      formField: {
        ...common,
        dropDownList: {
          entries,
          ...(selectedIndex >= 0 ? { result: selectedIndex } : {}),
        },
      },
    };
  })();
  return [{ paragraph: { children: [field] } }];
}

function parseSequenceDeclarations(element: Element, context: ParseContext): void {
  assertElementNames(element, [], ["text:sequence-decl"]);
  for (const declaration of childrenNamed(element, "text:sequence-decl")) {
    assertElementNames(declaration, ["text:name", "text:display-outline-level"]);
    const name = attributeString(declaration, "text:name");
    if (!name) throw unknownOdtElement(declaration);
    const level = attributeNumber(declaration, "text:display-outline-level") ?? 0;
    if (level !== 0) throw unknownOdtElement(declaration);
    context.pendingSequences.add(name);
  }
}

function parseVariableDeclarations(element: Element, context: ParseContext): void {
  assertElementNames(element, [], ["text:variable-decl"]);
  for (const declaration of childrenNamed(element, "text:variable-decl")) {
    assertElementNames(declaration, ["text:name", "office:value-type"]);
    const name = attributeString(declaration, "text:name");
    const valueType = attributeString(declaration, "office:value-type") ?? "string";
    if (!name || (valueType !== "float" && valueType !== "string")) {
      throw unknownOdtElement(declaration);
    }
    context.pendingVariables.set(name, valueType);
  }
}

function parseTrackedChanges(element: Element, context: ParseContext): void {
  assertElementNames(element, ["text:track-changes"], ["text:changed-region"]);
  context.trackRevisions = attributeString(element, "text:track-changes") !== "false";
  for (const region of childrenNamed(element, "text:changed-region")) {
    assertElementNames(region, ["xml:id", "text:id"], ["text:insertion", "text:deletion"]);
    const id = attributeString(region, "text:id") ?? attributeString(region, "xml:id") ?? "";
    const children =
      region.elements?.filter((child) =>
        ["text:insertion", "text:deletion"].includes(child.name ?? ""),
      ) ?? [];
    if (children.length !== 1) throw unknownOdtElement(region);
    const change = children[0]!;
    const info = childNamed(change, "office:change-info");
    const author = info ? textOf(childNamed(info, "dc:creator")) : "";
    const date = info ? textOf(childNamed(info, "dc:date")) : "";
    if (!author || !date) throw unknownOdtElement(change);
    const deletedContent =
      change.name === "text:deletion"
        ? (change.elements ?? []).filter((child) => child.name === "text:p")
        : [];
    const metadata = (change.elements ?? []).filter((child) => child.name === "office:change-info");
    if (
      metadata.length !== 1 ||
      metadata.length + deletedContent.length !== (change.elements ?? []).length ||
      deletedContent.length > 1
    ) {
      throw unknownOdtElement(change);
    }
    context.changes.set(id, {
      id: context.changes.size + 1,
      author,
      date,
      kind: change.name === "text:insertion" ? "insertion" : "deletion",
      children: deletedContent,
    });
  }
}
function encodeFieldArgument(value: string): string {
  return value.replace(/([\\"]|\s)/g, "\\$1");
}

function variableField(value: ReturnType<typeof parseVariableSet>): {
  simpleField: CanonicalSimpleField;
} {
  const formula = value.formula ? ` = ${value.formula}` : "";
  return {
    simpleField: {
      instruction: `VARIABLE ${encodeFieldArgument(value.name)}${formula}`,
      cachedValue: value.display ?? value.value,
    },
  };
}

function sequenceField(value: ReturnType<typeof parseSequence>): {
  simpleField: CanonicalSimpleField;
} {
  if (value.referenceName || value.formula) {
    throw malformedOdtElement("text:sequence", "sequence metadata has no canonical mapping");
  }
  return {
    simpleField: {
      instruction: `SEQ ${encodeFieldArgument(value.name)}`,
      cachedValue: value.display,
    },
  };
}

function chapterField(value: ReturnType<typeof parseChapter>): {
  simpleField: CanonicalSimpleField;
} {
  return {
    simpleField: {
      instruction: `STYLEREF ${value.outlineLevel ?? 1} \\n`,
      cachedValue: value.display,
    },
  };
}

function expressionField(value: ReturnType<typeof parseExpression>): {
  simpleField: CanonicalSimpleField;
} {
  return {
    simpleField: {
      instruction: `= ${value.formula ?? ""}`.trim(),
      cachedValue: value.display ?? value.value,
    },
  };
}

function hiddenTextField(value: ReturnType<typeof parseHiddenText>): {
  simpleField: CanonicalSimpleField;
} {
  if (value.fixed) {
    throw malformedOdtElement("text:hidden-text", "fixed hidden text has no canonical mapping");
  }
  return {
    simpleField: {
      instruction: `IF ${encodeFieldArgument(value.condition)} "${value.content}" ""`,
      cachedValue: value.hidden ? undefined : value.content,
    },
  };
}

function parseVariableSet(element: Element): {
  name: string;
  valueType: "float" | "string";
  value?: string;
  display?: string;
  formula?: string;
} {
  const rawType = attributeString(element, "office:value-type") ?? "string";
  if (rawType !== "float" && rawType !== "string") throw unknownOdtElement(element);
  return {
    name: attributeString(element, "text:name") ?? "",
    valueType: rawType,
    value: attributeString(element, "office:value"),
    display: textOf(element) || attributeString(element, "text:display"),
    formula: attributeString(element, "text:formula"),
  };
}

function parseSequence(element: Element): {
  name: string;
  referenceName?: string;
  display?: string;
  formula?: string;
} {
  return {
    name: attributeString(element, "text:name") ?? "",
    referenceName: attributeString(element, "text:ref-name"),
    display: textOf(element) || attributeString(element, "text:display"),
    formula: attributeString(element, "text:formula"),
  };
}

function parseChapter(element: Element): { display?: string; outlineLevel?: number } {
  return {
    display: textOf(element) || attributeString(element, "text:display"),
    outlineLevel: attributeNumber(element, "text:outline-level"),
  };
}

function parseExpression(element: Element): {
  formula?: string;
  valueType: "float" | "string";
  value?: string;
  display?: string;
} {
  const rawType = attributeString(element, "office:value-type") ?? "string";
  if (rawType !== "float" && rawType !== "string") throw unknownOdtElement(element);
  return {
    formula: attributeString(element, "text:formula"),
    valueType: rawType,
    value: attributeString(element, "office:value"),
    display: textOf(element) || attributeString(element, "text:display"),
  };
}

function parseHiddenText(element: Element): {
  condition: string;
  content: string;
  hidden?: boolean;
  fixed?: boolean;
} {
  return {
    condition: attributeString(element, "text:condition") ?? "",
    content: textOf(element),
    hidden: attributeString(element, "text:is-hidden") === "true" ? true : undefined,
    fixed: attributeString(element, "text:is-fixed") === "true" ? true : undefined,
  };
}

function parseAnnotation(element: Element): ParagraphChild {
  if (attributeString(element, "office:display") === "true") throw unknownOdtElement(element);
  return {
    comment: {
      author: attributeString(element, "dc:creator"),
      date: attributeString(element, "dc:date"),
      children: childrenNamed(element, "text:p").map((paragraph) => ({ text: textOf(paragraph) })),
    },
  };
}

function parseIndex(element: Element): Extract<SectionChild, { toc: unknown }>["toc"] {
  const source = childNamed(element, "text:illustration-index-source");
  const title = childNamed(source, "text:index-title-template");
  const scope = attributeString(source, "text:index-scope");
  if (scope !== undefined && scope !== "document") throw unknownOdtElement(element);
  if (attributeString(source, "text:relative-tab-stop-position") === "true")
    throw unknownOdtElement(element);
  if (attributeString(element, "text:style-name") || attributeString(title, "text:style-name")) {
    throw unknownOdtElement(element);
  }
  const entries = childrenNamed(childNamed(element, "text:index-body"), "text:p").map(
    (paragraph): SectionChild => ({ paragraph: { text: textOf(paragraph) } }),
  );
  return {
    alias: attributeString(element, "text:name"),
    captionLabelIncludingNumbers:
      attributeString(source, "text:use-caption") === "true" ? "Figure" : undefined,
    entries: title ? [{ paragraph: { text: textOf(title) } }, ...entries] : entries,
  };
}

function parseBibliography(element: Element): Extract<SectionChild, { sdt: unknown }>["sdt"] {
  if (attributeString(element, "text:style-name")) throw unknownOdtElement(element);
  return {
    properties: {
      alias: attributeString(element, "text:name"),
      bibliography: true,
      richText: true,
    },
    children: childrenNamed(element, "text:index-body").flatMap((body) =>
      childrenNamed(body, "text:p").map((paragraph) => ({
        paragraph: { text: textOf(paragraph) },
      })),
    ),
  };
}

function parseShape(element: Element): ShapeOptions {
  const name = attributeString(element, "draw:name");
  return {
    children: [],
    geometry: element.name === "draw:rect" ? "rect" : "ellipse",
    ...(name ? { altText: { name } } : {}),
    transformation: {
      ...(lengthToEmu(attributeString(element, "svg:x"))
        ? { offset: { left: lengthToEmu(attributeString(element, "svg:x")) } }
        : {}),
      ...(lengthToEmu(attributeString(element, "svg:y"))
        ? { offset: { top: lengthToEmu(attributeString(element, "svg:y")) } }
        : {}),
      width: lengthToEmu(attributeString(element, "svg:width")) ?? 0,
      height: lengthToEmu(attributeString(element, "svg:height")) ?? 0,
    },
  };
}

function parseShapeHyperlink(element: Element): ParagraphChild {
  return {
    hyperlink: {
      url: attributeString(element, "xlink:href") ?? "",
      children: childrenNamed(element, "draw:ellipse")
        .concat(childrenNamed(element, "draw:rect"))
        .map((shape) => ({ wpsShape: parseShape(shape) })),
    },
  };
}
function parseNote(element: Element, context: ParseContext): ParagraphChild[] {
  const body = childNamed(element, "text:note-body");
  const children = body ? parseBlocks(body.elements ?? [], context) : [];
  const isEndnote = attributeString(element, "text:note-class") === "endnote";
  const notes = isEndnote ? context.notes.endnotes : context.notes.footnotes;
  notes.push({ id: notes.length + 1, children });
  const reference = isEndnote
    ? { endnoteReference: notes.length }
    : { footnoteReference: notes.length };
  return [reference];
}

/** draw:frame + draw:image maps back to an inline picture run. */
function parsePictureFrame(frame: Element, context: ParseContext): ParagraphChild[] {
  const href = attributeString(childNamed(frame, "draw:image"), "xlink:href");
  const path = href?.replace(/^\//, "");
  const data = path ? context.binaries[path] : undefined;
  if (!data || !path) return [];
  const extension = path.split(".").pop() ?? "png";
  const imageType = extension === "jpeg" ? "jpg" : extension;
  if (!isRasterPictureType(imageType)) {
    throw unknownOdtElement(frame);
  }
  return [
    {
      picture: {
        type: imageType,
        data,
        transformation: {
          width: lengthToEmu(attributeString(frame, "svg:width")) ?? 0,
          height: lengthToEmu(attributeString(frame, "svg:height")) ?? 0,
        },
      },
    },
  ];
}

/** draw:frame + draw:object resolves an embedded chart subdocument. */
function parseChartFrame(frame: Element, context: ParseContext): RunOptions[] {
  const href = attributeString(childNamed(frame, "draw:object"), "xlink:href")
    ?.replace(/^\.\//, "")
    .replace(/^\//, "");
  const chart = href ? context.chartBodies.get(href) : undefined;
  if (!chart) return [];
  return [
    {
      chart: {
        ...chart,
        transformation: {
          width: lengthToEmu(attributeString(frame, "svg:width")) ?? 0,
          height: lengthToEmu(attributeString(frame, "svg:height")) ?? 0,
        },
      },
    } as RunOptions,
  ];
}

/** draw:custom-shape maps back to the shared docx shape model. */
function parseCustomShape(
  element: Element,
  context: ParseContext,
): { wpsShape: ShapeOptions } | undefined {
  const styleName = attributeString(element, "draw:style-name");
  const graphic = styleName ? context.graphicStyles.get(styleName) : undefined;
  const enhanced = childNamed(element, "draw:enhanced-geometry");
  const presetType = attributeString(enhanced, "draw:type");
  const preset = presetType ? PRESET_GEOMETRY_DOCX[presetType] : undefined;
  const x = lengthToEmu(attributeString(element, "svg:x")) ?? 0;
  const y = lengthToEmu(attributeString(element, "svg:y")) ?? 0;
  const width = lengthToEmu(attributeString(element, "svg:width")) ?? 0;
  const height = lengthToEmu(attributeString(element, "svg:height")) ?? 0;
  const fill: FillOptions | undefined = graphicFill(graphic);
  const outline: OutlineOptions | undefined = graphicOutline(graphic);
  const name = attributeString(element, "draw:name");
  return {
    wpsShape: {
      children: [],
      transformation: {
        ...(x || y ? { offset: { ...(x ? { left: x } : {}), ...(y ? { top: y } : {}) } } : {}),
        width,
        height,
      },
      ...(name ? { altText: { name } } : {}),
      ...(preset ? { geometry: preset } : {}),
      ...(fill ? { fill } : {}),
      ...(outline ? { outline } : {}),
    },
  };
}
