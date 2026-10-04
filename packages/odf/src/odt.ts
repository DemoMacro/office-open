import { toUint8Array } from "@office-open/core";
import type {
  DocumentOptions,
  ParagraphOptions,
  PictureOptions,
  RunOptions,
  SectionChild,
  SectionOptions,
} from "@office-open/docx";
import type { Element } from "@office-open/xml";

import { ODF_NAMESPACES, escapeText, metaXml, parseMeta } from "./meta";
import { parseOdfNodes, serializeOdfNodes, type OdfXmlNode } from "./odf-node";
import { generateOcf, readOcf, readXml, type OdfFiles, type OdfPackageFiles } from "./package";
import { parseTable, tableXml } from "./table";
import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  lengthToEmu,
  textOf,
  xmlElement,
} from "./xml";

const MIME = "application/vnd.oasis.opendocument.text";
const NAMESPACES = ODF_NAMESPACES;

export type OdtOptions = DocumentOptions & {
  odfExtensions?: OdfXmlNode[];
  textSections?: OdtTextSectionOptions[];
};

/** A text:section wrapper; children remain body blocks in document order. */
export interface OdtTextSectionOptions {
  /** Section identifier emitted as text:name (for example, "Notes"). */
  name: string;
  styleName?: string;
  protected?: boolean;
  children?: SectionChild[];
}

/** Section headers/footers render as master-page style:header/style:footer. */
function masterHeaderFooter(section: SectionOptions | undefined): string {
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
  outline?: AbstractNumbering;
  binaries: Record<string, Uint8Array>;
  textSections: OdtTextSectionOptions[];
  notes: { footnotes: NoteEntry[]; endnotes: NoteEntry[] };
}

/** Footnote/endnote entry, indexed from the shared docx document model. */
type NoteEntry = NonNullable<DocumentOptions["footnotes"]>[number];

type NoteChildren = NoteEntry["children"];

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
}

export function generateOdt(options: OdtOptions): Uint8Array {
  const styles: string[] = [];
  const blocks = options.sections.flatMap((section) => section.children);
  const images: OdtImage[] = [];
  const notes = notesContext(options);
  const sections = (options.textSections ?? []).map((section) =>
    xmlElement(
      "text:section",
      {
        "text:name": section.name,
        "text:style-name": section.styleName,
        "text:protected": section.protected,
      },
      [blocksXml(section.children ?? [], styles, images, notes)],
    ),
  );
  const body = [
    blocksXml(blocks, styles, images, notes),
    ...sections,
    ...serializeOdfNodes(options.odfExtensions),
  ].join("");
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
  return generateOcf(MIME, files);
}

/** Note ids auto-assign 1, 2, … per class, matching the docx model. */
function notesContext(options: OdtOptions): NotesContext {
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
  return { footnotes, endnotes };
}

function blocksXml(
  blocks: SectionChild[],
  styles: string[],
  images: OdtImage[],
  notes: NotesContext,
): string {
  const parts: string[] = [];
  let index = 0;
  // Consecutive list paragraphs of the same level and kind group into one
  // text:list — the ODF shape for Word's bullet/numbering runs.
  while (index < blocks.length) {
    const listInfo = listParagraphLevel(blocks[index]!);
    if (listInfo === undefined) {
      parts.push(blockXml(blocks[index]!, styles, images, notes));
      index += 1;
      continue;
    }
    const group: SectionChild[] = [];
    while (index < blocks.length) {
      const info = listParagraphLevel(blocks[index]!);
      if (info?.level !== listInfo.level || info?.ordered !== listInfo.ordered) break;
      group.push(blocks[index]!);
      index += 1;
    }
    parts.push(listXml(group, listInfo.level, listInfo.ordered, styles, images, notes));
  }
  return parts.join("");
}

export function parseOdt(data: Uint8Array): OdtOptions {
  const { files, binaries } = readOcf(data, MIME);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:text");
  const styleContainer = childNamed(content, "office:automatic-styles");
  const styleMap = parseStyles(styleContainer);
  const rawNodes = parseOdfNodes(body);
  const context: ParseContext = {
    styles: styleMap,
    listStyles: parseListStyles(styleContainer),
    outline: parseOutlineStyle(files),
    binaries,
    textSections: [],
    notes: { footnotes: [], endnotes: [] },
  };
  const children = parseBlocks(body?.elements ?? [], context);
  const result: OdtOptions = {
    ...parseMeta(files),
    sections: [{ properties: parsePageLayout(files), children }],
    odfExtensions: rawNodes.filter(
      (node) =>
        !["text:p", "text:h", "table:table", "text:list", "text:section"].includes(node.name),
    ),
  };
  if (context.textSections.length > 0) result.textSections = context.textSections;
  if (context.notes.footnotes.length > 0) result.footnotes = context.notes.footnotes;
  if (context.notes.endnotes.length > 0) result.endnotes = context.notes.endnotes;
  const fonts = childrenNamed(childNamed(content, "office:font-face-decls"), "style:font-face").map(
    parseFontFace,
  );
  if (fonts.length > 0) result.fonts = fonts;
  if (context.outline) result.numbering = { abstractNumberings: [context.outline] };
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
  section: SectionOptions | undefined,
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
): string {
  if ("paragraph" in child)
    return paragraphXml(normalizeParagraph(child.paragraph), styles, images, notes);
  if ("table" in child)
    return tableXml(child.table, styles, (block) => blockXml(block, styles, images, notes));
  return "";
}

function normalizeParagraph(input: string | ParagraphOptions): ParagraphOptions {
  return typeof input === "string" ? { text: input } : input;
}

function listParagraphLevel(child: SectionChild): { level: number; ordered: boolean } | undefined {
  if (!("paragraph" in child)) return undefined;
  const options = normalizeParagraph(child.paragraph);
  if (options.heading) return undefined;
  if (options.bullet?.level !== undefined) return { level: options.bullet.level, ordered: false };
  const numbering = options.numbering;
  if (typeof numbering === "object" && "reference" in numbering)
    return { level: numbering.level ?? 0, ordered: true };
  return undefined;
}

function listXml(
  group: SectionChild[],
  level: number,
  ordered: boolean,
  styles: string[],
  images: OdtImage[],
  notes: NotesContext,
): string {
  const styleName = addListStyle(styles, ordered);
  // ODF nesting is 1-based: list level 0 renders as a single text:list,
  // level 1 nests one text:list inside the first list-item, and so on.
  const items = group
    .map((child) =>
      xmlElement("text:list-item", undefined, [blockXml(child, styles, images, notes)]),
    )
    .join("");
  let xml = items;
  for (let depth = 0; depth < level; depth += 1) {
    xml = xmlElement("text:list-item", undefined, [
      xmlElement("text:list", { "text:style-name": styleName }, [xml]),
    ]);
  }
  return xmlElement("text:list", { "text:style-name": styleName }, [xml]);
}

function addListStyle(styles: string[], ordered: boolean): string {
  const name = `L${styles.length + 1}`;
  styles.push(
    xmlElement("text:list-style", { "style:name": name }, [
      ordered
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

function paragraphXml(
  options: ParagraphOptions,
  styles: string[],
  images: OdtImage[],
  notes: NotesContext,
): string {
  const alignment = typeof options.alignment === "string" ? options.alignment : undefined;
  const styleName =
    alignment || options.pageBreakBefore || options.tabStops?.length
      ? addParagraphStyle(
          { alignment, pageBreakBefore: options.pageBreakBefore, tabStops: options.tabStops },
          styles,
        )
      : undefined;
  const children = runXml(options, styles, images, notes);
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
    if ("footnoteReference" in child) {
      return noteXml(
        (child as { footnoteReference: number | { id: number } }).footnoteReference,
        notes,
        "footnote",
        styles,
        images,
      );
    }
    if ("endnoteReference" in child) {
      return noteXml(
        (child as { endnoteReference: number | { id: number } }).endnoteReference,
        notes,
        "endnote",
        styles,
        images,
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
    return "";
  });
}

/** Note reference renders inline as text:note carrying its body paragraphs. */
function noteXml(
  reference: number | { id: number },
  notes: NotesContext,
  noteClass: "footnote" | "endnote",
  styles: string[],
  images: OdtImage[],
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
  for (const element of elements) {
    if (element.name === "text:section") {
      const textSection: OdtTextSectionOptions = {
        name: attributeString(element, "text:name") ?? "",
        styleName: attributeString(element, "text:style-name"),
        protected:
          attributeString(element, "text:protected") === undefined
            ? undefined
            : attributeString(element, "text:protected") === "true",
        children: [],
      };
      context.textSections.push(textSection);
      textSection.children = parseBlocks(element.elements ?? [], context, listDepth, listState);
      // Typed sections flatten into document children; the wrapper is retained
      // separately so name/style/protection metadata round-trip.
      result.push(...textSection.children);
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
      result.push({ pageBreak: true } as unknown as SectionChild);
      continue;
    }
    if (element.name === "text:p" || element.name === "text:h" || element.name === "table:table") {
      const child = parseBlock(element, context);
      if (listDepth > 0 && "paragraph" in child) {
        const paragraph = normalizeParagraph(child.paragraph);
        result.push({
          paragraph: listState?.ordered
            ? { ...paragraph, numbering: { reference: listState.name, level: listDepth - 1 } }
            : { ...paragraph, bullet: { level: listDepth - 1 } },
        } as SectionChild);
      } else {
        result.push(child);
      }
    }
  }
  return result;
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
): (string | RunOptions | { tab: true })[] {
  const runs = (element.elements ?? []).flatMap(
    (child): (string | RunOptions | { tab: true })[] => {
      if (child.type === "text") return [String(child.text ?? "")];
      if (child.name === "text:line-break") return [{ break: 1 } as unknown as RunOptions];
      if (child.name === "text:soft-page-break")
        return [{ pageBreak: true } as unknown as RunOptions];
      if (child.name === "text:s") {
        const count = attributeNumber(child, "text:c") ?? 1;
        return [" ".repeat(count)];
      }
      if (child.name === "text:a") {
        const href = attributeString(child, "xlink:href") ?? "";
        const link = {
          hyperlink: {
            ...(href.startsWith("#") ? { anchor: href.slice(1) } : { url: href }),
            tooltip: attributeString(child, "office:title"),
            targetFrame: attributeString(child, "office:target-frame-name"),
            children: parseRuns(child, context),
          },
        };
        return [link as unknown as RunOptions];
      }
      if (child.name === "text:note") return parseNote(child, context);
      if (child.name === "text:bookmark") {
        const bookmark = { bookmark: { name: attributeString(child, "text:name") ?? "" } };
        return [bookmark as unknown as RunOptions];
      }
      if (child.name === "text:span") {
        const properties = context.styles.get(
          attributeString(child, "text:style-name") ?? "",
        )?.character;
        const run: RunOptions = {
          text: textOf(child),
          ...properties,
          underline: properties?.underline ? { type: "single" } : undefined,
        };
        return [run];
      }
      if (child.name === "text:tab") return [{ text: "", children: [{ tab: true }] }];
      if (child.name === "draw:frame") return parsePictureFrame(child, context);
      return [];
    },
  );
  // Adjacent plain runs (for example text:s splits) merge back into one string
  // so text shorthand survives the round trip.
  return runs.reduce<(string | RunOptions | { tab: true })[]>((merged, run) => {
    const last = merged.at(-1);
    if (typeof last === "string" && typeof run === "string") merged[merged.length - 1] = last + run;
    else merged.push(run);
    return merged;
  }, []);
}

/** text:note stores its body in the document notes and leaves a reference run. */
function parseNote(element: Element, context: ParseContext): (string | RunOptions)[] {
  const body = childNamed(element, "text:note-body");
  const children = body ? parseBlocks(body.elements ?? [], context) : [];
  const isEndnote = attributeString(element, "text:note-class") === "endnote";
  const notes = isEndnote ? context.notes.endnotes : context.notes.footnotes;
  notes.push({ id: notes.length + 1, children });
  const reference = isEndnote
    ? { endnoteReference: notes.length }
    : { footnoteReference: notes.length };
  return [reference as unknown as RunOptions];
}

/** draw:frame + draw:image maps back to an inline picture run. */
function parsePictureFrame(frame: Element, context: ParseContext): RunOptions[] {
  const href = attributeString(childNamed(frame, "draw:image"), "xlink:href");
  const path = href?.replace(/^\//, "");
  const data = path ? context.binaries[path] : undefined;
  if (!data || !path) return [];
  const extension = path.split(".").pop() ?? "png";
  return [
    {
      picture: {
        type: extension,
        data,
        transformation: {
          width: lengthToEmu(attributeString(frame, "svg:width")) ?? 0,
          height: lengthToEmu(attributeString(frame, "svg:height")) ?? 0,
        },
      },
    },
  ] as unknown as RunOptions[];
}
