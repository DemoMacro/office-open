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
    character: CharacterProperties;
  }
>;

/** Shared lookup state threaded through the ODT parse pipeline. */
interface ParseContext {
  styles: StyleMap;
  listStyles: Map<string, boolean>;
  binaries: Record<string, Uint8Array>;
  textSections: OdtTextSectionOptions[];
}

export function generateOdt(options: OdtOptions): Uint8Array {
  const styles: string[] = [];
  const blocks = options.sections.flatMap((section) => section.children);
  const sectionProperties = options.sections[0]?.properties;
  const images: OdtImage[] = [];
  const sections = (options.textSections ?? []).map((section) =>
    xmlElement(
      "text:section",
      {
        "text:name": section.name,
        "text:style-name": section.styleName,
        "text:protected": section.protected,
      },
      [blocksXml(section.children ?? [], styles, images)],
    ),
  );
  const body = [
    blocksXml(blocks, styles, images),
    ...sections,
    ...serializeOdfNodes(options.odfExtensions),
  ].join("");
  const files: OdfPackageFiles = {
    "content.xml": contentXml(body, styles),
    "styles.xml": documentStylesXml(sectionProperties),
    "meta.xml": metaXml(options),
  };
  for (const image of images) files[image.path] = image.data;
  return generateOcf(MIME, files);
}

function blocksXml(blocks: SectionChild[], styles: string[], images: OdtImage[]): string {
  const parts: string[] = [];
  let index = 0;
  // Consecutive list paragraphs of the same level and kind group into one
  // text:list — the ODF shape for Word's bullet/numbering runs.
  while (index < blocks.length) {
    const listInfo = listParagraphLevel(blocks[index]!);
    if (listInfo === undefined) {
      parts.push(blockXml(blocks[index]!, styles, images));
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
    parts.push(listXml(group, listInfo.level, listInfo.ordered, styles, images));
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
    binaries,
    textSections: [],
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
  return result;
}

function contentXml(body: string, styles: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:automatic-styles>${styles.join(
    "",
  )}</office:automatic-styles><office:body><office:text>${body}</office:text></office:body></office:document-content>`;
}

function documentStylesXml(properties: SectionOptions["properties"]): string {
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
  const pageLayout = layoutAttributes.length
    ? `<style:page-layout style:name="pm1"><style:page-layout-properties ${layoutAttributes.join(
        " ",
      )}/></style:page-layout>`
    : "";
  const masterStyles = pageLayout
    ? `<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles>`
    : "";
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NAMESPACES} office:version="1.3"><office:styles/><office:automatic-styles>${pageLayout}</office:automatic-styles>${masterStyles}</office:document-styles>`;
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

function blockXml(child: SectionChild, styles: string[], images: OdtImage[]): string {
  if ("paragraph" in child)
    return paragraphXml(normalizeParagraph(child.paragraph), styles, images);
  if ("table" in child)
    return tableXml(child.table, styles, (block) => blockXml(block, styles, images));
  return "";
}

function normalizeParagraph(input: string | ParagraphOptions): ParagraphOptions {
  return typeof input === "string" ? { text: input } : input;
}

function listParagraphLevel(child: SectionChild): { level: number; ordered: boolean } | undefined {
  if (!("paragraph" in child)) return undefined;
  const options = normalizeParagraph(child.paragraph);
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
): string {
  const styleName = addListStyle(styles, ordered);
  // ODF nesting is 1-based: list level 0 renders as a single text:list,
  // level 1 nests one text:list inside the first list-item, and so on.
  const items = group
    .map((child) => xmlElement("text:list-item", undefined, [blockXml(child, styles, images)]))
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

function paragraphXml(options: ParagraphOptions, styles: string[], images: OdtImage[]): string {
  const alignment = typeof options.alignment === "string" ? options.alignment : undefined;
  const styleName =
    alignment || options.pageBreakBefore
      ? addParagraphStyle({ alignment, pageBreakBefore: options.pageBreakBefore }, styles)
      : undefined;
  const children = runXml(options, styles, images);
  const heading = /^Heading([1-9])$/.exec(options.heading ?? "");
  const attributes = {
    "text:style-name": styleName,
    "text:outline-level": heading ? Number(heading[1]) : undefined,
  };
  return xmlElement(heading ? "text:h" : "text:p", attributes, children);
}

function runXml(options: ParagraphOptions, styles: string[], images: OdtImage[]): string[] {
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
  properties: { alignment?: string; pageBreakBefore?: boolean },
  styles: string[],
): string {
  const name = `P${styles.length + 1}`;
  styles.push(
    xmlElement(
      "style:style",
      { "style:name": name, "style:family": "paragraph", "style:parent-style-name": "Standard" },
      [
        xmlElement("style:paragraph-properties", {
          "fo:text-align": properties.alignment,
          "fo:break-before": properties.pageBreakBefore ? "page" : undefined,
        }),
      ],
    ),
  );
  return name;
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
  if (headingLevel && headingLevel <= 6) {
    result.heading = `Heading${headingLevel}` as ParagraphOptions["heading"];
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
