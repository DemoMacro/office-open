import type {
  DocumentOptions,
  ParagraphOptions,
  RunOptions,
  SectionChild,
  TableOptions,
} from "@office-open/docx";
import type { Element } from "@office-open/xml";

import { ODF_NAMESPACES, escapeText, metaXml, parseMeta } from "./meta";
import { parseOdfNodes, serializeOdfNodes, type OdfXmlNode } from "./odf-node";
import { generateOcf, readOcf, readXml, type OdfFiles } from "./package";
import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  lengthToEmu,
  textOf,
  xmlElement,
} from "./xml";

const MIME = "application/vnd.oasis.opendocument.text";
const NAMESPACES = ODF_NAMESPACES;

export type OdtOptions = DocumentOptions & { odfExtensions?: OdfXmlNode[] };

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

export function generateOdt(options: OdtOptions): Uint8Array {
  const styles: string[] = [];
  const blocks = options.sections.flatMap((section) => section.children);
  const parts: string[] = [];
  let index = 0;
  // Consecutive bullet paragraphs of the same level group into one text:list —
  // the ODF shape for Word's bullet-list runs.
  while (index < blocks.length) {
    const bulletLevel = bulletParagraphLevel(blocks[index]!);
    if (bulletLevel === undefined) {
      parts.push(blockXml(blocks[index]!, styles));
      index += 1;
      continue;
    }
    const group: SectionChild[] = [];
    while (index < blocks.length && bulletParagraphLevel(blocks[index]!) === bulletLevel) {
      group.push(blocks[index]!);
      index += 1;
    }
    parts.push(listXml(group, bulletLevel, styles));
  }
  const body = [parts.join(""), ...serializeOdfNodes(options.odfExtensions)].join("");
  const files: OdfFiles = {
    "content.xml": contentXml(body, styles),
    "styles.xml": documentStylesXml(),
    "meta.xml": metaXml(options),
  };
  return generateOcf(MIME, files);
}

export function parseOdt(data: Uint8Array): OdtOptions {
  const { files } = readOcf(data, MIME);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:text");
  const styleMap = parseStyles(childNamed(content, "office:automatic-styles"));
  const rawNodes = parseOdfNodes(body);
  const children = parseBlocks(body?.elements ?? [], styleMap);
  return {
    ...parseMeta(files),
    sections: [{ children }],
    odfExtensions: rawNodes.filter(
      (node) => !["text:p", "text:h", "table:table", "text:list"].includes(node.name),
    ),
  };
}

function contentXml(body: string, styles: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:automatic-styles>${styles.join(
    "",
  )}</office:automatic-styles><office:body><office:text>${body}</office:text></office:body></office:document-content>`;
}

function documentStylesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NAMESPACES} office:version="1.3"><office:styles/></office:document-styles>`;
}

function blockXml(child: SectionChild, styles: string[]): string {
  if ("paragraph" in child) return paragraphXml(normalizeParagraph(child.paragraph), styles);
  if ("table" in child) return tableXml(child.table, styles);
  return "";
}

function normalizeParagraph(input: string | ParagraphOptions): ParagraphOptions {
  return typeof input === "string" ? { text: input } : input;
}

function bulletParagraphLevel(child: SectionChild): number | undefined {
  if (!("paragraph" in child)) return undefined;
  const options = normalizeParagraph(child.paragraph);
  return options.bullet?.level;
}

function listXml(group: SectionChild[], level: number, styles: string[]): string {
  const styleName = addListStyle(styles);
  // ODF nesting is 1-based: bullet level 0 renders as a single text:list,
  // level 1 nests one text:list inside the first list-item, and so on.
  const items = group
    .map((child) => xmlElement("text:list-item", undefined, [blockXml(child, styles)]))
    .join("");
  let xml = items;
  for (let depth = 0; depth < level; depth += 1) {
    xml = xmlElement("text:list-item", undefined, [
      xmlElement("text:list", { "text:style-name": styleName }, [xml]),
    ]);
  }
  return xmlElement("text:list", { "text:style-name": styleName }, [xml]);
}

function addListStyle(styles: string[]): string {
  const name = `L${styles.length + 1}`;
  styles.push(
    xmlElement("text:list-style", { "style:name": name }, [
      xmlElement("text:list-level-style-bullet", { "text:level": 1, "text:bullet-char": "•" }, [
        xmlElement("style:list-level-properties", {
          "style:list-level-position-and-space-mode": "label-alignment",
        }),
        xmlElement("style:text-properties", { "fo:font-family": "OpenSymbol" }),
      ]),
    ]),
  );
  return name;
}

function paragraphXml(options: ParagraphOptions, styles: string[]): string {
  const alignment = typeof options.alignment === "string" ? options.alignment : undefined;
  const styleName =
    alignment || options.pageBreakBefore
      ? addParagraphStyle({ alignment, pageBreakBefore: options.pageBreakBefore }, styles)
      : undefined;
  const children = runXml(options, styles);
  const heading = /^Heading([1-9])$/.exec(options.heading ?? "");
  const attributes = {
    "text:style-name": styleName,
    "text:outline-level": heading ? Number(heading[1]) : undefined,
  };
  return xmlElement(heading ? "text:h" : "text:p", attributes, children);
}

function runXml(options: ParagraphOptions, styles: string[]): string[] {
  if (options.text !== undefined && options.children === undefined) {
    return [escapeText(options.text)];
  }
  return (options.children ?? []).map((child) => {
    if (typeof child === "string") return `<text:span>${escapeText(child)}</text:span>`;
    if ("pageBreak" in child) return "<text:soft-page-break/>";
    if ("columnBreak" in child) return "<text:line-break/>";
    if ("break" in child) return lineBreakXml(child as RunOptions);
    if ("text" in child) {
      const run = child as RunOptions;
      const styleName = addCharacterStyle(characterProperties(run), styles);
      return (
        lineBreakXml(run) +
        xmlElement("text:span", { "text:style-name": styleName }, [escapeText(run.text ?? "")])
      );
    }
    if ("tab" in child) return "<text:tab/>";
    return "";
  });
}

/** Run breaks render as text:line-break (ODF has no w:br/@clear equivalent). */
function lineBreakXml(run: RunOptions): string {
  if (!run.break) return "";
  const count = typeof run.break === "number" ? run.break : (run.break.count ?? 1);
  return "<text:line-break/>".repeat(Math.max(0, count));
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

function tableXml(table: TableOptions, styles: string[]): string {
  const explicitColumns = (table.columnWidths ?? []).map((width) => {
    const twips = typeof width === "number" ? width : Math.round(lengthToEmu(width)! / 635);
    const name = `T${styles.length + 1}`;
    styles.push(
      xmlElement("style:style", { "style:name": name, "style:family": "table-column" }, [
        xmlElement("style:table-column-properties", {
          "style:column-width": `${Number((twips / 567).toFixed(4))}cm`,
        }),
      ]),
    );
    return xmlElement("table:table-column", { "table:style-name": name });
  });
  const rows = table.rows.map((row) => {
    const cells = ("cells" in row ? row.cells : []).map((cell) => {
      const span = "columnSpan" in cell ? cell.columnSpan : undefined;
      const children =
        "children" in cell ? cell.children.map((child) => blockXml(child, styles)) : [];
      const xml = xmlElement(
        "table:table-cell",
        {
          "office:value-type": "string",
          "table:number-columns-spanned": span,
        },
        children,
      );
      const covered = Array.from({ length: Math.max(0, (span ?? 1) - 1) }, () =>
        xmlElement("table:covered-table-cell"),
      );
      return xml + covered.join("");
    });
    return xmlElement("table:table-row", undefined, cells);
  });
  const rowSpans = table.rows.map((row) =>
    ("cells" in row ? row.cells : []).reduce(
      (total, cell) => total + ("columnSpan" in cell ? (cell.columnSpan ?? 1) : 1),
      0,
    ),
  );
  const columnCount = Math.max(explicitColumns.length, ...rowSpans, 1);
  const columns = Array.from(
    { length: columnCount },
    (_, column) => explicitColumns[column] ?? xmlElement("table:table-column"),
  );
  const bodyRows = rows.length
    ? rows
    : [xmlElement("table:table-row", undefined, [xmlElement("table:table-cell")])];
  return xmlElement("table:table", { "table:name": `Table${styles.length + 1}` }, [
    columns.join(""),
    bodyRows.join(""),
  ]);
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

function parseBlock(element: Element, styles: StyleMap): SectionChild {
  if (element.name === "table:table") return parseTable(element, styles);
  const paragraph = parseParagraph(element, styles);
  return { paragraph };
}

function parseBlocks(elements: Element[], styles: StyleMap, listDepth = 0): SectionChild[] {
  const result: SectionChild[] = [];
  for (const element of elements) {
    if (element.name === "text:list") {
      // A typed text:list unwraps to bullet paragraphs at the nesting depth;
      // list-header content (rare) keeps the generic fall-through below.
      for (const item of childrenNamed(element, "text:list-item")) {
        result.push(...parseBlocks(item.elements ?? [], styles, listDepth + 1));
      }
      continue;
    }
    if (element.name === "text:soft-page-break") {
      result.push({ pageBreak: true } as unknown as SectionChild);
      continue;
    }
    if (element.name === "text:p" || element.name === "text:h" || element.name === "table:table") {
      const child = parseBlock(element, styles);
      if (listDepth > 0 && "paragraph" in child) {
        const paragraph = normalizeParagraph(child.paragraph);
        result.push({
          paragraph: { ...paragraph, bullet: { level: listDepth - 1 } },
        } as SectionChild);
      } else {
        result.push(child);
      }
    }
  }
  return result;
}

function parseParagraph(element: Element, styles: StyleMap): ParagraphOptions {
  const style = styles.get(attributeString(element, "text:style-name") ?? "");
  const headingLevel = attributeNumber(element, "text:outline-level");
  const runs = parseRuns(element, styles);
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

function parseRuns(element: Element, styles: StyleMap): (string | RunOptions | { tab: true })[] {
  return (element.elements ?? []).flatMap((child): (string | RunOptions | { tab: true })[] => {
    if (child.type === "text") return [String(child.text ?? "")];
    if (child.name === "text:line-break") return [{ break: 1 } as unknown as RunOptions];
    if (child.name === "text:soft-page-break")
      return [{ pageBreak: true } as unknown as RunOptions];
    if (child.name === "text:s") {
      const count = attributeNumber(child, "text:c") ?? 1;
      return [" ".repeat(count)];
    }
    if (child.name === "text:span") {
      const properties = styles.get(attributeString(child, "text:style-name") ?? "")?.character;
      const run: RunOptions = {
        text: textOf(child),
        ...properties,
        underline: properties?.underline ? { type: "single" } : undefined,
      };
      return [run];
    }
    if (child.name === "text:tab") return [{ text: "", children: [{ tab: true }] }];
    return [];
  });
}

function parseTable(element: Element, styles: StyleMap): SectionChild {
  const columnWidths = childrenNamed(element, "table:table-column").map((column) => {
    const style = styles.get(attributeString(column, "table:style-name") ?? "");
    return style?.columnWidth ?? 5000;
  });
  const rows = childrenNamed(element, "table:table-row").map((row) => ({
    cells: childrenNamed(row, "table:table-cell").map((cell) => {
      const cellChildren =
        cell.elements
          ?.filter(
            (child) =>
              child.name === "text:p" || child.name === "text:h" || child.name === "table:table",
          )
          .map((child) => parseBlock(child, styles)) ?? [];
      return {
        children: cellChildren,
        columnSpan: attributeNumber(cell, "table:number-columns-spanned"),
      };
    }),
  }));
  return { table: { columnWidths, rows } };
}
