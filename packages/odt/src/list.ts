import type { DocumentOptions, SectionChild } from "@office-open/docx";
import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  readXml,
  xmlElement,
  type OdfFiles,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { blockXml } from "./body";
import type { OdtChart, OdtImage } from "./drawing";
import { lengthToTwips, twipsToLength } from "./master-pages";
import type { OdtEmbeddedObjectOptions } from "./package";
import { normalizeParagraph } from "./paragraph";
import { NUM_FORMAT_DOCX, NUM_FORMAT_ODF, type NotesContext } from "./section";

export type AbstractNumbering = NonNullable<
  NonNullable<DocumentOptions["numbering"]>["abstractNumberings"]
>[number];

export type NumberingLevel = AbstractNumbering["levels"][number];

/** Docx level suffix → ODF text:label-followed-by token. */
export const SUFFIX_ODF: Record<string, string> = {
  nothing: "nothing",
  space: "space",
  tab: "listtab",
};

/** ODF text:label-followed-by token → docx level suffix. */
export const SUFFIX_DOCX: Record<string, string> = {
  listtab: "tab",
  nothing: "nothing",
  space: "space",
};

/** Heading-linked abstract numbering renders as text:outline-style. */
export function outlineStyleXml(numbering: DocumentOptions["numbering"]): string {
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

export function outlineLevelXml(level: NumberingLevel): string {
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
export function parseOutlineStyle(files: OdfFiles): AbstractNumbering | undefined {
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

/** Non-outline abstract numberings render as reusable text:list-style definitions. */
export function listStyleXml(definition: AbstractNumbering): string {
  return xmlElement(
    "text:list-style",
    { "style:name": definition.reference },
    definition.levels.map(listLevelXml),
  );
}

export function listLevelXml(level: NumberingLevel): string {
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

export function listParagraphLevel(
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

export function listXml(
  group: SectionChild[],
  info: { level: number; ordered: boolean; reference?: string },
  styles: string[],
  images: OdtImage[],
  notes: NotesContext,
  numbering: DocumentOptions["numbering"],
  charts: OdtChart[],
  embeddedObjects: OdtEmbeddedObjectOptions[] = [],
): string {
  const styleName = addListStyle(styles, info, numbering);
  // ODF nesting is 1-based: list level 0 renders as a single text:list,
  // level 1 nests one text:list inside the first list-item, and so on.
  const items = group
    .map((child) =>
      xmlElement("text:list-item", undefined, [
        blockXml(child, styles, images, notes, numbering, charts, embeddedObjects),
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

export function addListStyle(
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
export function parseListStyles(container: Element | undefined): Map<string, boolean> {
  const result = new Map<string, boolean>();
  for (const style of childrenNamed(container, "text:list-style")) {
    const name = attributeString(style, "style:name") ?? "";
    result.set(name, childrenNamed(style, "text:list-level-style-number").length > 0);
  }
  return result;
}

/** text:list-style definitions → abstract numberings keyed by style name. */
export function parseListNumberings(container: Element | undefined): AbstractNumbering[] {
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

export function parseListLevel(element: Element): NumberingLevel {
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
