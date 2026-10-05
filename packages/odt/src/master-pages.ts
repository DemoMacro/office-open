import type { ParagraphOptions, SectionChild, SectionOptions } from "@office-open/docx";
import {
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  lengthToEmu,
  readXml,
  textOf,
  xmlElement,
  type OdfFiles,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { normalizeParagraph, spacesXml } from "./paragraph";

/** Section headers/footers render as master-page style:header/style:footer. */
export function masterHeaderFooter(
  section: Pick<SectionOptions, "headers" | "footers"> | undefined,
): string {
  return (
    headerFooterXml(section?.headers?.default, "style:header") +
    headerFooterXml(section?.footers?.default, "style:footer")
  );
}

export function headerFooterXml(children: SectionChild[] | undefined, element: string): string {
  const blocks = (children ?? [])
    .filter((child) => "paragraph" in child)
    .map((child) => headerParagraphXml(normalizeParagraph(child.paragraph)));
  if (blocks.length === 0) return "";
  return xmlElement(element, undefined, blocks);
}

export function headerParagraphXml(options: ParagraphOptions): string {
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

export function headerFieldXml(field: { instruction?: string; cachedValue?: string }): string {
  const value = field.cachedValue ?? "1";
  if (/NUMPAGES/i.test(field.instruction ?? ""))
    return xmlElement("text:page-count", undefined, [value]);
  return xmlElement("text:page-number", { "text:select-page": "current" }, [value]);
}

/** Master-page header/footer paragraphs → docx section header/footer children. */
export function parseMasterHeaderFooter(files: OdfFiles): {
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

export function parseHeaderParagraphs(container: Element): SectionChild[] {
  return childrenNamed(container, "text:p").map((paragraph) => {
    const runs = parseHeaderRuns(paragraph);
    const options: ParagraphOptions =
      runs.length === 1 && typeof runs[0] === "string" ? { text: runs[0] } : { children: runs };
    return { paragraph: options } as SectionChild;
  });
}

export function parseHeaderRuns(paragraph: Element): NonNullable<ParagraphOptions["children"]> {
  return (paragraph.elements ?? []).flatMap((child): NonNullable<ParagraphOptions["children"]> => {
    if (child.type === "text") return [String(child.text ?? "")];
    if (child.name === "text:page-number")
      return [{ simpleField: { instruction: " PAGE ", cachedValue: textOf(child) ?? "" } }];
    if (child.name === "text:page-count")
      return [{ simpleField: { instruction: " NUMPAGES ", cachedValue: textOf(child) ?? "" } }];
    return [];
  });
}

export function parsePageLayout(files: OdfFiles): SectionOptions["properties"] {
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

export function twipsToLength(value: number | string): string {
  return emuToLength(typeof value === "number" ? value * 635 : lengthToEmu(value));
}

export function lengthToTwips(value: string | undefined): number | undefined {
  const emu = lengthToEmu(value);
  return emu === undefined ? undefined : Math.round(emu / 635);
}
