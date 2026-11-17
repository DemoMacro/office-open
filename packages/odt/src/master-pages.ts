import type { SectionChild, SectionOptions } from "@office-open/docx";
import {
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  lengthToEmu,
  readXml,
  xmlElement,
  type OdfFiles,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { OdtGenerateError, OdtParseError } from "./error";
import type { OdtMasterPageOptions, OdtPageLayoutOptions } from "./options";

export interface MasterHeaderFooterContext {
  renderBlocks: (children: SectionChild[]) => string;
}

/** Section headers/footers render as master-page style:header/style:footer. */
export function masterHeaderFooter(
  section: Pick<SectionOptions, "headers" | "footers"> | undefined,
  context: MasterHeaderFooterContext,
): string {
  const header = section?.headers;
  const footer = section?.footers;
  if (
    (header?.default !== undefined && header.odd !== undefined) ||
    (footer?.default !== undefined && footer.odd !== undefined)
  ) {
    throw new OdtGenerateError(
      "ODT has no separate odd and default master-page header/footer regions",
      "styles.xml",
      "/office:document-styles/office:master-styles/style:master-page",
      "style:master-page",
      "ambiguous-master-header-footer-slots",
    );
  }
  return (
    headerFooterXml(header?.default ?? header?.odd, "style:header", context) +
    headerFooterXml(header?.even, "style:header-left", context) +
    headerFooterXml(header?.first, "style:header-first", context) +
    headerFooterXml(footer?.default ?? footer?.odd, "style:footer", context) +
    headerFooterXml(footer?.even, "style:footer-left", context) +
    headerFooterXml(footer?.first, "style:footer-first", context)
  );
}

export function headerFooterXml(
  children: SectionChild[] | undefined,
  element: string,
  context: MasterHeaderFooterContext,
): string {
  if (children === undefined) return "";
  return xmlElement(element, undefined, [context.renderBlocks(children)]);
}

/** Master-page header/footer regions → canonical section slots. */
export function parseMasterHeaderFooter(
  files: OdfFiles,
  parseBlocks: (elements: Element[], path: string) => SectionChild[],
): {
  headers?: SectionOptions["headers"];
  footers?: SectionOptions["footers"];
} {
  const master = childNamed(
    childNamed(readXml(files, "styles.xml"), "office:master-styles"),
    "style:master-page",
  );
  return master ? parseMasterRegions(master, parseBlocks) : {};
}

export function parseMasterPages(
  files: OdfFiles,
  parseBlocks: (elements: Element[], path: string) => SectionChild[],
): OdtMasterPageOptions[] {
  const styles = readXml(files, "styles.xml");
  return childrenNamed(childNamed(styles, "office:master-styles"), "style:master-page").flatMap(
    (master) => {
      const name = attributeString(master, "style:name");
      if (name === undefined) return [];
      const regions = parseMasterRegions(master, parseBlocks);
      return [
        {
          name,
          ...(attributeString(master, "style:page-layout-name")
            ? { pageLayoutName: attributeString(master, "style:page-layout-name") }
            : {}),
          ...regions,
        },
      ];
    },
  );
}

function parseMasterRegions(
  master: Element,
  parseBlocks: (elements: Element[], path: string) => SectionChild[],
): {
  headers?: SectionOptions["headers"];
  footers?: SectionOptions["footers"];
} {
  const regions = [
    ["style:header", "headers", "default"],
    ["style:header-left", "headers", "even"],
    ["style:header-first", "headers", "first"],
    ["style:footer", "footers", "default"],
    ["style:footer-left", "footers", "even"],
    ["style:footer-first", "footers", "first"],
  ] as const;
  const result: {
    headers?: SectionOptions["headers"];
    footers?: SectionOptions["footers"];
  } = {};
  for (const [elementName, family, slot] of regions) {
    const region = childNamed(master, elementName);
    if (!region) continue;
    const display = attributeString(region, "style:display");
    if (display !== undefined && display !== "true" && display !== "1") {
      throw new OdtParseError(
        `styles.xml: ${elementName} display state has no canonical section mapping`,
        "styles.xml",
        `/office:document-styles/office:master-styles/style:master-page/${elementName}/@style:display`,
        elementName,
        "unsupported-master-header-footer-display",
      );
    }
    const blocks = (region.elements ?? []).filter((child) => child.type === "element");
    const parsed = parseBlocks(
      blocks,
      `/office:document-styles/office:master-styles/style:master-page/${elementName}`,
    );
    (result[family] ??= {})[slot] = parsed;
  }
  return result;
}

export function parsePageLayouts(files: OdfFiles): OdtPageLayoutOptions[] {
  const styles = readXml(files, "styles.xml");
  return childrenNamed(childNamed(styles, "office:automatic-styles"), "style:page-layout").flatMap(
    (layout) => {
      const name = attributeString(layout, "style:name");
      if (name === undefined) return [];
      return [
        {
          name,
          properties: parsePageLayoutProperties(childNamed(layout, "style:page-layout-properties")),
        },
      ];
    },
  );
}

export function parsePageLayout(files: OdfFiles): SectionOptions["properties"] {
  const styles = readXml(files, "styles.xml");
  const master = childNamed(childNamed(styles, "office:master-styles"), "style:master-page");
  const pageLayoutName = attributeString(master, "style:page-layout-name");
  const pageLayouts = childrenNamed(
    childNamed(styles, "office:automatic-styles"),
    "style:page-layout",
  );
  const layout = childNamed(
    pageLayouts.find((item) => attributeString(item, "style:name") === pageLayoutName) ??
      pageLayouts[0],
    "style:page-layout-properties",
  );
  return parsePageLayoutProperties(layout);
}

function parsePageLayoutProperties(layout: Element | undefined): SectionOptions["properties"] {
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
