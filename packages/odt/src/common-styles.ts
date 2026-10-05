import type { DocumentOptions, SectionOptions } from "@office-open/docx";
import {
  attributeString,
  childNamed,
  childrenNamed,
  readXml,
  xmlElement,
  type OdfFiles,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { NAMESPACES } from "./body";
import { outlineStyleXml } from "./list";
import { lengthToTwips, masterHeaderFooter, twipsToLength } from "./master-pages";
import { notesConfigurationXml } from "./notes";

/** Docx font-table entry, indexed from the shared document model. */
export type FontEntry = NonNullable<DocumentOptions["fonts"]>[number];

export type DocumentDefaults = NonNullable<
  NonNullable<NonNullable<DocumentOptions["styles"]>["default"]>["document"]
>;

/** Docx font table → ODF font-face declarations (no embedded data in ODF). */
export function fontFaceDecls(fonts: DocumentOptions["fonts"]): string {
  if (!fonts?.length) return "";
  return xmlElement("office:font-face-decls", undefined, fonts.map(fontFaceXml));
}

export function fontFaceXml(font: FontEntry): string {
  return xmlElement("style:font-face", {
    "style:name": font.name,
    "style:font-family-generic": font.family,
    "style:font-pitch": font.pitch,
    "svg:panose-1": font.panose1,
  });
}

export function parseFontFace(element: Element): FontEntry {
  return {
    name: attributeString(element, "style:name") ?? "",
    family: attributeString(element, "style:font-family-generic"),
    pitch: attributeString(element, "style:font-pitch"),
    panose1: attributeString(element, "svg:panose-1"),
  };
}

export function documentStylesXml(
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

/** Docx document defaults → ODF style:default-style (paragraph family). */
export function defaultStyleXml(styles: DocumentOptions["styles"]): string {
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
export function parseDefaultStyle(files: OdfFiles): DocumentDefaults | undefined {
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
