import type {
  DocumentOptions,
  ParagraphStylePropertiesOptions,
  RunStylePropertiesOptions,
  SectionOptions,
} from "@office-open/docx";
import {
  attributeString,
  childNamed,
  childrenNamed,
  readXml,
  type GraphicStyleResources,
  graphicStyleResourcesXml,
  xmlElement,
  type OdfFiles,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { NAMESPACES } from "./body";
import { OdtParseError } from "./error";
import { outlineStyleXml } from "./list";
import { lengthToTwips, twipsToLength } from "./master-pages";
import { notesConfigurationXml } from "./notes";
import type { OdtMasterPageOptions, OdtPageLayoutOptions } from "./options";

/** Docx font-table entry, indexed from the shared document model. */
export type FontEntry = NonNullable<DocumentOptions["fonts"]>[number];

export type DocumentDefaults = NonNullable<
  NonNullable<NonNullable<DocumentOptions["styles"]>["default"]>["document"]
>;

type CommonStyle =
  | NonNullable<NonNullable<DocumentOptions["styles"]>["paragraphStyles"]>[number]
  | NonNullable<NonNullable<DocumentOptions["styles"]>["characterStyles"]>[number];

type ParsedCommonStyle = {
  id: string;
  name: string;
  family?: string;
  alignment?: ParagraphStylePropertiesOptions["alignment"];
  basedOn?: string;
  next?: string;
  indent?: ParagraphStylePropertiesOptions["indent"];
  spacing?: ParagraphStylePropertiesOptions["spacing"];
  run?: RunStylePropertiesOptions;
  relativeFontSize?: number;
};

function textPropertiesXml(style: Pick<CommonStyle, "run">): string {
  const run = style.run;
  return xmlElement("style:text-properties", {
    "fo:font-weight": run?.bold ? "bold" : undefined,
    "fo:font-style": run?.italic ? "italic" : undefined,
    "style:text-underline-style": run?.underline?.type ? "solid" : undefined,
    "style:text-line-through-style": run?.strike ? "solid" : undefined,
    "fo:font-size": typeof run?.size === "number" ? `${run.size}pt` : undefined,
    "fo:color": typeof run?.color === "string" ? `#${run.color.slice(-6)}` : undefined,
    "fo:font-family": typeof run?.font === "string" ? run.font : undefined,
  });
}

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
  headerFooter: string,
  pageLayouts?: readonly OdtPageLayoutOptions[],
  masterPages?: readonly OdtMasterPageOptions[],
  masterPageXml?: readonly string[],
  graphicResources?: GraphicStyleResources,
): string {
  const properties = section?.properties;
  const needsDefaultLayout =
    pageLayouts?.length || properties !== undefined || headerFooter !== "" || masterPages?.length;
  const layouts = pageLayouts?.length
    ? pageLayouts
    : needsDefaultLayout
      ? [{ name: "pm1", properties }]
      : [];
  const pageLayout = layouts
    .map(
      (layout) =>
        `<style:page-layout style:name="${layout.name}">${pageLayoutPropertiesXml(
          layout.properties,
        )}</style:page-layout>`,
    )
    .join("");
  const masters = masterPageXml?.length
    ? masterPageXml.join("")
    : masterPages?.length
      ? masterPages
          .map(
            (master) =>
              `<style:master-page style:name="${master.name}"${
                master.pageLayoutName ? ` style:page-layout-name="${master.pageLayoutName}"` : ""
              }/>`,
          )
          .join("")
      : headerFooter !== ""
        ? `<style:master-page style:name="Standard" style:page-layout-name="${layouts[0]?.name ?? "pm1"}">${headerFooter}</style:master-page>`
        : "";
  const masterStyles = masters ? `<office:master-styles>${masters}</office:master-styles>` : "";
  const notes =
    commonStylesXml(styles) +
    notesConfigurationXml(settings) +
    outlineStyleXml(numbering) +
    defaultStyleXml(styles);
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NAMESPACES} office:version="1.3"><office:styles>${notes}${graphicResources ? graphicStyleResourcesXml(graphicResources).join("") : ""}</office:styles><office:automatic-styles>${pageLayout}</office:automatic-styles>${masterStyles}</office:document-styles>`;
}

function pageLayoutPropertiesXml(properties: SectionOptions["properties"]): string {
  const pageSize = typeof properties?.pageSize === "object" ? properties.pageSize : undefined;
  const pageMargin = typeof properties?.pageMargin === "object" ? properties.pageMargin : undefined;
  const attributes = [
    pageSize?.width !== undefined && `fo:page-width="${twipsToLength(pageSize.width)}"`,
    pageSize?.height !== undefined && `fo:page-height="${twipsToLength(pageSize.height)}"`,
    pageSize?.orientation && `style:print-orientation="${pageSize.orientation}"`,
    pageMargin?.top !== undefined && `fo:margin-top="${twipsToLength(pageMargin.top)}"`,
    pageMargin?.right !== undefined && `fo:margin-right="${twipsToLength(pageMargin.right)}"`,
    pageMargin?.bottom !== undefined && `fo:margin-bottom="${twipsToLength(pageMargin.bottom)}"`,
    pageMargin?.left !== undefined && `fo:margin-left="${twipsToLength(pageMargin.left)}"`,
  ].filter(Boolean);
  return `<style:page-layout-properties${attributes.length ? ` ${attributes.join(" ")}` : ""}/>`;
}

export function commonStylesXml(styles: DocumentOptions["styles"]): string {
  return [
    ...(styles?.paragraphStyles ?? []).map((style) => {
      const paragraph = style.paragraph;
      return xmlElement(
        "style:style",
        {
          "style:name": style.id,
          "style:family": "paragraph",
          "style:display-name": style.name,
          "style:parent-style-name": style.basedOn,
          "style:next-style-name": style.next,
        },
        [
          xmlElement("style:paragraph-properties", {
            "fo:text-align": paragraph?.alignment,
            "fo:margin-left":
              paragraph?.indent?.left !== undefined
                ? twipsToLength(paragraph.indent.left)
                : undefined,
            "fo:margin-right":
              paragraph?.indent?.right !== undefined
                ? twipsToLength(paragraph.indent.right)
                : undefined,
            "fo:margin-top":
              paragraph?.spacing?.before !== undefined
                ? twipsToLength(paragraph.spacing.before)
                : undefined,
            "fo:margin-bottom":
              paragraph?.spacing?.after !== undefined
                ? twipsToLength(paragraph.spacing.after)
                : undefined,
          }),
          textPropertiesXml(style),
        ],
      );
    }),
    ...(styles?.characterStyles ?? []).map((style) =>
      xmlElement(
        "style:style",
        {
          "style:name": style.id,
          "style:family": "text",
          "style:display-name": style.name,
          "style:parent-style-name": style.basedOn,
          "style:next-style-name": style.next,
        },
        [textPropertiesXml(style)],
      ),
    ),
  ].join("");
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

export function parseCommonStyles(files: OdfFiles): DocumentOptions["styles"] {
  const container = childNamed(readXml(files, "styles.xml"), "office:styles");
  const parsed = new Map<string, ParsedCommonStyle>();
  for (const style of childrenNamed(container, "style:style")) {
    const id = attributeString(style, "style:name") ?? "";
    const character = childNamed(style, "style:text-properties");
    const paragraph = childNamed(style, "style:paragraph-properties");
    const indentLeft = lengthToTwips(attributeString(paragraph, "fo:margin-left"));
    const indentRight = lengthToTwips(attributeString(paragraph, "fo:margin-right"));
    const spacingBefore = lengthToTwips(attributeString(paragraph, "fo:margin-top"));
    const spacingAfter = lengthToTwips(attributeString(paragraph, "fo:margin-bottom"));
    const indent =
      indentLeft !== undefined || indentRight !== undefined
        ? {
            ...(indentLeft !== undefined ? { left: indentLeft } : {}),
            ...(indentRight !== undefined ? { right: indentRight } : {}),
          }
        : undefined;
    const spacing =
      spacingBefore !== undefined || spacingAfter !== undefined
        ? {
            ...(spacingBefore !== undefined ? { before: spacingBefore } : {}),
            ...(spacingAfter !== undefined ? { after: spacingAfter } : {}),
          }
        : undefined;
    const run = parseRunProperties(character, id);
    parsed.set(id, {
      id,
      name: attributeString(style, "style:display-name") ?? id,
      family: attributeString(style, "style:family"),
      alignment: parseAlignment(attributeString(paragraph, "fo:text-align"), id),
      basedOn: attributeString(style, "style:parent-style-name"),
      next: attributeString(style, "style:next-style-name"),
      indent,
      spacing,
      ...(run.run ? { run: run.run } : {}),
      ...(run.relativeFontSize === undefined ? {} : { relativeFontSize: run.relativeFontSize }),
    });
  }
  const defaultSize = parseDefaultStyle(files)?.run?.size ?? 12;
  const resolvedSizes = new Map<string, number>();
  function effectiveFontSize(name: string, path = new Set<string>()): number {
    const cached = resolvedSizes.get(name);
    if (cached !== undefined) return cached;
    const style = parsed.get(name);
    if (!style) return defaultSize;
    if (path.has(name)) throw invalidStyle(name, "style:parent-style-name", "cyclic style chain");
    const parentPath = new Set(path).add(name);
    const inherited = style.basedOn ? effectiveFontSize(style.basedOn, parentPath) : defaultSize;
    const size =
      style.run?.size ??
      (style.relativeFontSize === undefined
        ? inherited
        : (inherited * style.relativeFontSize) / 100);
    resolvedSizes.set(name, size);
    return size;
  }
  for (const [name, style] of parsed) {
    if (style.relativeFontSize === undefined) continue;
    style.run = { ...style.run, size: effectiveFontSize(name) };
  }
  const paragraphStyles: NonNullable<NonNullable<DocumentOptions["styles"]>["paragraphStyles"]> = [
    ...parsed,
  ]
    .filter(([, style]) => style.family === "paragraph")
    .map(([name, style]) => ({
      id: style.id,
      name: style.name ?? name,
      ...(style.basedOn ? { basedOn: style.basedOn } : {}),
      ...(style.next ? { next: style.next } : {}),
      ...(style.alignment !== undefined || style.indent !== undefined || style.spacing !== undefined
        ? {
            paragraph: {
              alignment: style.alignment,
              ...(style.indent ? { indent: style.indent } : {}),
              ...(style.spacing ? { spacing: style.spacing } : {}),
            },
          }
        : {}),
      ...(style.run ? { run: style.run } : {}),
    }));
  const characterStyles: NonNullable<NonNullable<DocumentOptions["styles"]>["characterStyles"]> = [
    ...parsed,
  ]
    .filter(([, style]) => style.family === "text")
    .map(([name, style]) => ({
      id: style.id,
      name: style.name ?? name,
      ...(style.basedOn ? { basedOn: style.basedOn } : {}),
      ...(style.next ? { next: style.next } : {}),
      ...(style.run ? { run: style.run } : {}),
    }));
  const result = {
    ...(paragraphStyles.length > 0 ? { paragraphStyles } : {}),
    ...(characterStyles.length > 0 ? { characterStyles } : {}),
  };
  return Object.keys(result).length > 0 ? result : undefined;
}

function parseRunProperties(
  element: Element | undefined,
  styleName: string,
): { run?: RunStylePropertiesOptions; relativeFontSize?: number } {
  const size = attributeString(element, "fo:font-size");
  const color = attributeString(element, "fo:color");
  const weight = attributeString(element, "fo:font-weight");
  const fontStyle = attributeString(element, "fo:font-style");
  const underline = attributeString(element, "style:text-underline-style");
  const strike = attributeString(element, "style:text-line-through-style");
  const relativeFontSize = size?.endsWith("%") ? Number(size.slice(0, -1)) : undefined;
  if (
    relativeFontSize !== undefined &&
    (!Number.isFinite(relativeFontSize) || relativeFontSize < 0)
  )
    throw invalidStyle(styleName, "fo:font-size", "invalid relative size");
  const run: RunStylePropertiesOptions = {
    ...(weight === undefined ? {} : { bold: weight === "bold" }),
    ...(fontStyle === undefined ? {} : { italic: fontStyle === "italic" }),
    ...(underline === undefined || underline === "none"
      ? {}
      : { underline: { type: "single" as const } }),
    ...(strike === undefined || strike === "none" ? {} : { strike: true }),
    ...(size === undefined || relativeFontSize !== undefined
      ? {}
      : { size: parsePointSize(size, styleName) }),
    ...(color?.startsWith("#") ? { color: color.slice(1) } : {}),
    ...(attributeString(element, "fo:font-family")
      ? { font: attributeString(element, "fo:font-family") }
      : {}),
  };
  return {
    ...(Object.keys(run).length > 0 ? { run } : {}),
    ...(relativeFontSize === undefined ? {} : { relativeFontSize }),
  };
}

function parsePointSize(value: string, styleName: string): number {
  const size = value.endsWith("pt")
    ? Number(value.slice(0, -2))
    : (lengthToTwips(value) ?? Number.NaN) / 20;
  if (!Number.isFinite(size) || size < 0)
    throw invalidStyle(styleName, "fo:font-size", "invalid size");
  return size;
}

function parseAlignment(
  value: string | undefined,
  styleName: string,
): ParagraphStylePropertiesOptions["alignment"] {
  switch (value) {
    case undefined:
      return undefined;
    case "start":
    case "end":
    case "left":
    case "right":
    case "center":
      return value;
    case "justify":
      return "both";
    default:
      throw invalidStyle(styleName, "fo:text-align", `unsupported value: ${value}`);
  }
}

function invalidStyle(styleName: string, elementName: string, reason: string): OdtParseError {
  return new OdtParseError(
    `styles.xml: style ${styleName}: ${reason}`,
    "styles.xml",
    `/office:document-styles/office:styles/style:style[@style:name="${styleName}"]/${elementName}`,
    elementName,
    reason,
  );
}
