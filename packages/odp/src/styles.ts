import type { TextParagraphPropertiesOptions } from "@office-open/core";
import {
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  lengthToEmu,
  parseStyleMaps,
  parseStyleProperty,
  removeStylePropertyAttributes,
  styleMapsXml,
  stylePropertyXml,
  xmlElement,
} from "@office-open/odf";
import type { StyleFamily, StyleResourceOptions } from "@office-open/odf";
import type { BackgroundOptions, TableCellOptions } from "@office-open/pptx";
import type { Element } from "@office-open/xml";

import { OdpParseError } from "./error";
import type { TextProperties } from "./text-run";

export function addTextStyle(properties: TextProperties, styles: string[]): string | undefined {
  if (Object.values(properties).every((value) => value === undefined)) return undefined;
  const name = `C${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "text" }, [
      xmlElement("style:text-properties", {
        "fo:font-weight": properties.bold ? "bold" : undefined,
        "fo:font-style": properties.italic ? "italic" : undefined,
        "style:text-underline-style": properties.underline ? "solid" : undefined,
        "fo:font-size": properties.size ? `${properties.size}pt` : undefined,
        "fo:color": properties.fill ? `#${properties.fill}` : undefined,
        "fo:font-family": properties.font ? fontFamilyValue(properties.font) : undefined,
      }),
    ]),
  );
  return name;
}

export function addParagraphStyle(
  properties: TextParagraphPropertiesOptions,
  styles: string[],
): string {
  const name = `P${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "paragraph" }, [
      xmlElement("style:paragraph-properties", {
        "fo:text-align": properties.alignment,
        "fo:margin-left":
          properties.marginIndent === undefined ? undefined : emuToLength(properties.marginIndent),
        "fo:margin-right":
          properties.marginRight === undefined ? undefined : emuToLength(properties.marginRight),
        "fo:text-indent":
          properties.indent === undefined ? undefined : emuToLength(properties.indent),
        "fo:margin-top":
          properties.spaceBefore === undefined ? undefined : `${properties.spaceBefore}pt`,
        "fo:margin-bottom":
          properties.spaceAfter === undefined ? undefined : `${properties.spaceAfter}pt`,
        "fo:line-height":
          properties.lineSpacingPercent === undefined
            ? properties.lineSpacingPoints === undefined
              ? undefined
              : `${properties.lineSpacingPoints}pt`
            : `${properties.lineSpacingPercent}%`,
      }),
    ]),
  );
  return name;
}

/** A family name with spaces or commas is quoted, as fo:font-family requires. */
function fontFamilyValue(family: string): string {
  return /[\s,]/.test(family) ? `'${family}'` : family;
}

export function styleResourcesXml(overlays: StyleResourceOptions[] | undefined): string[] {
  return (overlays ?? []).map((overlay) =>
    xmlElement("style:style", { "style:name": overlay.name, "style:family": overlay.family }, [
      ...styleMapsXml(overlay.maps),
      ...overlay.properties.map((property) => stylePropertyXml(property, "emu")),
    ]),
  );
}

/** Column width style — pptx widths are EMU. */
export function addColumnStyle(width: number | string, styles: string[]): string {
  const name = `TC${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "table-column" }, [
      xmlElement("style:table-column-properties", {
        "style:column-width": typeof width === "number" ? emuToLength(width) : width,
      }),
    ]),
  );
  return name;
}

/** Row height style — pptx heights are EMU. */
export function addRowStyle(
  height: number | string | undefined,
  styles: string[],
): string | undefined {
  if (height === undefined) return undefined;
  const name = `TR${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "table-row" }, [
      xmlElement("style:table-row-properties", {
        "style:row-height": typeof height === "number" ? emuToLength(height) : height,
        "style:use-optimal-row-height": "false",
      }),
    ]),
  );
  return name;
}

/** The cell formatting ODF can carry: fill, vertical anchor, padding, and the four side borders. */
export type TableCellStyle = Pick<
  TableCellOptions,
  "fill" | "verticalAlign" | "margins" | "borders"
>;

const CELL_SIDES = ["top", "bottom", "left", "right"] as const;
const CELL_VERTICAL_ALIGN = { top: "top", center: "middle", bottom: "bottom" } as const;

/** RRGGBB of a plain solid fill; themed, transformed, and non-solid fills have no ODF colour form. */
function solidHexColor(fill: TableCellOptions["fill"]): string | undefined {
  const color = typeof fill === "object" && fill.type === "solid" ? fill.color : fill;
  const value =
    typeof color === "string"
      ? color
      : typeof color === "object" && "value" in color && !("transforms" in color)
        ? color.value
        : undefined;
  if (typeof value !== "string") return undefined;
  const hex = value.startsWith("#") ? value.slice(1) : value;
  return /^[0-9a-f]{6}$/i.test(hex) ? hex : undefined;
}

const odfLength = (value: number | string): string =>
  typeof value === "number" ? emuToLength(value) : value;

/** fo:border value of one cell side: width, line style, colour (1pt solid black when unset). */
function cellBorderValue(
  border: NonNullable<TableCellStyle["borders"]>["top"],
): string | undefined {
  if (!border) return undefined;
  const dash = border.dashStyle;
  const line =
    dash === undefined || dash === "solid"
      ? "solid"
      : dash === "dot" || dash === "sysDot"
        ? "dotted"
        : "dashed";
  return `${odfLength(border.width ?? 12700)} ${line} #${solidHexColor(border.color) ?? "000000"}`;
}

/**
 * Cell style: fill, anchor and padding as table-cell-properties, side borders
 * as paragraph-properties — the placement Impress reads cell borders from.
 */
export function addCellStyle(cell: TableCellStyle, styles: string[]): string | undefined {
  const fill = solidHexColor(cell.fill);
  const background =
    fill !== undefined
      ? `#${fill}`
      : typeof cell.fill === "object" && cell.fill.type === "none"
        ? "transparent"
        : undefined;
  const cellProperties = {
    "fo:background-color": background,
    "style:vertical-align":
      cell.verticalAlign && Object.hasOwn(CELL_VERTICAL_ALIGN, cell.verticalAlign)
        ? CELL_VERTICAL_ALIGN[cell.verticalAlign as keyof typeof CELL_VERTICAL_ALIGN]
        : undefined,
    ...Object.fromEntries(
      CELL_SIDES.map((side) => [
        `fo:padding-${side}`,
        cell.margins?.[side] === undefined ? undefined : odfLength(cell.margins[side]),
      ]),
    ),
  };
  const borders = Object.fromEntries(
    CELL_SIDES.map((side) => [`fo:border-${side}`, cellBorderValue(cell.borders?.[side])]),
  );
  const hasCellProperties = Object.values(cellProperties).some((value) => value !== undefined);
  const hasBorders = Object.values(borders).some((value) => value !== undefined);
  if (!hasCellProperties && !hasBorders) return undefined;
  const name = `TE${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "table-cell" }, [
      ...(hasCellProperties ? [xmlElement("style:table-cell-properties", cellProperties)] : []),
      ...(hasBorders ? [xmlElement("style:paragraph-properties", borders)] : []),
    ]),
  );
  return name;
}

/** Row heights (EMU) by table-row style name. */
export function parseRowHeights(container: Element | undefined): Map<string, number> {
  const result = new Map<string, number>();
  for (const style of childrenNamed(container, "style:style")) {
    if (attributeString(style, "style:family") !== "table-row") continue;
    const height = attributeString(
      childNamed(style, "style:table-row-properties"),
      "style:row-height",
    );
    const emu = lengthToEmu(height);
    if (emu !== undefined) result.set(attributeString(style, "style:name") ?? "", emu);
  }
  return result;
}

function parseCellBorder(value: string | undefined): NonNullable<TableCellStyle["borders"]>["top"] {
  const match = value?.match(/^(\S+)\s+(solid|dashed|dotted)\s+#([0-9a-f]{6})$/i);
  const width = lengthToEmu(match?.[1]);
  if (!match || width === undefined) return undefined;
  const line = match[2]!.toLowerCase();
  return {
    width,
    color: match[3]!,
    ...(line === "solid" ? {} : { dashStyle: line === "dotted" ? "dot" : "dash" }),
  };
}

/** Cell formatting by table-cell style name (the form `addCellStyle` writes). */
export function parseCellStyles(container: Element | undefined): Map<string, TableCellStyle> {
  const result = new Map<string, TableCellStyle>();
  for (const style of childrenNamed(container, "style:style")) {
    if (attributeString(style, "style:family") !== "table-cell") continue;
    const cellProperties = childNamed(style, "style:table-cell-properties");
    const paragraphProperties = childNamed(style, "style:paragraph-properties");
    const background = attributeString(cellProperties, "fo:background-color");
    const verticalAlign = Object.entries(CELL_VERTICAL_ALIGN).find(
      ([, odf]) => odf === attributeString(cellProperties, "style:vertical-align"),
    )?.[0] as TableCellStyle["verticalAlign"];
    const margins = Object.fromEntries(
      CELL_SIDES.flatMap((side) => {
        const emu = lengthToEmu(attributeString(cellProperties, `fo:padding-${side}`));
        return emu === undefined ? [] : [[side, emu] as const];
      }),
    );
    const borders = Object.fromEntries(
      CELL_SIDES.flatMap((side) => {
        const border = parseCellBorder(attributeString(paragraphProperties, `fo:border-${side}`));
        return border ? [[side, border] as const] : [];
      }),
    );
    const cell: TableCellStyle = {
      ...(background && /^#[0-9a-f]{6}$/i.test(background)
        ? { fill: background.slice(1) }
        : background === "transparent"
          ? { fill: { type: "none" } }
          : {}),
      ...(verticalAlign ? { verticalAlign } : {}),
      ...(Object.keys(margins).length > 0 ? { margins } : {}),
      ...(Object.keys(borders).length > 0 ? { borders } : {}),
    };
    if (Object.keys(cell).length > 0) result.set(attributeString(style, "style:name") ?? "", cell);
  }
  return result;
}

export function parseTextStyles(container: Element | undefined): Map<string, TextProperties> {
  const result = new Map<string, TextProperties>();
  const paragraphStyles = parseParagraphStyles([container]);
  for (const style of childrenNamed(container, "style:style")) {
    const familyName = attributeString(style, "style:family");
    const styleName = attributeString(style, "style:name") ?? "";
    if (familyName === "paragraph") {
      const paragraph = paragraphStyles.get(styleName);
      if (paragraph) result.set(styleName, { paragraph });
      continue;
    }
    if (familyName !== "text") continue;
    const properties = childNamed(style, "style:text-properties");
    const size = attributeString(properties, "fo:font-size");
    const color = attributeString(properties, "fo:color");
    const family = attributeString(properties, "fo:font-family");
    result.set(styleName, {
      bold: attributeString(properties, "fo:font-weight") === "bold",
      italic: attributeString(properties, "fo:font-style") === "italic",
      underline:
        attributeString(properties, "style:text-underline-style") === "solid"
          ? "single"
          : undefined,
      size: size?.endsWith("pt") ? Number(size.slice(0, -2)) : undefined,
      ...(color && /^#[0-9a-f]{6}$/i.test(color) ? { fill: color.slice(1) } : {}),
      ...(family ? { font: family.replace(/^(['"])(.*)\1$/, "$2") } : {}),
      ...(paragraphStyles.has(attributeString(style, "style:name") ?? "")
        ? { paragraph: paragraphStyles.get(attributeString(style, "style:name") ?? "") }
        : {}),
    });
  }
  return result;
}

export function parseParagraphStyles(
  containers: readonly (Element | undefined)[],
): Map<string, TextParagraphPropertiesOptions> {
  const result = new Map<string, TextParagraphPropertiesOptions>();
  for (const container of containers) {
    for (const style of childrenNamed(container, "style:style")) {
      if (attributeString(style, "style:family") !== "paragraph") continue;
      const name = attributeString(style, "style:name");
      const source = childNamed(style, "style:paragraph-properties");
      if (!name || !source) continue;
      const alignment = attributeString(source, "fo:text-align");
      const margins = [
        ["fo:margin-left", "marginIndent"],
        ["fo:margin-right", "marginRight"],
        ["fo:text-indent", "indent"],
      ] as const;
      const properties: TextParagraphPropertiesOptions = {};
      if (alignment === "start" || alignment === "left") properties.alignment = "left";
      else if (alignment === "center") properties.alignment = "center";
      else if (alignment === "end" || alignment === "right") properties.alignment = "right";
      else if (alignment === "both" || alignment === "justify") properties.alignment = "justify";
      for (const [attribute, field] of margins) {
        const value = lengthToEmu(attributeString(source, attribute));
        if (value !== undefined) properties[field] = value;
      }
      const before = pointValue(attributeString(source, "fo:margin-top"));
      const after = pointValue(attributeString(source, "fo:margin-bottom"));
      if (before !== undefined) properties.spaceBefore = before;
      if (after !== undefined) properties.spaceAfter = after;
      const lineHeight = attributeString(source, "fo:line-height");
      if (lineHeight?.endsWith("%")) {
        const value = Number(lineHeight.slice(0, -1));
        if (Number.isFinite(value)) properties.lineSpacingPercent = value;
      } else {
        const value = pointValue(lineHeight);
        if (value !== undefined) properties.lineSpacingPoints = value;
      }
      if (Object.keys(properties).length > 0) result.set(name, properties);
    }
  }
  return result;
}

export function parseDrawingPageBackground(
  styleResources: readonly StyleResourceOptions[],
  styleName: string | undefined,
): BackgroundOptions | undefined {
  if (!styleName) return undefined;
  const resource = styleResources.find(
    (style) => style.name === styleName && style.family === "drawing-page",
  );
  const properties = resource?.properties.find(
    (property) => property.name === "style:drawing-page-properties",
  );
  const fill = properties?.attributes["draw:fill"];
  const color = properties?.attributes["draw:fill-color"];
  if (fill === "none") return { fill: { type: "none" } };
  if (fill !== "solid" || typeof color !== "string" || !/^#[0-9a-f]{6}$/iu.test(color))
    return undefined;
  return { fill: { type: "solid", color: color.slice(1) } };
}

function pointValue(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  if (value.endsWith("pt")) {
    const points = Number(value.slice(0, -2));
    return Number.isFinite(points) ? points : undefined;
  }
  const emu = lengthToEmu(value);
  return emu === undefined ? undefined : emu / 12700;
}

const ODP_STYLE_PROPERTIES: Record<string, string[]> = {
  paragraph: ["style:paragraph-properties", "loext:graphic-properties", "style:text-properties"],
  text: ["style:text-properties"],
  graphic: [
    "style:graphic-properties",
    "loext:graphic-properties",
    "style:paragraph-properties",
    "style:text-properties",
  ],
  presentation: [
    "style:graphic-properties",
    "loext:graphic-properties",
    "style:paragraph-properties",
    "style:text-properties",
  ],
  table: ["style:table-properties"],
  "table-column": ["style:table-column-properties"],
  "table-row": ["style:table-row-properties"],
  "table-cell": [
    "style:table-cell-properties",
    "style:graphic-properties",
    "loext:graphic-properties",
    "style:paragraph-properties",
    "style:text-properties",
  ],
  section: ["style:section-properties"],
  "drawing-page": ["style:drawing-page-properties"],
  ruby: ["style:ruby-properties"],
};

export function parseStyleResources(container: Element | undefined): StyleResourceOptions[] {
  const result: StyleResourceOptions[] = [];
  for (const style of childrenNamed(container, "style:style")) {
    const name = attributeString(style, "style:name") ?? "";
    const family = attributeString(style, "style:family") as StyleFamily;
    const path = `/office:document-content/office:automatic-styles/style:style[@style:name="${name}"]`;
    const allowed = Object.hasOwn(ODP_STYLE_PROPERTIES, family)
      ? ODP_STYLE_PROPERTIES[family]
      : undefined;
    if (!allowed) {
      const reason = "unknown style family";
      throw new OdpParseError(
        `content.xml: ${path}: ${family}: ${reason}`,
        "content.xml",
        path,
        family,
        reason,
      );
    }
    const properties: StyleResourceOptions["properties"] = [];
    const maps = parseStyleMaps(style);
    for (const property of style.elements ?? []) {
      if (property.type !== "element" || property.name === "style:map") continue;
      const parsed = parseStyleProperty(property, { family, scale: "emu" });
      if (!parsed.ok) {
        const name = parsed.attributeName ?? parsed.elementName;
        throw new OdpParseError(
          `content.xml: ${path}: ${name}: ${parsed.reason}`,
          "content.xml",
          path,
          name,
          parsed.reason,
        );
      }
      const value = removeStylePropertyAttributes(parsed.value, (attribute) =>
        isMappedTextAttribute(property.name ?? "", attribute),
      );
      if (hasStyleProjection(value)) properties.push(value);
    }
    if (properties.length > 0 || maps.length > 0) {
      result.push({ name, family, properties, ...(maps.length > 0 ? { maps } : {}) });
    }
  }
  return result;
}

function isMappedTextAttribute(element: string, attribute: string): boolean {
  return (
    element === "style:text-properties" &&
    [
      "fo:font-weight",
      "fo:font-style",
      "style:text-underline-style",
      "fo:font-size",
      "fo:color",
      "fo:font-family",
    ].includes(attribute)
  );
}

function hasStyleProjection(property: StyleResourceOptions["properties"][number]): boolean {
  return Object.keys(property.attributes).length > 0 || (property.children?.length ?? 0) > 0;
}
