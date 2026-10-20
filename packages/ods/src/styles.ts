import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  lengthToEmu,
  xmlElement,
  type XmlAttributes,
} from "@office-open/odf";
import type { StyleFamily, StyleResourceOptions } from "@office-open/odf";
import type {
  AlignmentOptions,
  BorderOptions,
  BorderSideOptions,
  CellFillOptions,
  StyleOptions,
} from "@office-open/xlsx";
import type { Element } from "@office-open/xml";

import { OdsParseError } from "./error";
import { numFmtDataStyle } from "./numbering";

export interface DimensionStyle {
  width?: number;
  height?: number;
  hidden?: boolean;
  showFormulas?: boolean;
  showRowHeaders?: boolean;
  showColumnHeaders?: boolean;
}

/** Registers a table-cell style (numFmt data style + font/fill); returns its name. */
export function addCellStyle(style: StyleOptions, styles: string[]): string {
  const data = style.numFmt ? numFmtDataStyle(style.numFmt) : undefined;
  const dataName = `N${styles.length + 1}`;
  if (data) styles.push(xmlElement(data.name, { "style:name": dataName }, data.children));
  const styleName = `ce${styles.length + 1}`;
  styles.push(
    xmlElement(
      "style:style",
      {
        "style:name": styleName,
        "style:family": "table-cell",
        "style:data-style-name": data ? dataName : undefined,
      },
      cellStyleChildren(style),
    ),
  );
  return styleName;
}

export function styleResourcesXml(overlays: StyleResourceOptions[] | undefined): string[] {
  return (overlays ?? []).map((overlay) =>
    xmlElement(
      "style:style",
      { "style:name": overlay.name, "style:family": overlay.family },
      overlay.properties.map((property) =>
        xmlElement(property.name, resourceAttributesXml(property.attributes)),
      ),
    ),
  );
}

function resourceAttributesXml(
  attributes: Record<string, string | number | boolean>,
): Record<string, string | number | boolean> {
  return Object.fromEntries(
    Object.entries(attributes).map(([name, value]) => [
      name,
      typeof value === "string"
        ? value
        : isPercentAttribute(name)
          ? `${value}%`
          : typeof value === "boolean"
            ? value
            : name === "style:column-width"
              ? emuToLength(value * 9525)
              : name === "style:row-height" || name === "style:min-row-height"
                ? `${value}pt`
                : value,
    ]),
  );
}

function isPercentAttribute(name: string): boolean {
  return /(line-height|rel-width|rel-height|opacity|transparency)/.test(name);
}

/** Font/fill properties as ODF style children; empty array when styleless. */
export function cellStyleChildren(style: StyleOptions): string[] {
  const font = style.font;
  const fill = style.fill;
  const alignment = style.alignment;
  const textAttributes = {
    "fo:color": odfHex(font?.color?.rgb),
    "fo:font-size": font?.size !== undefined ? `${font.size}pt` : undefined,
    "fo:font-weight": font?.bold ? "bold" : undefined,
    "fo:font-style": font?.italic ? "italic" : undefined,
    "fo:underline-style": font?.underline ? "solid" : undefined,
    "fo:text-line-through-style": font?.strike ? "solid" : undefined,
  };
  const cellAttributes = {
    "fo:background-color": odfHex(solidFillColor(fill)),
    "style:vertical-align": odfVertical(alignment?.vertical),
    "fo:wrap-option": alignment?.wrapText ? "wrap" : undefined,
    ...borderAttributes(style.border),
  };
  const paragraphAttributes = { "fo:text-align": odfHorizontal(alignment?.horizontal) };
  const children: string[] = [];
  const gnumericFill = gnumericFillAttributes(fill);
  if (Object.values(cellAttributes).some((value) => value !== undefined))
    children.push(
      xmlElement("style:table-cell-properties", {
        ...gnumericFill,
        ...cellAttributes,
        "fo:background-color":
          gnumericFill["fo:background-color"] ?? cellAttributes["fo:background-color"],
      }),
    );
  if (Object.values(paragraphAttributes).some((value) => value !== undefined))
    children.push(xmlElement("style:paragraph-properties", paragraphAttributes));
  if (Object.values(textAttributes).some((value) => value !== undefined))
    children.push(xmlElement("style:text-properties", textAttributes));
  return children;
}

/** XLSX cell borders → ODF border shorthand with explicit RGB colors. */
export function borderAttributes(border: BorderSideOptions | undefined): XmlAttributes {
  const attributes = {
    "fo:border-top": odfBorder(border?.top),
    "fo:border-bottom": odfBorder(border?.bottom),
    "fo:border-left": odfBorder(border?.left),
    "fo:border-right": odfBorder(border?.right),
    // xlsx diagonal flags map onto ODF's two named diagonal directions.
    "style:diagonal-tl-br": border?.diagonalDown ? odfBorder(border.diagonal) : undefined,
    "style:diagonal-bl-tr": border?.diagonalUp ? odfBorder(border.diagonal) : undefined,
  };
  return Object.values(attributes).some((value) => value !== undefined) ? attributes : {};
}

/** XLSX border side → CSS-style ODF shorthand (`width style color`). */
export function odfBorder(side: BorderOptions | undefined): string | undefined {
  if (!side || side.style === undefined) return undefined;
  if (side.style === "none") return "none";
  const width = side.style === "thick" ? "2.5pt" : mediumBorder(side.style) ? "1pt" : "0.5pt";
  const style = odfLineStyle(side.style);
  return [width, style, odfHex(side.color?.rgb)].filter(Boolean).join(" ");
}

export function mediumBorder(style: BorderOptions["style"]): boolean {
  return (
    style === "medium" ||
    style === "mediumDashed" ||
    style === "mediumDashDot" ||
    style === "mediumDashDotDot"
  );
}

/** Closest ODF/CSS line style; compound XLSX dashes become dashed. */
export function odfLineStyle(style: NonNullable<BorderOptions["style"]>): string {
  if (
    style === "dashed" ||
    style === "mediumDashed" ||
    style === "dashDot" ||
    style === "mediumDashDot"
  )
    return "dashed";
  if (style === "dotted" || style === "dashDotDot" || style === "mediumDashDotDot") return "dotted";
  if (style === "double") return "double";
  return "solid";
}

/** ODF border shorthand → the closest typed XLSX border side. */
export function parseBorder(value: string | undefined): BorderOptions | undefined {
  if (!value || value === "none") return value === "none" ? { style: "none" } : undefined;
  const match = /^(-?\d+(?:\.\d+)?(?:cm|mm|pt|pc|in|px))\s+(\S+)(?:\s+#([0-9a-fA-F]{6}))?$/.exec(
    value,
  );
  if (!match) return undefined;
  const [, width, lineStyle, color] = match;
  const points = lengthToEmu(width)! / 12700;
  const style = ((): BorderOptions["style"] => {
    if (lineStyle === "dashed") return points >= 0.75 ? "mediumDashed" : "dashed";
    if (lineStyle === "dotted") return "dotted";
    if (lineStyle === "double") return "double";
    if (points >= 1.75) return "thick";
    if (points >= 0.75) return "medium";
    if (points < 0.25) return "hair";
    return "thin";
  })();
  return { style, color: color ? { rgb: `FF${color.toUpperCase()}` } : undefined };
}

/** xlsx horizontal → ODF fo:text-align. */
export function odfHorizontal(value: AlignmentOptions["horizontal"]): string | undefined {
  if (value === "left") return "start";
  if (value === "right") return "end";
  if (value === "center" || value === "justify") return value;
  return undefined;
}

/** xlsx vertical → ODF style:vertical-align. */
export function odfVertical(value: AlignmentOptions["vertical"]): string | undefined {
  if (value === "center") return "middle";
  if (value === "top" || value === "bottom") return value;
  return undefined;
}

/** Solid-fill foreground color; pattern/gradient fills keep their typed shape. */
export function solidFillColor(fill: CellFillOptions | undefined): string | undefined {
  return fill?.type === undefined || fill.type === "solid" ? fill?.foregroundColor?.rgb : undefined;
}

export function parseGnumericFill(properties: Element): CellFillOptions | undefined {
  const background = attributeString(properties, "gnm:background-colour");
  const foreground = attributeString(properties, "gnm:pattern-colour");
  const pattern = attributeNumber(properties, "gnm:pattern");
  if (background === undefined && foreground === undefined && pattern === undefined)
    return undefined;
  if (pattern === undefined || !Number.isInteger(pattern) || pattern < 0 || pattern > 24) {
    throw new OdsParseError(
      "Invalid Gnumeric cell fill pattern",
      "content.xml",
      "/office:document-content/office:automatic-styles/style:style/style:table-cell-properties",
      "gnm:pattern",
      "invalid-fill-pattern",
    );
  }
  const kind = GNUMERIC_PATTERN_FILL[pattern]!;
  const foregroundColor = gnumericColor(foreground, "gnm:pattern-colour") ?? { rgb: "FF000000" };
  const backgroundColor = gnumericColor(background, "gnm:background-colour");
  return {
    type: kind.patternType === "solid" ? "solid" : "pattern",
    patternType: kind.patternType,
    legacyPatternType: kind.legacyPatternType,
    foregroundColor,
    ...(backgroundColor ? { backgroundColor } : {}),
  };
}

function gnumericColor(value: string | undefined, name: string): { rgb: string } | undefined {
  if (value === undefined) return undefined;
  if (!/^#[0-9a-f]{6}$/iu.test(value)) {
    throw new OdsParseError(
      "Invalid Gnumeric cell fill color",
      "content.xml",
      "/office:document-content/office:automatic-styles/style:style/style:table-cell-properties",
      name,
      "invalid-fill-color",
    );
  }
  return { rgb: `FF${value.slice(1).toUpperCase()}` };
}

export function gnumericFillAttributes(fill: CellFillOptions | undefined): XmlAttributes {
  const index = GNUMERIC_PATTERN_INDEX.get(
    fill?.legacyPatternType ?? fill?.patternType ?? (fill?.type === "solid" ? "solid" : undefined),
  );
  if (index === undefined) return {};
  const foreground = fill?.foregroundColor;
  const background = fill?.backgroundColor;
  return {
    "gnm:pattern": index,
    "gnm:pattern-colour": foreground?.rgb ? odfHex(foreground.rgb) : "#000000",
    "gnm:background-colour": background?.rgb ? odfHex(background.rgb) : undefined,
    "fo:background-color":
      index === 0 ? "transparent" : background?.rgb ? odfHex(background.rgb) : undefined,
  };
}

/** xlsx hex (RRGGBB or AARRGGBB) → ODF #RRGGBB. */
export function odfHex(hex: string | undefined): string | undefined {
  return hex ? `#${hex.slice(-6)}` : undefined;
}

export function addDimensionStyle(
  style: DimensionStyle,
  styles: string[],
  family: "column" | "row",
): string {
  const name = `${family === "column" ? "co" : "ro"}${styles.length + 1}`;
  const properties =
    family === "column"
      ? xmlElement("style:table-column-properties", {
          "style:column-width": `${style.width ?? 0}px`,
          "style:use-optimal-column-width": style.hidden ? undefined : "true",
        })
      : xmlElement("style:table-row-properties", {
          "style:row-height": style.height ? `${style.height}pt` : undefined,
        });
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": `table-${family}` }, [
      properties,
    ]),
  );
  return name;
}

export function addTableStyle(style: DimensionStyle, styles: string[]): string {
  const name = `ta${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "table" }, [
      xmlElement("style:table-properties", {
        "table:display": style.hidden ? false : undefined,
        "gnm:display-formulas": style.showFormulas,
        "gnm:display-col-header": style.showColumnHeaders,
        "gnm:display-row-header": style.showRowHeaders,
      }),
    ]),
  );
  return name;
}

export function parseDimensionStyles(container: Element | undefined): Map<string, DimensionStyle> {
  const result = new Map<string, DimensionStyle>();
  for (const style of childrenNamed(container, "style:style")) {
    const name = attributeString(style, "style:name") ?? "";
    const column = childNamed(style, "style:table-column-properties");
    const row = childNamed(style, "style:table-row-properties");
    const width = attributeString(column, "style:column-width");
    const height = attributeString(row, "style:row-height");
    const table = childNamed(style, "style:table-properties");
    const showColumnHeaders = attributeString(table, "gnm:display-col-header");
    const showRowHeaders = attributeString(table, "gnm:display-row-header");
    if (
      showColumnHeaders !== undefined &&
      showRowHeaders !== undefined &&
      showColumnHeaders !== showRowHeaders
    ) {
      const reason = "Gnumeric header visibility differs";
      throw new OdsParseError(
        `content.xml: /office:document-content/office:automatic-styles/style:style[@style:name="${name}"]: ${reason}`,
        "content.xml",
        `/office:document-content/office:automatic-styles/style:style[@style:name="${name}"]`,
        "gnm:display-col-header",
        "unsupported-header-visibility",
      );
    }
    result.set(name, {
      width: width?.endsWith("px") ? Number(width.slice(0, -2)) : undefined,
      height: height?.endsWith("pt") ? Number(height.slice(0, -2)) : undefined,
      hidden:
        attributeString(column, "style:column-hidden") === "true" ||
        attributeString(row, "style:row-hidden") === "true" ||
        attributeString(table, "table:display") === "false" ||
        undefined,
      showFormulas:
        attributeString(table, "gnm:display-formulas") === undefined
          ? undefined
          : attributeString(table, "gnm:display-formulas") === "true",
      showColumnHeaders: showColumnHeaders === undefined ? undefined : showColumnHeaders === "true",
      showRowHeaders: showRowHeaders === undefined ? undefined : showRowHeaders === "true",
    });
  }
  return result;
}

const ODS_STYLE_PROPERTIES: Record<string, string[]> = {
  paragraph: ["style:paragraph-properties", "loext:graphic-properties", "style:text-properties"],
  text: ["style:text-properties"],
  graphic: [
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
    "style:paragraph-properties",
    "style:text-properties",
  ],
  section: ["style:section-properties"],
  "drawing-page": ["style:drawing-page-properties"],
  ruby: ["style:ruby-properties"],
};

const GNUMERIC_PATTERN_FILL: Record<
  number,
  {
    patternType?: CellFillOptions["patternType"];
    legacyPatternType?: string;
  }
> = {
  0: { patternType: "none" },
  1: { patternType: "solid" },
  2: { patternType: "mediumGray" },
  3: { patternType: "darkGray" },
  4: { patternType: "lightGray" },
  5: { patternType: "gray125" },
  6: { patternType: "gray0625" },
  7: { patternType: "darkHorizontal" },
  8: { patternType: "darkVertical" },
  9: { patternType: "darkDown" },
  10: { patternType: "darkUp" },
  11: { patternType: "darkTrellis" },
  12: { legacyPatternType: "gnumericThickDiagonalCross" },
  13: { patternType: "lightHorizontal" },
  14: { patternType: "lightVertical" },
  15: { patternType: "lightUp" },
  16: { patternType: "lightDown" },
  17: { patternType: "lightGrid" },
  18: { patternType: "lightTrellis" },
  19: { legacyPatternType: "gnumericSmallCircles" },
  20: { legacyPatternType: "gnumericSemiCircles" },
  21: { legacyPatternType: "gnumericThatch" },
  22: { legacyPatternType: "gnumericLargeCircles" },
  23: { legacyPatternType: "gnumericBricks" },
  24: { legacyPatternType: "gnumericForegroundSolid" },
};

const GNUMERIC_PATTERN_INDEX = new Map(
  Object.entries(GNUMERIC_PATTERN_FILL).map(([index, kind]) => [
    kind.legacyPatternType ?? kind.patternType,
    Number(index),
  ]),
);

export function parseStyleResources(container: Element | undefined): StyleResourceOptions[] {
  const result: StyleResourceOptions[] = [];
  for (const style of childrenNamed(container, "style:style")) {
    const name = attributeString(style, "style:name") ?? "";
    const family = attributeString(style, "style:family") as StyleFamily;
    const path = `/office:document-content/office:automatic-styles/style:style[@style:name="${name}"]`;
    if (!ODS_STYLE_PROPERTIES[family]) {
      const reason = "unknown style family";
      throw new OdsParseError(
        `content.xml: ${path}: ${family}: ${reason}`,
        "content.xml",
        path,
        family,
        reason,
      );
    }
    const propertyElements = (style.elements ?? []).filter(
      (child) =>
        child.type === "element" && child.name !== "style:tab-stops" && child.name !== "style:map",
    );
    const properties: StyleResourceOptions["properties"] = [];
    for (const property of propertyElements) {
      const propertyName = property.name ?? "";
      if (!ODS_STYLE_PROPERTIES[family]!.includes(propertyName)) {
        const reason = "unknown style property element";
        throw new OdsParseError(
          `content.xml: ${path}: ${propertyName}: ${reason}`,
          "content.xml",
          path,
          propertyName,
          reason,
        );
      }
      const attributes: Record<string, string | number | boolean> = {};
      for (const [attribute, value] of Object.entries(property.attributes ?? {})) {
        if (isMappedDimensionAttribute(property.name ?? "", attribute)) continue;
        if (
          !/^(fo|style|text|draw|table|svg|dr3d|presentation|smil|officeooo|loext|css3t|tableooo):/.test(
            attribute,
          )
        ) {
          const reason = "attribute has no ODF RNG mapping";
          throw new OdsParseError(
            `content.xml: ${path}: ${attribute}: ${reason}`,
            "content.xml",
            path,
            attribute,
            reason,
          );
        }
        attributes[attribute] = resourceValue(attribute, String(value));
      }
      if (Object.keys(attributes).length > 0)
        properties.push({ name: property.name ?? "", attributes });
    }
    if (properties.length > 0) result.push({ name, family, properties });
  }
  return result;
}

function isMappedDimensionAttribute(element: string, attribute: string): boolean {
  return (
    (element === "style:table-column-properties" &&
      ["style:column-width", "style:use-optimal-column-width"].includes(attribute)) ||
    (element === "style:table-row-properties" &&
      [
        "style:row-height",
        "style:use-optimal-row-height",
        "style:column-hidden",
        "style:row-hidden",
      ].includes(attribute)) ||
    (element === "style:table-properties" &&
      [
        "table:display",
        "gnm:display-formulas",
        "gnm:display-col-header",
        "gnm:display-row-header",
      ].includes(attribute)) ||
    (element === "style:table-cell-properties" && attribute.startsWith("gnm:"))
  );
}

function resourceValue(name: string, value: string): string | number {
  if (/^\d+(?:\.\d+)?%$/.test(value)) return Number.parseFloat(value);
  if (name === "style:column-width") return Math.round((lengthToEmu(value) ?? 0) / 9525);
  if (name === "style:row-height" || name === "style:min-row-height") {
    const emu = lengthToEmu(value);
    return emu === undefined ? value : emu / 12700;
  }
  if (/(angle|rotation)/.test(name) && Number.isFinite(Number(value))) return Number(value);
  return value;
}

/** ODF #RRGGBB → xlsx RRGGBB hex. */
export function odfColor(value: string | undefined): string | undefined {
  return value?.startsWith("#") ? value.slice(1) : undefined;
}
