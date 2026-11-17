import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  lengthToEmu,
  parseStyleMaps,
  parseStyleProperty,
  removeStylePropertyAttributes,
  styleMapsXml,
  stylePropertyXml,
  xmlElement,
  type XmlAttributes,
} from "@office-open/odf";
import type { StyleFamily, StyleResourceOptions } from "@office-open/odf";
import type {
  AlignmentOptions,
  BorderOptions,
  BorderSideOptions,
  CellFillOptions,
  RichTextRunPropertiesOptions,
  StyleOptions,
  FontOptions,
} from "@office-open/xlsx";
import type { Element } from "@office-open/xml";

import { OdsGenerateError, OdsParseError } from "./error";
import { numFmtDataStyle } from "./numbering";

const dataStyleCaches = new WeakMap<string[], Map<string, string>>();
const styleCaches = new WeakMap<string[], Map<string, string>>();

function cachedName(
  caches: WeakMap<string[], Map<string, string>>,
  styles: string[],
  key: string,
): string | undefined {
  let cache = caches.get(styles);
  if (cache === undefined) {
    cache = new Map();
    caches.set(styles, cache);
  }
  return cache.get(key);
}

function cacheName(
  caches: WeakMap<string[], Map<string, string>>,
  styles: string[],
  key: string,
  name: string,
): void {
  let cache = caches.get(styles);
  if (cache === undefined) {
    cache = new Map();
    caches.set(styles, cache);
  }
  cache.set(key, name);
}

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
  const children = cellStyleChildren(style);
  let dataName: string | undefined;
  if (data) {
    const dataKey = JSON.stringify([data.name, data.children]);
    dataName = cachedName(dataStyleCaches, styles, dataKey);
    if (dataName === undefined) {
      dataName = `N${styles.length + 1}`;
      styles.push(xmlElement(data.name, { "style:name": dataName }, data.children));
      cacheName(dataStyleCaches, styles, dataKey, dataName);
    }
  }
  const styleKey = JSON.stringify([dataName, children]);
  const cachedStyleName = cachedName(styleCaches, styles, styleKey);
  if (cachedStyleName !== undefined) return cachedStyleName;
  const styleName = `ce${styles.length + 1}`;
  styles.push(
    xmlElement(
      "style:style",
      {
        "style:name": styleName,
        "style:family": "table-cell",
        "style:data-style-name": data ? dataName : undefined,
      },
      children,
    ),
  );
  cacheName(styleCaches, styles, styleKey, styleName);
  return styleName;
}

export function styleResourcesXml(overlays: StyleResourceOptions[] | undefined): string[] {
  return (overlays ?? []).map((overlay) =>
    xmlElement("style:style", { "style:name": overlay.name, "style:family": overlay.family }, [
      ...styleMapsXml(overlay.maps),
      ...overlay.properties.map((property) => stylePropertyXml(property, "spreadsheet")),
    ]),
  );
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
    ...underlineAttributes(font),
    ...strikeAttributes(font),
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
  if (Object.values({ ...gnumericFill, ...cellAttributes }).some((value) => value !== undefined))
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

/** Registers one rich-text run style in the ODF automatic styles section. */
export function addTextStyle(
  properties: RichTextRunPropertiesOptions,
  styles: string[],
): string | undefined {
  if (Object.values(properties).every((value) => value === undefined)) return undefined;
  const name = `tx${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "text" }, [
      xmlElement("style:text-properties", {
        "fo:font-family": properties.font,
        "fo:font-size": properties.size === undefined ? undefined : `${properties.size}pt`,
        "fo:font-weight": properties.bold ? "bold" : undefined,
        "fo:font-style": properties.italic ? "italic" : undefined,
        "fo:color": properties.color?.rgb ? odfHex(properties.color.rgb) : undefined,
        ...richUnderlineAttributes(properties.underline),
        "style:text-line-through-style": properties.strike ? "solid" : undefined,
      }),
    ]),
  );
  return name;
}

function richUnderlineAttributes(
  underline: RichTextRunPropertiesOptions["underline"],
): XmlAttributes {
  if (!underline || underline === "none") return {};
  if (underline === "singleAccounting" || underline === "doubleAccounting") {
    throw new OdsGenerateError(
      `ODF cannot represent underline mode: ${underline}`,
      "content.xml",
      "/office:document-content/office:automatic-styles/style:style/style:text-properties",
      "style:text-underline-style",
      underline,
    );
  }
  return {
    "style:text-underline-style": "solid",
    "style:text-underline-type": underline === "double" ? "double" : "single",
  };
}

/** Parses ODF character styles into the shared XLSX rich-text run shape. */
export function parseTextStyles(
  containers: readonly (Element | undefined)[],
): Map<string, RichTextRunPropertiesOptions> {
  const source = new Map<string, { element: Element; parent?: string }>();
  for (const container of containers) {
    for (const style of childrenNamed(container, "style:style")) {
      if (attributeString(style, "style:family") !== "text") continue;
      const name = attributeString(style, "style:name");
      if (name) {
        source.set(name, {
          element: style,
          parent: attributeString(style, "style:parent-style-name"),
        });
      }
    }
  }
  const resolved = new Map<string, RichTextRunPropertiesOptions>();
  const visiting = new Set<string>();
  function resolve(name: string): RichTextRunPropertiesOptions {
    const existing = resolved.get(name);
    if (existing) return existing;
    if (visiting.has(name)) {
      throw new OdsParseError(
        `styles.xml: cyclic text style inheritance: ${name}`,
        "styles.xml",
        `/office:automatic-styles/style:style[@style:name="${name}"]`,
        "style:parent-style-name",
        "cyclic-style-inheritance",
      );
    }
    const entry = source.get(name);
    if (!entry) return {};
    visiting.add(name);
    const parent = entry.parent ? resolve(entry.parent) : {};
    const text = childNamed(entry.element, "style:text-properties");
    const size = attributeString(text, "fo:font-size");
    const color = attributeString(text, "fo:color");
    const underline = attributeString(text, "style:text-underline-style");
    const underlineType = attributeString(text, "style:text-underline-type");
    const strike = attributeString(text, "style:text-line-through-style");
    const own: RichTextRunPropertiesOptions = {
      ...parent,
      ...(attributeString(text, "fo:font-family")
        ? { font: attributeString(text, "fo:font-family")!.replace(/^(['"])(.*)\1$/u, "$2") }
        : {}),
      ...(size?.endsWith("pt") ? { size: Number(size.slice(0, -2)) } : {}),
      ...(attributeString(text, "fo:font-weight") === undefined
        ? {}
        : { bold: attributeString(text, "fo:font-weight") === "bold" }),
      ...(attributeString(text, "fo:font-style") === undefined
        ? {}
        : { italic: attributeString(text, "fo:font-style") === "italic" }),
      ...(color && /^#[0-9a-f]{6}$/iu.test(color)
        ? { color: { rgb: `FF${color.slice(1).toUpperCase()}` } }
        : {}),
      ...(underline && underline !== "none"
        ? { underline: underlineType === "double" ? "double" : "single" }
        : underline === "none"
          ? { underline: "none" }
          : {}),
      ...(strike && strike !== "none"
        ? { strike: true }
        : strike === "none"
          ? { strike: false }
          : {}),
    };
    visiting.delete(name);
    resolved.set(name, own);
    return own;
  }
  for (const name of source.keys()) resolve(name);
  return resolved;
}

/** Maps canonical font underlines to ODF attributes, rejecting ODF-incompatible modes. */
export function underlineAttributes(font: FontOptions | undefined): XmlAttributes {
  const underline = font?.underline;
  const style = font?.underlineStyle ?? (underline ? "solid" : undefined);
  const type = font?.underlineType ?? (underline === "double" ? "double" : "single");
  if (!style) return {};
  if (underline !== "singleAccounting" && underline !== "doubleAccounting") {
    return {
      "style:text-underline-style": style,
      "style:text-underline-type": type,
    };
  }
  throw new OdsGenerateError(
    `ODF cannot represent underline mode: ${underline}`,
    "styles.xml",
    "/office:automatic-styles/style:style/style:text-properties",
    "style:text-underline-type",
    `${style}/${type}`,
  );
}

/** Parses an ODF underline while rejecting values with no canonical equivalent. */
export function parseOdfUnderline(
  element: Element | undefined,
): Pick<FontOptions, "underline" | "underlineStyle" | "underlineType"> | undefined {
  const style = attributeString(element, "style:text-underline-style");
  const legacyStyle = attributeString(element, "fo:underline-style");
  const value = style ?? legacyStyle;
  if (value === undefined || value === "none") return undefined;
  const type = attributeString(element, "style:text-underline-type") ?? "single";
  if (type !== "single" && type !== "double") throwUnderline(value, type);
  if (value === "solid") return { underline: type === "double" ? "double" : true };
  if (value === "dotted" || value === "dash") {
    return {
      underline: type === "double" ? "double" : true,
      underlineStyle: value,
      underlineType: type,
    };
  }
  throwUnderline(value, type);
}

function throwUnderline(value: string, type: string): never {
  throw new OdsParseError(
    `styles.xml: unsupported ODF underline ${value}/${type}`,
    "styles.xml",
    "/office:automatic-styles/style:style/style:text-properties",
    "style:text-underline-style",
    "no canonical workbook mapping",
  );
}

/** Maps canonical strike-through to its exact ODF single-line form. */
export function strikeAttributes(font: FontOptions | undefined): XmlAttributes {
  if (!font?.strike) return {};
  return {
    "style:text-line-through-style": font.strikeStyle ?? "solid",
    "style:text-line-through-type": font.strikeType ?? "single",
  };
}

/** Parses an ODF strike-through while rejecting non-canonical line styles. */
export function parseOdfStrike(
  element: Element | undefined,
): Pick<FontOptions, "strike" | "strikeStyle" | "strikeType"> | undefined {
  const style = attributeString(element, "style:text-line-through-style");
  const legacyStyle = attributeString(element, "fo:text-line-through-style");
  const value = style ?? legacyStyle;
  if (value === undefined || value === "none") return undefined;
  const type = attributeString(element, "style:text-line-through-type") ?? "single";
  if (type !== "single" && type !== "double") throwStrike(value, type);
  if (
    value === "solid" ||
    value === "dotted" ||
    value === "dash" ||
    value === "longDash" ||
    value === "dotDash" ||
    value === "dotDotDash"
  ) {
    return {
      strike: true,
      ...(value === "solid" && type === "single" ? {} : { strikeStyle: value, strikeType: type }),
    };
  }
  throw new OdsParseError(
    `styles.xml: unsupported ODF strike-through ${value}/${type}`,
    "styles.xml",
    "/office:automatic-styles/style:style/style:text-properties",
    "style:text-line-through-style",
    "no canonical workbook mapping",
  );
}

function throwStrike(value: string, type: string): never {
  throw new OdsParseError(
    `styles.xml: unsupported ODF strike-through ${value}/${type}`,
    "styles.xml",
    "/office:automatic-styles/style:style/style:text-properties",
    "style:text-line-through-style",
    "no canonical workbook mapping",
  );
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
  const properties =
    family === "column"
      ? xmlElement("style:table-column-properties", {
          "style:column-width": `${style.width ?? 0}px`,
          "style:use-optimal-column-width": style.hidden ? undefined : "true",
        })
      : xmlElement("style:table-row-properties", {
          "style:row-height": style.height ? `${style.height}pt` : undefined,
        });
  const key = JSON.stringify([family, properties]);
  const cachedNameValue = cachedName(styleCaches, styles, key);
  if (cachedNameValue !== undefined) return cachedNameValue;
  const name = `${family === "column" ? "co" : "ro"}${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": `table-${family}` }, [
      properties,
    ]),
  );
  cacheName(styleCaches, styles, key, name);
  return name;
}

export function addTableStyle(style: DimensionStyle, styles: string[]): string {
  const properties = xmlElement("style:table-properties", {
    "table:display": style.hidden ? false : undefined,
    "gnm:display-formulas": style.showFormulas,
    "gnm:display-col-header": style.showColumnHeaders,
    "gnm:display-row-header": style.showRowHeaders,
  });
  const key = JSON.stringify(["table", properties]);
  const cachedNameValue = cachedName(styleCaches, styles, key);
  if (cachedNameValue !== undefined) return cachedNameValue;
  const name = `ta${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "table" }, [properties]),
  );
  cacheName(styleCaches, styles, key, name);
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
    const columnHidden = dimensionBoolean(column, "style:column-hidden");
    const rowHidden = dimensionBoolean(row, "style:row-hidden");
    const display = dimensionBoolean(table, "table:display");
    const displayFormulas = dimensionBoolean(table, "gnm:display-formulas");
    const displayColumnHeaders = dimensionBoolean(table, "gnm:display-col-header");
    const displayRowHeaders = dimensionBoolean(table, "gnm:display-row-header");
    result.set(name, {
      width: width === undefined ? undefined : lengthToPixels(width, column),
      height: height === undefined ? undefined : lengthToPoints(height, row),
      hidden: columnHidden || rowHidden || display === false || undefined,
      showFormulas: displayFormulas,
      showColumnHeaders: displayColumnHeaders,
      showRowHeaders: displayRowHeaders,
    });
  }
  return result;
}

function dimensionBoolean(element: Element | undefined, name: string): boolean | undefined {
  const value = attributeString(element, name);
  if (value === undefined) return undefined;
  if (value === "true" || value === "1") return true;
  if (value === "false" || value === "0") return false;
  throw new OdsParseError(
    `content.xml: invalid boolean value for ${name}: ${value}`,
    "content.xml",
    `/office:document-content/office:automatic-styles/style:style/${name}`,
    name,
    "invalid boolean value",
  );
}

function lengthToPixels(value: string, element: Element | undefined): number | undefined {
  const emu = lengthToEmu(value);
  if (emu === undefined) return undefined;
  const pixels = emu / 9525;
  if (!Number.isFinite(pixels)) {
    throw new OdsParseError(
      `content.xml: invalid dimension value: ${value}`,
      "content.xml",
      "/office:document-content/office:automatic-styles/style:style",
      element?.name ?? "style:table-column-properties",
      "invalid dimension",
    );
  }
  return pixels;
}

function lengthToPoints(value: string, element: Element | undefined): number | undefined {
  const emu = lengthToEmu(value);
  if (emu === undefined) return undefined;
  const points = emu / 12700;
  if (!Number.isFinite(points)) {
    throw new OdsParseError(
      `content.xml: invalid dimension value: ${value}`,
      "content.xml",
      "/office:document-content/office:automatic-styles/style:style",
      element?.name ?? "style:table-row-properties",
      "invalid dimension",
    );
  }
  return points;
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
    if (!Object.hasOwn(ODS_STYLE_PROPERTIES, family)) {
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
      (child) => child.type === "element" && child.name !== "style:map",
    );
    const properties: StyleResourceOptions["properties"] = [];
    const maps = parseStyleMaps(style);
    for (const property of propertyElements) {
      const sourceAttributes = Object.fromEntries(
        Object.entries(property.attributes ?? {}).filter(
          ([attribute]) => !isMappedDimensionAttribute(property.name ?? "", attribute),
        ),
      );
      if (Object.keys(sourceAttributes).length === 0 && !hasChildElements(property)) continue;
      const parsed = parseStyleProperty(
        { ...property, attributes: sourceAttributes },
        { family, scale: "spreadsheet" },
      );
      if (!parsed.ok) {
        const name = parsed.attributeName ?? parsed.elementName;
        throw new OdsParseError(
          `content.xml: ${path}: ${name}: ${parsed.reason}`,
          "content.xml",
          path,
          name,
          parsed.reason,
        );
      }
      const value = removeStylePropertyAttributes(parsed.value, (attribute) =>
        isMappedDimensionAttribute(property.name ?? "", attribute),
      );
      if (hasStyleProjection(value)) properties.push(value);
    }
    if (properties.length > 0 || maps.length > 0) {
      result.push({ name, family, properties, ...(maps.length > 0 ? { maps } : {}) });
    }
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

function hasChildElements(element: Element): boolean {
  return (element.elements ?? []).some((child) => child.type === "element");
}

function hasStyleProjection(property: StyleResourceOptions["properties"][number]): boolean {
  return Object.keys(property.attributes).length > 0 || (property.children?.length ?? 0) > 0;
}

/** ODF #RRGGBB → xlsx RRGGBB hex. */
export function odfColor(value: string | undefined): string | undefined {
  return value?.startsWith("#") ? value.slice(1) : undefined;
}
