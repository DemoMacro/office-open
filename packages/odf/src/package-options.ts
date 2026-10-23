import type { DataType } from "@office-open/core";
import type { Element } from "@office-open/xml";

import type { OcfManifestOptions } from "./container/manifest";
import { emuToLength, lengthToEmu, type XmlAttributes } from "./runtime/xml";

/** Style family allowed by ODF automatic/common style containers. */
export type StyleFamily =
  | "paragraph"
  | "text"
  | "graphic"
  | "table"
  | "table-column"
  | "table-row"
  | "table-cell"
  | "section"
  | "page-layout"
  | "drawing-page"
  | "presentation"
  | "ruby";

/** Source-only OCF member carried outside modeled canonical content. */
export interface PackageMemberOptions {
  path: string;
  mediaType?: string;
  data: DataType;
}

type StylePropertyValue = string | number | boolean;

const PARAGRAPH_STYLE_ATTRIBUTES = [
  "fo:background-color",
  "fo:border",
  "fo:border-bottom",
  "fo:border-left",
  "fo:border-right",
  "fo:border-top",
  "fo:break-after",
  "fo:break-before",
  "fo:hyphenation-keep",
  "fo:hyphenation-ladder-count",
  "fo:keep-together",
  "fo:keep-with-next",
  "fo:line-height",
  "fo:margin",
  "fo:margin-bottom",
  "fo:margin-left",
  "fo:margin-right",
  "fo:margin-top",
  "fo:orphans",
  "fo:padding",
  "fo:padding-bottom",
  "fo:padding-left",
  "fo:padding-right",
  "fo:padding-top",
  "fo:text-align",
  "fo:text-align-last",
  "fo:text-indent",
  "fo:widows",
  "style:auto-text-indent",
  "style:background-transparency",
  "style:border-line-width",
  "style:border-line-width-bottom",
  "style:border-line-width-left",
  "style:border-line-width-right",
  "style:border-line-width-top",
  "style:contextual-spacing",
  "style:font-independent-line-spacing",
  "style:join-border",
  "style:justify-single-word",
  "style:line-break",
  "style:line-height-at-least",
  "style:line-spacing",
  "style:page-number",
  "style:punctuation-wrap",
  "style:register-true",
  "style:shadow",
  "style:snap-to-layout-grid",
  "style:tab-stop-distance",
  "style:text-autospace",
  "style:vertical-align",
  "style:writing-mode",
  "style:writing-mode-automatic",
  "text:line-number",
  "text:number-lines",
] as const;

const TEXT_STYLE_ATTRIBUTES = [
  "fo:background-color",
  "fo:color",
  "fo:country",
  "fo:font-family",
  "fo:font-size",
  "fo:font-style",
  "fo:font-variant",
  "fo:font-weight",
  "fo:hyphenate",
  "fo:hyphenation-push-char-count",
  "fo:hyphenation-remain-char-count",
  "fo:language",
  "fo:letter-spacing",
  "fo:script",
  "fo:text-shadow",
  "fo:text-transform",
  "style:country-asian",
  "style:country-complex",
  "style:font-charset",
  "style:font-charset-asian",
  "style:font-charset-complex",
  "style:font-family-asian",
  "style:font-family-complex",
  "style:font-family-generic",
  "style:font-family-generic-asian",
  "style:font-family-generic-complex",
  "style:font-name",
  "style:font-name-asian",
  "style:font-name-complex",
  "style:font-pitch",
  "style:font-pitch-asian",
  "style:font-pitch-complex",
  "style:font-relief",
  "style:font-size-asian",
  "style:font-size-complex",
  "style:font-size-rel",
  "style:font-size-rel-asian",
  "style:font-size-rel-complex",
  "style:font-style-asian",
  "style:font-style-complex",
  "style:font-style-name",
  "style:font-style-name-asian",
  "style:font-style-name-complex",
  "style:font-weight-asian",
  "style:font-weight-complex",
  "style:language-asian",
  "style:language-complex",
  "style:letter-kerning",
  "style:rfc-language-tag",
  "style:rfc-language-tag-asian",
  "style:rfc-language-tag-complex",
  "style:script-asian",
  "style:script-complex",
  "style:script-type",
  "style:text-blinking",
  "style:text-combine",
  "style:text-combine-end-char",
  "style:text-combine-start-char",
  "style:text-emphasize",
  "style:text-line-through-color",
  "style:text-line-through-mode",
  "style:text-line-through-style",
  "style:text-line-through-text",
  "style:text-line-through-text-style",
  "style:text-line-through-type",
  "style:text-line-through-width",
  "style:text-outline",
  "style:text-overline-color",
  "style:text-overline-mode",
  "style:text-overline-style",
  "style:text-overline-type",
  "style:text-overline-width",
  "style:text-position",
  "style:text-rotation-angle",
  "style:text-rotation-scale",
  "style:text-scale",
  "style:text-underline-color",
  "style:text-underline-mode",
  "style:text-underline-style",
  "style:text-underline-type",
  "style:text-underline-width",
  "style:use-window-font-color",
  "text:condition",
  "text:display",
] as const;

const GRAPHIC_STYLE_ATTRIBUTES = [
  "dr3d:ambient-color",
  "dr3d:back-scale",
  "dr3d:backface-culling",
  "dr3d:close-back",
  "dr3d:close-front",
  "dr3d:depth",
  "dr3d:diffuse-color",
  "dr3d:edge-rounding",
  "dr3d:edge-rounding-mode",
  "dr3d:emissive-color",
  "dr3d:end-angle",
  "dr3d:horizontal-segments",
  "dr3d:lighting-mode",
  "dr3d:normals-direction",
  "dr3d:normals-kind",
  "dr3d:shadow",
  "dr3d:shininess",
  "dr3d:specular-color",
  "dr3d:texture-filter",
  "dr3d:texture-generation-mode-x",
  "dr3d:texture-generation-mode-y",
  "dr3d:texture-kind",
  "dr3d:texture-mode",
  "dr3d:vertical-segments",
  "draw:auto-grow-height",
  "draw:auto-grow-width",
  "draw:blue",
  "draw:caption-angle",
  "draw:caption-angle-type",
  "draw:caption-escape",
  "draw:caption-escape-direction",
  "draw:caption-fit-line-length",
  "draw:caption-gap",
  "draw:caption-line-length",
  "draw:caption-type",
  "draw:color-inversion",
  "draw:color-mode",
  "draw:contrast",
  "draw:decimal-places",
  "draw:draw-aspect",
  "draw:end-guide",
  "draw:end-line-spacing-horizontal",
  "draw:end-line-spacing-vertical",
  "draw:fill",
  "draw:fill-color",
  "draw:fill-gradient-name",
  "draw:fill-hatch-name",
  "draw:fill-hatch-solid",
  "draw:fill-image-height",
  "draw:fill-image-name",
  "draw:fill-image-ref-point",
  "draw:fill-image-ref-point-x",
  "draw:fill-image-ref-point-y",
  "draw:fill-image-width",
  "draw:fit-to-contour",
  "draw:fit-to-size",
  "draw:frame-display-border",
  "draw:frame-display-scrollbar",
  "draw:frame-margin-horizontal",
  "draw:frame-margin-vertical",
  "draw:gamma",
  "draw:gradient-step-count",
  "draw:green",
  "draw:guide-distance",
  "draw:guide-overhang",
  "draw:image-opacity",
  "draw:line-distance",
  "draw:luminance",
  "draw:marker-end",
  "draw:marker-end-center",
  "draw:marker-end-width",
  "draw:marker-start",
  "draw:marker-start-center",
  "draw:marker-start-width",
  "draw:measure-align",
  "draw:measure-vertical-align",
  "draw:ole-draw-aspect",
  "draw:opacity",
  "draw:opacity-name",
  "draw:parallel",
  "draw:placing",
  "draw:red",
  "draw:secondary-fill-color",
  "draw:shadow",
  "draw:shadow-color",
  "draw:shadow-offset-x",
  "draw:shadow-offset-y",
  "draw:shadow-opacity",
  "draw:show-unit",
  "draw:start-guide",
  "draw:start-line-spacing-horizontal",
  "draw:start-line-spacing-vertical",
  "draw:stroke",
  "draw:stroke-dash",
  "draw:stroke-dash-names",
  "draw:stroke-linejoin",
  "draw:symbol-color",
  "draw:textarea-horizontal-align",
  "draw:textarea-vertical-align",
  "draw:tile-repeat-offset",
  "draw:unit",
  "draw:visible-area-height",
  "draw:visible-area-left",
  "draw:visible-area-top",
  "draw:visible-area-width",
  "draw:wrap-influence-on-position",
  "fo:background-color",
  "fo:border",
  "fo:border-bottom",
  "fo:border-left",
  "fo:border-right",
  "fo:border-top",
  "fo:clip",
  "fo:margin",
  "fo:margin-bottom",
  "fo:margin-left",
  "fo:margin-right",
  "fo:margin-top",
  "fo:max-height",
  "fo:max-width",
  "fo:min-height",
  "fo:min-width",
  "fo:padding",
  "fo:padding-bottom",
  "fo:padding-left",
  "fo:padding-right",
  "fo:padding-top",
  "fo:wrap-option",
  "style:background-transparency",
  "style:border-line-width",
  "style:border-line-width-bottom",
  "style:border-line-width-left",
  "style:border-line-width-right",
  "style:border-line-width-top",
  "style:editable",
  "style:flow-with-text",
  "style:horizontal-pos",
  "style:horizontal-rel",
  "style:mirror",
  "style:number-wrapped-paragraphs",
  "style:overflow-behavior",
  "style:print-content",
  "style:protect",
  "style:rel-height",
  "style:rel-width",
  "style:repeat",
  "style:run-through",
  "style:shadow",
  "style:shrink-to-fit",
  "style:vertical-pos",
  "style:vertical-rel",
  "style:wrap",
  "style:wrap-contour",
  "style:wrap-contour-mode",
  "style:wrap-dynamic-threshold",
  "style:writing-mode",
  "svg:fill-rule",
  "svg:height",
  "svg:stroke-color",
  "svg:stroke-linecap",
  "svg:stroke-opacity",
  "svg:stroke-width",
  "svg:width",
  "svg:x",
  "svg:y",
  "text:anchor-page-number",
  "text:anchor-type",
  "text:animation",
  "text:animation-delay",
  "text:animation-direction",
  "text:animation-repeat",
  "text:animation-start-inside",
  "text:animation-steps",
  "text:animation-stop-inside",
] as const;

const TABLE_STYLE_ATTRIBUTES = [
  "fo:background-color",
  "fo:break-after",
  "fo:break-before",
  "fo:keep-with-next",
  "fo:margin",
  "fo:margin-bottom",
  "fo:margin-left",
  "fo:margin-right",
  "fo:margin-top",
  "style:may-break-between-rows",
  "style:page-number",
  "style:rel-width",
  "style:shadow",
  "style:width",
  "style:writing-mode",
  "table:align",
  "table:border-model",
  "table:display",
  "table:tab-color",
] as const;
const TABLE_COLUMN_STYLE_ATTRIBUTES = [
  "fo:break-after",
  "fo:break-before",
  "style:column-width",
  "style:rel-column-width",
  "style:use-optimal-column-width",
] as const;
const TABLE_ROW_STYLE_ATTRIBUTES = [
  "fo:background-color",
  "fo:break-after",
  "fo:break-before",
  "fo:keep-together",
  "style:min-row-height",
  "style:row-height",
  "style:use-optimal-row-height",
] as const;
const TABLE_CELL_STYLE_ATTRIBUTES = [
  "fo:background-color",
  "fo:border",
  "fo:border-bottom",
  "fo:border-left",
  "fo:border-right",
  "fo:border-top",
  "fo:padding",
  "fo:padding-bottom",
  "fo:padding-left",
  "fo:padding-right",
  "fo:padding-top",
  "fo:wrap-option",
  "style:border-line-width",
  "style:border-line-width-bottom",
  "style:border-line-width-left",
  "style:border-line-width-right",
  "style:border-line-width-top",
  "style:cell-protect",
  "style:decimal-places",
  "style:diagonal-bl-tr",
  "style:diagonal-bl-tr-widths",
  "style:diagonal-tl-br",
  "style:diagonal-tl-br-widths",
  "style:direction",
  "style:glyph-orientation-vertical",
  "style:print-content",
  "style:repeat-content",
  "style:rotation-align",
  "style:rotation-angle",
  "style:shadow",
  "style:shrink-to-fit",
  "style:text-align-source",
  "style:vertical-align",
  "style:writing-mode",
] as const;
const SECTION_STYLE_ATTRIBUTES = [
  "fo:background-color",
  "fo:margin-left",
  "fo:margin-right",
  "style:editable",
  "style:protect",
  "style:writing-mode",
  "text:dont-balance-text-columns",
] as const;
const PAGE_LAYOUT_STYLE_ATTRIBUTES = [
  "fo:background-color",
  "fo:border",
  "fo:border-bottom",
  "fo:border-left",
  "fo:border-right",
  "fo:border-top",
  "fo:margin",
  "fo:margin-bottom",
  "fo:margin-left",
  "fo:margin-right",
  "fo:margin-top",
  "fo:padding",
  "fo:padding-bottom",
  "fo:padding-left",
  "fo:padding-right",
  "fo:padding-top",
  "fo:page-height",
  "fo:page-width",
  "style:border-line-width",
  "style:border-line-width-bottom",
  "style:border-line-width-left",
  "style:border-line-width-right",
  "style:border-line-width-top",
  "style:first-page-number",
  "style:footnote-max-height",
  "style:layout-grid-base-height",
  "style:layout-grid-base-width",
  "style:layout-grid-color",
  "style:layout-grid-display",
  "style:layout-grid-lines",
  "style:layout-grid-mode",
  "style:layout-grid-print",
  "style:layout-grid-ruby-below",
  "style:layout-grid-ruby-height",
  "style:layout-grid-snap-to",
  "style:layout-grid-standard-mode",
  "style:num-format",
  "style:num-letter-sync",
  "style:num-prefix",
  "style:num-suffix",
  "style:paper-tray-name",
  "style:print",
  "style:print-orientation",
  "style:print-page-order",
  "style:register-truth-ref-style-name",
  "style:scale-to",
  "style:scale-to-X",
  "style:scale-to-Y",
  "style:scale-to-pages",
  "style:shadow",
  "style:table-centering",
  "style:writing-mode",
] as const;
const HEADER_FOOTER_STYLE_ATTRIBUTES = [
  "fo:background-color",
  "fo:border",
  "fo:border-bottom",
  "fo:border-left",
  "fo:border-right",
  "fo:border-top",
  "fo:margin",
  "fo:margin-bottom",
  "fo:margin-left",
  "fo:margin-right",
  "fo:margin-top",
  "fo:min-height",
  "fo:padding",
  "fo:padding-bottom",
  "fo:padding-left",
  "fo:padding-right",
  "fo:padding-top",
  "style:border-line-width",
  "style:border-line-width-bottom",
  "style:border-line-width-left",
  "style:border-line-width-right",
  "style:border-line-width-top",
  "style:dynamic-spacing",
  "style:shadow",
  "svg:height",
] as const;
const DRAWING_PAGE_STYLE_ATTRIBUTES = [
  "draw:background-size",
  "draw:fill",
  "draw:fill-color",
  "draw:fill-gradient-name",
  "draw:fill-hatch-name",
  "draw:fill-hatch-solid",
  "draw:fill-image-height",
  "draw:fill-image-name",
  "draw:fill-image-ref-point",
  "draw:fill-image-ref-point-x",
  "draw:fill-image-ref-point-y",
  "draw:fill-image-width",
  "draw:gradient-step-count",
  "draw:opacity",
  "draw:opacity-name",
  "draw:secondary-fill-color",
  "draw:tile-repeat-offset",
  "presentation:background-objects-visible",
  "presentation:background-visible",
  "presentation:display-date-time",
  "presentation:display-footer",
  "presentation:display-header",
  "presentation:display-page-number",
  "presentation:duration",
  "presentation:transition-speed",
  "presentation:transition-style",
  "presentation:transition-type",
  "presentation:visibility",
  "smil:direction",
  "smil:fadeColor",
  "smil:subtype",
  "smil:type",
  "style:repeat",
  "svg:fill-rule",
] as const;
const RUBY_STYLE_ATTRIBUTES = ["style:ruby-align", "style:ruby-position"] as const;

const BOOLEAN_STYLE_ATTRIBUTES = [
  "dr3d:close-back",
  "dr3d:close-front",
  "dr3d:enabled",
  "dr3d:specular",
  "draw:auto-grow-height",
  "draw:auto-grow-width",
  "draw:caption-fit-line-length",
  "draw:color-inversion",
  "draw:fill-hatch-solid",
  "draw:fit-to-contour",
  "draw:fit-to-size",
  "draw:frame-display-border",
  "draw:frame-display-scrollbar",
  "draw:marker-end-center",
  "draw:marker-start-center",
  "draw:parallel",
  "draw:show-unit",
  "fo:hyphenate",
  "presentation:background-objects-visible",
  "presentation:background-visible",
  "presentation:display-date-time",
  "presentation:display-footer",
  "presentation:display-header",
  "presentation:display-page-number",
  "style:auto-text-indent",
  "style:contextual-spacing",
  "style:display",
  "style:dynamic-spacing",
  "style:editable",
  "style:flow-with-text",
  "style:font-independent-line-spacing",
  "style:join-border",
  "style:justify-single-word",
  "style:layout-grid-display",
  "style:layout-grid-print",
  "style:layout-grid-ruby-below",
  "style:layout-grid-snap-to",
  "style:layout-grid-standard-mode",
  "style:letter-kerning",
  "style:may-break-between-rows",
  "style:num-letter-sync",
  "style:print-content",
  "style:protect",
  "style:register-true",
  "style:repeat-content",
  "style:shrink-to-fit",
  "style:snap-to-layout-grid",
  "style:text-blinking",
  "style:text-outline",
  "style:use-optimal-column-width",
  "style:use-optimal-row-height",
  "style:use-window-font-color",
  "style:wrap-contour",
  "style:writing-mode-automatic",
  "table:display",
  "text:animation-start-inside",
  "text:animation-stop-inside",
  "text:dont-balance-text-columns",
  "text:line-break",
  "text:number-lines",
] as const;

const INTEGER_STYLE_ATTRIBUTES = [
  "draw:decimal-places",
  "draw:gradient-step-count",
  "draw:ole-draw-aspect",
  "dr3d:horizontal-segments",
  "dr3d:vertical-segments",
  "fo:hyphenation-ladder-count",
  "fo:hyphenation-push-char-count",
  "fo:hyphenation-remain-char-count",
  "fo:orphans",
  "fo:widows",
  "style:decimal-places",
  "style:first-page-number",
  "style:layout-grid-lines",
  "style:number-wrapped-paragraphs",
  "style:page-number",
  "style:scale-to-X",
  "style:scale-to-Y",
  "style:scale-to-pages",
  "text:anchor-page-number",
  "text:line-number",
  "svg:stroke-opacity",
] as const;

type ParagraphStyleAttributeName = (typeof PARAGRAPH_STYLE_ATTRIBUTES)[number];
type TextStyleAttributeName = (typeof TEXT_STYLE_ATTRIBUTES)[number];
type GraphicStyleAttributeName = (typeof GRAPHIC_STYLE_ATTRIBUTES)[number];
type TableStyleAttributeName = (typeof TABLE_STYLE_ATTRIBUTES)[number];
type TableColumnStyleAttributeName = (typeof TABLE_COLUMN_STYLE_ATTRIBUTES)[number];
type TableRowStyleAttributeName = (typeof TABLE_ROW_STYLE_ATTRIBUTES)[number];
type TableCellStyleAttributeName = (typeof TABLE_CELL_STYLE_ATTRIBUTES)[number];
type SectionStyleAttributeName = (typeof SECTION_STYLE_ATTRIBUTES)[number];
type PageLayoutStyleAttributeName = (typeof PAGE_LAYOUT_STYLE_ATTRIBUTES)[number];
type HeaderFooterStyleAttributeName = (typeof HEADER_FOOTER_STYLE_ATTRIBUTES)[number];
type DrawingPageStyleAttributeName = (typeof DRAWING_PAGE_STYLE_ATTRIBUTES)[number];
type RubyStyleAttributeName = (typeof RUBY_STYLE_ATTRIBUTES)[number];

export interface ParagraphStylePropertiesOptions {
  name: "style:paragraph-properties";
  attributes: Partial<Record<ParagraphStyleAttributeName, StylePropertyValue>>;
}
export interface TextStylePropertiesOptions {
  name: "style:text-properties";
  attributes: Partial<Record<TextStyleAttributeName, StylePropertyValue>>;
}
export interface GraphicStylePropertiesOptions {
  name: "style:graphic-properties" | "loext:graphic-properties";
  attributes: Partial<Record<GraphicStyleAttributeName, StylePropertyValue>>;
}
export interface TableStylePropertiesOptions {
  name: "style:table-properties";
  attributes: Partial<Record<TableStyleAttributeName, StylePropertyValue>>;
}
export interface TableColumnStylePropertiesOptions {
  name: "style:table-column-properties";
  attributes: Partial<Record<TableColumnStyleAttributeName, StylePropertyValue>>;
}
export interface TableRowStylePropertiesOptions {
  name: "style:table-row-properties";
  attributes: Partial<Record<TableRowStyleAttributeName, StylePropertyValue>>;
}
export interface TableCellStylePropertiesOptions {
  name: "style:table-cell-properties";
  attributes: Partial<Record<TableCellStyleAttributeName, StylePropertyValue>>;
}
export interface SectionStylePropertiesOptions {
  name: "style:section-properties";
  attributes: Partial<Record<SectionStyleAttributeName, StylePropertyValue>>;
}
export interface PageLayoutStylePropertiesOptions {
  name: "style:page-layout-properties";
  attributes: Partial<Record<PageLayoutStyleAttributeName, StylePropertyValue>>;
}
export interface HeaderFooterStylePropertiesOptions {
  name: "style:header-footer-properties";
  attributes: Partial<Record<HeaderFooterStyleAttributeName, StylePropertyValue>>;
}
export interface DrawingPageStylePropertiesOptions {
  name: "style:drawing-page-properties";
  attributes: Partial<Record<DrawingPageStyleAttributeName, StylePropertyValue>>;
}
export interface RubyStylePropertiesOptions {
  name: "style:ruby-properties";
  attributes: Partial<Record<RubyStyleAttributeName, StylePropertyValue>>;
}

export type StylePropertyOptions =
  | ParagraphStylePropertiesOptions
  | TextStylePropertiesOptions
  | GraphicStylePropertiesOptions
  | TableStylePropertiesOptions
  | TableColumnStylePropertiesOptions
  | TableRowStylePropertiesOptions
  | TableCellStylePropertiesOptions
  | SectionStylePropertiesOptions
  | PageLayoutStylePropertiesOptions
  | HeaderFooterStylePropertiesOptions
  | DrawingPageStylePropertiesOptions
  | RubyStylePropertiesOptions;

/** Automatic style properties retained as a typed projection, never as raw XML. */
export interface StyleResourceOptions {
  name: string;
  family: StyleFamily;
  properties: StylePropertyOptions[];
}

/** ODF flavor used to normalize ODF length attributes without losing scale. */
export type StylePropertyScale = "hundredth-millimeter" | "spreadsheet" | "emu";

const STYLE_PROPERTY_ATTRIBUTES = {
  "style:paragraph-properties": PARAGRAPH_STYLE_ATTRIBUTES,
  "style:text-properties": TEXT_STYLE_ATTRIBUTES,
  "style:graphic-properties": GRAPHIC_STYLE_ATTRIBUTES,
  "loext:graphic-properties": GRAPHIC_STYLE_ATTRIBUTES,
  "style:table-properties": TABLE_STYLE_ATTRIBUTES,
  "style:table-column-properties": TABLE_COLUMN_STYLE_ATTRIBUTES,
  "style:table-row-properties": TABLE_ROW_STYLE_ATTRIBUTES,
  "style:table-cell-properties": TABLE_CELL_STYLE_ATTRIBUTES,
  "style:section-properties": SECTION_STYLE_ATTRIBUTES,
  "style:page-layout-properties": PAGE_LAYOUT_STYLE_ATTRIBUTES,
  "style:header-footer-properties": HEADER_FOOTER_STYLE_ATTRIBUTES,
  "style:drawing-page-properties": DRAWING_PAGE_STYLE_ATTRIBUTES,
  "style:ruby-properties": RUBY_STYLE_ATTRIBUTES,
} as const;

const STYLE_FAMILY_PROPERTIES: Record<
  StyleFamily,
  readonly (keyof typeof STYLE_PROPERTY_ATTRIBUTES)[]
> = {
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
    "style:graphic-properties",
    "loext:graphic-properties",
    "style:paragraph-properties",
    "style:text-properties",
  ],
  section: ["style:section-properties"],
  "page-layout": ["style:page-layout-properties", "style:header-footer-properties"],
  "drawing-page": ["style:drawing-page-properties"],
  presentation: [
    "style:graphic-properties",
    "loext:graphic-properties",
    "style:paragraph-properties",
    "style:text-properties",
  ],
  ruby: ["style:ruby-properties"],
};

export interface StylePropertyParseContext {
  family: StyleFamily;
  scale: StylePropertyScale;
}

export type StylePropertyParseResult =
  | { ok: true; value: StylePropertyOptions }
  | {
      ok: false;
      attributeName?: string;
      elementName: string;
      reason: string;
    };

export function parseStyleProperty(
  element: Element,
  context: StylePropertyParseContext,
): StylePropertyParseResult {
  const elementName = element.name ?? "";
  const allowedElements = STYLE_FAMILY_PROPERTIES[context.family];
  if (!allowedElements.includes(elementName as keyof typeof STYLE_PROPERTY_ATTRIBUTES))
    return { ok: false, elementName, reason: "unknown style property element" };
  const allowedAttributes =
    STYLE_PROPERTY_ATTRIBUTES[elementName as keyof typeof STYLE_PROPERTY_ATTRIBUTES];
  const attributes: Record<string, StylePropertyValue> = {};
  for (const [attributeName, source] of Object.entries(element.attributes ?? {})) {
    if (!allowedAttributes.includes(attributeName as never))
      return { ok: false, attributeName, elementName, reason: "attribute has no ODF RNG mapping" };
    const value = parseStyleAttributeValue(attributeName, String(source), context.scale);
    if (value === undefined)
      return { ok: false, attributeName, elementName, reason: "invalid style property value" };
    attributes[attributeName] = value;
  }
  return {
    ok: true,
    value: { name: elementName, attributes } as StylePropertyOptions,
  };
}

export function stylePropertyAttributes(
  property: StylePropertyOptions,
  scale: StylePropertyScale,
): XmlAttributes {
  return Object.fromEntries(
    Object.entries(property.attributes).map(([name, value]) => [
      name,
      styleAttributeValueXml(name, value, scale),
    ]),
  );
}

export function removeStylePropertyAttributes(
  property: StylePropertyOptions,
  shouldRemove: (attributeName: string) => boolean,
): StylePropertyOptions {
  const attributes = Object.fromEntries(
    Object.entries(property.attributes).filter(([name]) => !shouldRemove(name)),
  );
  return { ...property, attributes } as StylePropertyOptions;
}

function parseStyleAttributeValue(
  name: string,
  source: string,
  scale: StylePropertyScale,
): StylePropertyValue | undefined {
  if (isBooleanStyleAttribute(name)) {
    if (source !== "true" && source !== "false") return undefined;
    return source === "true";
  }
  if (isPercentStyleAttribute(name)) {
    if (!source.endsWith("%")) return source;
    const value = Number.parseFloat(source);
    return Number.isFinite(value) ? value : undefined;
  }
  if (isLengthStyleAttribute(name)) {
    const emu = lengthToEmu(source);
    if (emu === undefined) return source;
    if (scale === "emu") return emu;
    if (name === "style:row-height" || name === "style:min-row-height") return emu / 12700;
    return scale === "spreadsheet" ? Math.round(emu / 9525) : Math.round(emu / 635);
  }
  if (isAngleStyleAttribute(name)) {
    const value = Number(source);
    return Number.isFinite(value) ? value : undefined;
  }
  if ((INTEGER_STYLE_ATTRIBUTES as readonly string[]).includes(name)) {
    const value = Number(source);
    return Number.isInteger(value) ? value : undefined;
  }
  return source;
}

function styleAttributeValueXml(
  name: string,
  value: StylePropertyValue,
  scale: StylePropertyScale,
): StylePropertyValue {
  if (typeof value !== "number") return value;
  if (isPercentStyleAttribute(name)) return `${value}%`;
  if (isAngleStyleAttribute(name) || (INTEGER_STYLE_ATTRIBUTES as readonly string[]).includes(name))
    return value;
  if (!isLengthStyleAttribute(name)) return value;
  if (scale === "emu") return emuToLength(value);
  if (name === "style:row-height" || name === "style:min-row-height") return `${value}pt`;
  if (scale === "spreadsheet") return emuToLength(value * 9525);
  return emuToLength(value * 635);
}

function isBooleanStyleAttribute(name: string): boolean {
  return (BOOLEAN_STYLE_ATTRIBUTES as readonly string[]).includes(name);
}

function isPercentStyleAttribute(name: string): boolean {
  return /(?:line-height|rel-width|rel-height|opacity|transparency|scale)$/.test(name);
}

function isAngleStyleAttribute(name: string): boolean {
  return /(?:angle|rotation)$/.test(name);
}

function isLengthStyleAttribute(name: string): boolean {
  return /(?:^|[-:])(?:width|height|margin|spacing|distance|indent|padding|position|x|y|depth|radius)(?:$|[-0-9])/.test(
    name,
  );
}

/** Unified OCF and style resources attached to one canonical family model. */
export interface PackageOptions {
  /** Source package MIME type; retained for template and legacy packages. */
  packageMediaType?: string;
  packageManifest?: OcfManifestOptions;
  packageMembers?: PackageMemberOptions[];
  styleResources?: StyleResourceOptions[];
}
