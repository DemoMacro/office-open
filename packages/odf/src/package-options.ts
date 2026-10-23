import type { DataType } from "@office-open/core";
import { encodeBase64, toUint8Array } from "@office-open/core";
import type { Element } from "@office-open/xml";

import type { OcfManifestOptions } from "./container/manifest";
import {
  attributeString,
  emuToLength,
  escapeText,
  lengthToEmu,
  unknownAttributeNames,
  xmlElement,
  type XmlAttributes,
} from "./runtime/xml";

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

export type ListLevelPropertiesParseResult =
  | { ok: true; value: ListLevelStylePropertiesOptions }
  | {
      ok: false;
      attributeName?: string;
      elementName: string;
      reason: string;
    };

export type ChartPropertiesParseResult =
  | { ok: true; value: ChartStylePropertiesOptions }
  | {
      ok: false;
      attributeName?: string;
      elementName: string;
      reason: string;
    };

export interface ChartStylePropertiesOptions {
  name: "style:chart-properties";
  attributes: Partial<Record<(typeof CHART_PROPERTIES_ATTRIBUTES)[number], StylePropertyValue>>;
  symbolImage?: ChartSymbolImageOptions;
}

export interface ChartSymbolImageOptions {
  attributes: Partial<Record<(typeof SYMBOL_IMAGE_ATTRIBUTES)[number], string>>;
}

export interface ListLevelStylePropertiesOptions {
  name: "style:list-level-properties";
  attributes: Partial<Record<(typeof LIST_LEVEL_ATTRIBUTES)[number], StylePropertyValue>>;
  labelAlignment?: ListLevelLabelAlignmentOptions;
}

export interface ListLevelLabelAlignmentOptions {
  attributes: Partial<Record<(typeof LABEL_ALIGNMENT_ATTRIBUTES)[number], StylePropertyValue>>;
}

export interface StyleBackgroundImageOptions {
  name: "style:background-image";
  attributes: Partial<Record<(typeof BACKGROUND_IMAGE_ATTRIBUTES)[number], StylePropertyValue>>;
  data?: DataType;
}

export interface StyleTabStopOptions {
  attributes: Partial<Record<(typeof TAB_STOP_ATTRIBUTES)[number], StylePropertyValue>>;
}

export interface StyleTabStopsOptions {
  name: "style:tab-stops";
  stops: StyleTabStopOptions[];
}

export interface StyleMapOptions {
  name: "style:map";
  condition: string;
  applyStyleName: string;
  baseCellAddress?: string;
}

export interface StyleDropCapOptions {
  name: "style:drop-cap";
  attributes: Partial<Record<(typeof DROP_CAP_ATTRIBUTES)[number], StylePropertyValue>>;
}

export interface StyleColumnOptions {
  attributes: Partial<Record<(typeof COLUMN_ATTRIBUTES)[number], StylePropertyValue>>;
}

export interface StyleColumnSeparatorOptions {
  attributes: Partial<Record<(typeof COLUMN_SEPARATOR_ATTRIBUTES)[number], StylePropertyValue>>;
}

export interface StyleColumnsOptions {
  name: "style:columns";
  attributes: Partial<Record<"fo:column-count" | "fo:column-gap", StylePropertyValue>>;
  columns: StyleColumnOptions[];
  separator?: StyleColumnSeparatorOptions;
}

export interface StyleListLevelOptions {
  name:
    | "text:list-level-style-number"
    | "text:list-level-style-bullet"
    | "text:list-level-style-image";
  attributes: Partial<
    Record<
      | (typeof NUMBER_LEVEL_ATTRIBUTES)[number]
      | (typeof BULLET_LEVEL_ATTRIBUTES)[number]
      | (typeof IMAGE_LEVEL_ATTRIBUTES)[number],
      StylePropertyValue
    >
  >;
  listProperties?: ListLevelStylePropertiesOptions;
  textProperties?: TextStylePropertiesOptions;
}

export interface StyleListStyleOptions {
  name: "text:list-style";
  attributes: Partial<Record<(typeof LIST_STYLE_ATTRIBUTES)[number], StylePropertyValue>>;
  levels: StyleListLevelOptions[];
}

export interface StyleFootnoteSeparatorOptions {
  name: "style:footnote-sep";
  attributes: Partial<Record<(typeof FOOTNOTE_SEPARATOR_ATTRIBUTES)[number], StylePropertyValue>>;
}

export interface StyleSoundOptions {
  name: "presentation:sound";
  attributes: Partial<Record<(typeof SOUND_ATTRIBUTES)[number], StylePropertyValue>>;
}

export interface StyleNotesConfigurationOptions {
  name: "text:notes-configuration";
  attributes: Partial<Record<(typeof NOTES_CONFIGURATION_ATTRIBUTES)[number], StylePropertyValue>>;
  continuationForward?: string;
  continuationBackward?: string;
}

export type StylePropertyChildOptions =
  | StyleBackgroundImageOptions
  | StyleColumnsOptions
  | StyleDropCapOptions
  | StyleFootnoteSeparatorOptions
  | StyleListStyleOptions
  | StyleNotesConfigurationOptions
  | StyleSoundOptions
  | StyleTabStopsOptions;

const TAB_STOP_ATTRIBUTES = [
  "style:char",
  "style:leader-color",
  "style:leader-style",
  "style:leader-text",
  "style:leader-text-style",
  "style:leader-type",
  "style:leader-width",
  "style:position",
  "style:type",
] as const;

const DROP_CAP_ATTRIBUTES = [
  "style:distance",
  "style:length",
  "style:lines",
  "style:style-name",
] as const;

const BACKGROUND_IMAGE_ATTRIBUTES = [
  "draw:opacity",
  "style:filter-name",
  "style:position",
  "style:repeat",
  "xlink:actuate",
  "xlink:href",
  "xlink:show",
  "xlink:type",
] as const;

const COLUMN_ATTRIBUTES = [
  "fo:end-indent",
  "fo:space-after",
  "fo:space-before",
  "fo:start-indent",
  "style:rel-width",
] as const;

const COLUMN_SEPARATOR_ATTRIBUTES = [
  "style:color",
  "style:height",
  "style:style",
  "style:vertical-align",
  "style:width",
] as const;

const LIST_STYLE_ATTRIBUTES = [
  "style:display-name",
  "style:name",
  "text:consecutive-numbering",
] as const;

const NUMBER_LEVEL_ATTRIBUTES = [
  "style:num-format",
  "style:num-letter-sync",
  "style:num-prefix",
  "style:num-suffix",
  "text:display-levels",
  "text:level",
  "text:start-value",
  "text:style-name",
] as const;

const BULLET_LEVEL_ATTRIBUTES = [
  "style:num-prefix",
  "style:num-suffix",
  "text:bullet-char",
  "text:bullet-relative-size",
  "text:level",
  "text:style-name",
] as const;

const IMAGE_LEVEL_ATTRIBUTES = [
  "text:level",
  "xlink:actuate",
  "xlink:href",
  "xlink:show",
  "xlink:type",
] as const;

const NOTES_CONFIGURATION_ATTRIBUTES = [
  "style:num-format",
  "style:num-letter-sync",
  "style:num-prefix",
  "style:num-suffix",
  "text:citation-body-style-name",
  "text:citation-style-name",
  "text:default-style-name",
  "text:footnotes-position",
  "text:master-page-name",
  "text:note-class",
  "text:start-numbering-at",
  "text:start-value",
] as const;

const SYMBOL_IMAGE_ATTRIBUTES = ["xlink:href"] as const;

const LABEL_ALIGNMENT_ATTRIBUTES = [
  "fo:margin-left",
  "fo:text-indent",
  "text:label-followed-by",
  "text:list-tab-stop-position",
] as const;

const LIST_LEVEL_ATTRIBUTES = [
  "fo:height",
  "fo:text-align",
  "fo:width",
  "style:font-name",
  "style:vertical-pos",
  "style:vertical-rel",
  "svg:y",
  "text:list-level-position-and-space-mode",
  "text:min-label-distance",
  "text:min-label-width",
  "text:space-before",
] as const;

const SOUND_ATTRIBUTES = [
  "presentation:play-full",
  "xlink:actuate",
  "xlink:href",
  "xlink:show",
  "xlink:type",
  "xml:id",
] as const;

const FOOTNOTE_SEPARATOR_ATTRIBUTES = [
  "style:adjustment",
  "style:color",
  "style:distance-after-sep",
  "style:distance-before-sep",
  "style:line-style",
  "style:rel-width",
  "style:width",
] as const;

const CHART_PROPERTIES_ATTRIBUTES = [
  "chart:angle-offset",
  "chart:auto-position",
  "chart:auto-size",
  "chart:axis-label-position",
  "chart:axis-position",
  "chart:connect-bars",
  "chart:data-label-number",
  "chart:data-label-series",
  "chart:data-label-symbol",
  "chart:data-label-text",
  "chart:deep",
  "chart:display-label",
  "chart:error-category",
  "chart:error-lower-indicator",
  "chart:error-lower-limit",
  "chart:error-lower-range",
  "chart:error-margin",
  "chart:error-percentage",
  "chart:error-upper-indicator",
  "chart:error-upper-limit",
  "chart:error-upper-range",
  "chart:gap-width",
  "chart:group-bars-per-axis",
  "chart:hole-size",
  "chart:include-hidden-cells",
  "chart:interpolation",
  "chart:lines",
  "chart:interval-major",
  "chart:interval-minor-divisor",
  "chart:japanese-candle-stick",
  "chart:label-arrangement",
  "chart:label-position",
  "chart:label-position-negative",
  "chart:link-data-style-to-source",
  "chart:logarithmic",
  "chart:maximum",
  "chart:mean-value",
  "chart:minimum",
  "chart:origin",
  "chart:overlap",
  "chart:percentage",
  "chart:pie-offset",
  "chart:regression-force-intercept",
  "chart:regression-intercept-value",
  "chart:regression-max-degree",
  "chart:regression-moving-type",
  "chart:regression-name",
  "chart:regression-period",
  "chart:regression-type",
  "chart:reverse-direction",
  "chart:right-angled-axes",
  "chart:scale-text",
  "chart:series-source",
  "chart:solid-type",
  "chart:sort-by-x-values",
  "chart:spline-order",
  "chart:spline-resolution",
  "chart:stacked",
  "chart:symbol-height",
  "chart:symbol-name",
  "chart:symbol-type",
  "chart:symbol-width",
  "chart:text-overlap",
  "chart:three-dimensional",
  "chart:tick-mark-position",
  "chart:tick-marks-major-inner",
  "chart:tick-marks-major-outer",
  "chart:tick-marks-minor-inner",
  "chart:tick-marks-minor-outer",
  "chart:treat-empty-cells",
  "chart:vertical",
  "chart:visible",
  "style:direction",
  "style:rotation-angle",
  "text:line-break",
] as const;

const CHILD_LENGTH_ATTRIBUTES = new Set([
  "fo:column-gap",
  "fo:end-indent",
  "fo:height",
  "fo:margin-left",
  "fo:space-after",
  "fo:space-before",
  "fo:start-indent",
  "fo:text-indent",
  "fo:width",
  "style:distance",
  "style:distance-after-sep",
  "style:distance-before-sep",
  "style:leader-width",
  "style:position",
  "style:width",
  "chart:symbol-height",
  "chart:symbol-width",
  "svg:y",
  "text:list-tab-stop-position",
  "text:min-label-distance",
  "text:min-label-width",
  "text:space-before",
]);

const CHILD_PERCENT_ATTRIBUTES = new Set([
  "chart:hole-size",
  "draw:opacity",
  "style:height",
  "style:rel-width",
  "text:bullet-relative-size",
]);

const CHILD_BOOLEAN_ATTRIBUTES = new Set([
  "chart:auto-position",
  "chart:auto-size",
  "chart:connect-bars",
  "chart:data-label-series",
  "chart:data-label-symbol",
  "chart:data-label-text",
  "chart:deep",
  "chart:display-label",
  "chart:error-lower-indicator",
  "chart:error-upper-indicator",
  "chart:group-bars-per-axis",
  "chart:include-hidden-cells",
  "chart:lines",
  "chart:japanese-candle-stick",
  "chart:link-data-style-to-source",
  "chart:logarithmic",
  "chart:mean-value",
  "chart:percentage",
  "chart:regression-force-intercept",
  "chart:reverse-direction",
  "chart:right-angled-axes",
  "chart:scale-text",
  "chart:sort-by-x-values",
  "chart:stacked",
  "chart:text-overlap",
  "chart:three-dimensional",
  "chart:tick-marks-major-inner",
  "chart:tick-marks-major-outer",
  "chart:tick-marks-minor-inner",
  "chart:tick-marks-minor-outer",
  "chart:vertical",
  "chart:visible",
  "presentation:play-full",
  "style:num-letter-sync",
  "text:consecutive-numbering",
  "text:line-break",
]);

const CHILD_INTEGER_ATTRIBUTES = new Set([
  "chart:angle-offset",
  "chart:gap-width",
  "chart:interval-minor-divisor",
  "chart:overlap",
  "chart:pie-offset",
  "chart:regression-max-degree",
  "chart:regression-period",
  "chart:spline-order",
  "chart:spline-resolution",
  "fo:column-count",
  "style:lines",
  "style:rotation-angle",
  "text:display-levels",
  "text:start-value",
]);

const CHILD_DOUBLE_ATTRIBUTES = new Set([
  "chart:axis-position",
  "chart:error-lower-limit",
  "chart:error-margin",
  "chart:error-percentage",
  "chart:error-upper-limit",
  "chart:interval-major",
  "chart:maximum",
  "chart:minimum",
  "chart:origin",
  "chart:regression-intercept-value",
]);

export interface ParagraphStylePropertiesOptions {
  name: "style:paragraph-properties";
  attributes: Partial<Record<ParagraphStyleAttributeName, StylePropertyValue>>;
  children?: StylePropertyChildOptions[];
}
export interface TextStylePropertiesOptions {
  name: "style:text-properties";
  attributes: Partial<Record<TextStyleAttributeName, StylePropertyValue>>;
  children?: StylePropertyChildOptions[];
}
export interface GraphicStylePropertiesOptions {
  name: "style:graphic-properties" | "loext:graphic-properties";
  attributes: Partial<Record<GraphicStyleAttributeName, StylePropertyValue>>;
  children?: StylePropertyChildOptions[];
}
export interface TableStylePropertiesOptions {
  name: "style:table-properties";
  attributes: Partial<Record<TableStyleAttributeName, StylePropertyValue>>;
  children?: StylePropertyChildOptions[];
}
export interface TableColumnStylePropertiesOptions {
  name: "style:table-column-properties";
  attributes: Partial<Record<TableColumnStyleAttributeName, StylePropertyValue>>;
  children?: StylePropertyChildOptions[];
}
export interface TableRowStylePropertiesOptions {
  name: "style:table-row-properties";
  attributes: Partial<Record<TableRowStyleAttributeName, StylePropertyValue>>;
  children?: StylePropertyChildOptions[];
}
export interface TableCellStylePropertiesOptions {
  name: "style:table-cell-properties";
  attributes: Partial<Record<TableCellStyleAttributeName, StylePropertyValue>>;
  children?: StylePropertyChildOptions[];
}
export interface SectionStylePropertiesOptions {
  name: "style:section-properties";
  attributes: Partial<Record<SectionStyleAttributeName, StylePropertyValue>>;
  children?: StylePropertyChildOptions[];
}
export interface PageLayoutStylePropertiesOptions {
  name: "style:page-layout-properties";
  attributes: Partial<Record<PageLayoutStyleAttributeName, StylePropertyValue>>;
  children?: StylePropertyChildOptions[];
}
export interface HeaderFooterStylePropertiesOptions {
  name: "style:header-footer-properties";
  attributes: Partial<Record<HeaderFooterStyleAttributeName, StylePropertyValue>>;
  children?: StylePropertyChildOptions[];
}
export interface DrawingPageStylePropertiesOptions {
  name: "style:drawing-page-properties";
  attributes: Partial<Record<DrawingPageStyleAttributeName, StylePropertyValue>>;
  children?: StylePropertyChildOptions[];
}
export interface RubyStylePropertiesOptions {
  name: "style:ruby-properties";
  attributes: Partial<Record<RubyStyleAttributeName, StylePropertyValue>>;
  children?: StylePropertyChildOptions[];
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
  maps?: StyleMapOptions[];
}

const STYLE_MAP_ATTRIBUTES = [
  "style:condition",
  "style:apply-style-name",
  "style:base-cell-address",
] as const;

export function parseStyleMaps(style: Element): StyleMapOptions[] {
  return (style.elements ?? [])
    .filter((child): child is Element => child.type === "element" && child.name === "style:map")
    .map((element) => {
      const condition = attributeString(element, "style:condition");
      const applyStyleName = attributeString(element, "style:apply-style-name");
      if (condition === undefined || applyStyleName === undefined) {
        throw new Error("style:map requires condition and apply-style-name");
      }
      const unknown = unknownAttributeNames(element, STYLE_MAP_ATTRIBUTES);
      if (unknown.length > 0) throw new Error(`Unknown style:map attribute: ${unknown[0]}`);
      return {
        name: "style:map",
        condition,
        applyStyleName,
        ...(attributeString(element, "style:base-cell-address") === undefined
          ? {}
          : { baseCellAddress: attributeString(element, "style:base-cell-address") }),
      };
    });
}

export function styleMapsXml(maps: StyleMapOptions[] | undefined): string[] {
  return (maps ?? []).map((map) =>
    xmlElement("style:map", {
      "style:condition": map.condition,
      "style:apply-style-name": map.applyStyleName,
      "style:base-cell-address": map.baseCellAddress,
    }),
  );
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

const STYLE_PROPERTY_CHILDREN: Partial<
  Record<keyof typeof STYLE_PROPERTY_ATTRIBUTES, readonly string[]>
> = {
  "style:paragraph-properties": ["style:background-image", "style:drop-cap", "style:tab-stops"],
  "style:text-properties": [],
  "style:graphic-properties": ["style:background-image", "style:columns", "text:list-style"],
  "loext:graphic-properties": ["style:background-image", "style:columns", "text:list-style"],
  "style:table-properties": ["style:background-image"],
  "style:table-column-properties": [],
  "style:table-row-properties": ["style:background-image"],
  "style:table-cell-properties": ["style:background-image"],
  "style:section-properties": [
    "style:background-image",
    "style:columns",
    "text:notes-configuration",
  ],
  "style:page-layout-properties": ["style:background-image", "style:columns", "style:footnote-sep"],
  "style:header-footer-properties": ["style:background-image"],
  "style:drawing-page-properties": ["presentation:sound"],
  "style:ruby-properties": [],
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

export type StylePropertyChildParseResult =
  | { ok: true; value: StylePropertyChildOptions }
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
  if (
    !allowedElements.includes(elementName as keyof typeof STYLE_PROPERTY_ATTRIBUTES) ||
    !isStylePropertyElement(elementName)
  )
    return { ok: false, elementName, reason: "unknown style property element" };
  const allowedAttributes = STYLE_PROPERTY_ATTRIBUTES[elementName];
  const allowedChildren = STYLE_PROPERTY_CHILDREN[elementName] ?? [];
  const attributes: Record<string, StylePropertyValue> = {};
  for (const [attributeName, source] of Object.entries(element.attributes ?? {})) {
    if (!allowedAttributes.includes(attributeName as never))
      return { ok: false, attributeName, elementName, reason: "attribute has no ODF RNG mapping" };
    const value = parseStyleAttributeValue(attributeName, String(source), context.scale);
    if (value === undefined)
      return { ok: false, attributeName, elementName, reason: "invalid style property value" };
    attributes[attributeName] = value;
  }
  const children: StylePropertyChildOptions[] = [];
  for (const child of element.elements ?? []) {
    if (child.type !== "element") continue;
    const childName = child.name ?? "";
    if (!allowedChildren.includes(childName as never))
      return { ok: false, elementName: childName, reason: "child has no ODF RNG mapping" };
    const parsedChild = parseStyleChild(child, context);
    if (!parsedChild.ok) return parsedChild;
    children.push(parsedChild.value);
  }
  return {
    ok: true,
    value: {
      name: elementName,
      attributes,
      ...(children.length ? { children } : {}),
    } as StylePropertyOptions,
  };
}

function isStylePropertyElement(name: string): name is keyof typeof STYLE_PROPERTY_ATTRIBUTES {
  return Object.hasOwn(STYLE_PROPERTY_ATTRIBUTES, name);
}

export function parseListLevelProperties(
  element: Element,
  scale: StylePropertyScale,
): ListLevelPropertiesParseResult {
  const parsed = parseAttributes(element, LIST_LEVEL_ATTRIBUTES, scale);
  if (!parsed.ok) return parsed;
  let labelAlignment: ListLevelLabelAlignmentOptions | undefined;
  for (const child of element.elements ?? []) {
    if (child.type !== "element") continue;
    if (child.name !== "style:list-level-label-alignment")
      return {
        ok: false,
        elementName: child.name ?? "",
        reason: "child has no ODF RNG mapping",
      };
    if (labelAlignment)
      return {
        ok: false,
        elementName: child.name ?? "",
        reason: "duplicate style property child",
      };
    const parsedAlignment = parseAttributes(child, LABEL_ALIGNMENT_ATTRIBUTES, scale);
    if (!parsedAlignment.ok) return parsedAlignment;
    if (parsedAlignment.value["text:label-followed-by"] === undefined)
      return {
        ok: false,
        attributeName: "text:label-followed-by",
        elementName: child.name ?? "",
        reason: "required style property attribute is missing",
      };
    labelAlignment = {
      attributes: parsedAlignment.value as ListLevelLabelAlignmentOptions["attributes"],
    };
  }
  return {
    ok: true,
    value: {
      name: "style:list-level-properties",
      attributes: parsed.value as ListLevelStylePropertiesOptions["attributes"],
      ...(labelAlignment ? { labelAlignment } : {}),
    },
  };
}

export function parseChartStyleProperties(
  element: Element | undefined,
  scale: StylePropertyScale,
): ChartPropertiesParseResult {
  if (!element) return { ok: true, value: { name: "style:chart-properties", attributes: {} } };
  const parsed = parseAttributes(element, CHART_PROPERTIES_ATTRIBUTES, scale);
  if (!parsed.ok) return parsed;
  let symbolImage: ChartSymbolImageOptions | undefined;
  for (const child of element.elements ?? []) {
    if (child.type !== "element") continue;
    if (child.name !== "chart:symbol-image")
      return { ok: false, elementName: child.name ?? "", reason: "child has no ODF RNG mapping" };
    if (symbolImage)
      return {
        ok: false,
        elementName: child.name ?? "",
        reason: "duplicate style property child",
      };
    const parsedImage = parseAttributes(child, SYMBOL_IMAGE_ATTRIBUTES, scale);
    if (!parsedImage.ok) return parsedImage;
    symbolImage = { attributes: parsedImage.value as ChartSymbolImageOptions["attributes"] };
  }
  return {
    ok: true,
    value: {
      name: "style:chart-properties",
      attributes: parsed.value as ChartStylePropertiesOptions["attributes"],
      ...(symbolImage ? { symbolImage } : {}),
    },
  };
}

export function listLevelPropertiesXml(
  properties: ListLevelStylePropertiesOptions,
  scale: StylePropertyScale,
): string {
  return xmlElement(
    properties.name,
    childAttributesXml(properties.attributes, scale),
    properties.labelAlignment
      ? [
          xmlElement(
            "style:list-level-label-alignment",
            childAttributesXml(properties.labelAlignment.attributes, scale),
          ),
        ]
      : [],
  );
}

export function chartPropertiesXml(
  properties: ChartStylePropertiesOptions,
  scale: StylePropertyScale,
): string {
  return xmlElement(
    properties.name,
    childAttributesXml(properties.attributes, scale),
    properties.symbolImage
      ? [
          xmlElement(
            "chart:symbol-image",
            childAttributesXml(properties.symbolImage.attributes, scale),
          ),
        ]
      : [],
  );
}

export function stylePropertyXml(
  property: StylePropertyOptions,
  scale: StylePropertyScale,
): string {
  return xmlElement(
    property.name,
    stylePropertyAttributes(property, scale),
    (property.children ?? []).map((child) => stylePropertyChildXml(child, scale)),
  );
}

type AttributeParseResult<T extends readonly string[]> =
  | { ok: true; value: Partial<Record<T[number], StylePropertyValue>> }
  | {
      ok: false;
      attributeName?: string;
      elementName: string;
      reason: string;
    };

function parseAttributes<const T extends readonly string[]>(
  element: Element,
  allowed: T,
  scale: StylePropertyScale,
): AttributeParseResult<T> {
  const value: Partial<Record<T[number], StylePropertyValue>> = {};
  for (const [name, source] of Object.entries(element.attributes ?? {})) {
    const attributeName = allowed.find((candidate): candidate is T[number] => candidate === name);
    if (!attributeName)
      return {
        ok: false,
        attributeName: name,
        elementName: element.name ?? "",
        reason: "attribute has no ODF RNG mapping",
      };
    const parsed = parseStyleChildValue(name, String(source), scale);
    if (parsed === undefined)
      return {
        ok: false,
        attributeName: name,
        elementName: element.name ?? "",
        reason: "invalid style property value",
      };
    value[attributeName] = parsed;
  }
  return { ok: true, value };
}

function parseStyleChild(
  element: Element,
  context: StylePropertyParseContext,
): StylePropertyChildParseResult {
  const name = element.name ?? "";
  const scale = context.scale;
  if (name === "style:background-image") {
    const parsed = parseAttributes(element, BACKGROUND_IMAGE_ATTRIBUTES, scale);
    if (!parsed.ok) return parsed;
    const binary = element.elements?.find(
      (child) => child.type === "element" && child.name === "office:binary-data",
    );
    if (
      element.elements?.some(
        (child) => child.type === "element" && child.name !== "office:binary-data",
      )
    )
      return { ok: false, elementName: name, reason: "child has no ODF RNG mapping" };
    const encoded = binary?.elements
      ?.filter((child) => child.type === "text")
      .map((child) => String(child.text ?? ""))
      .join("");
    return {
      ok: true,
      value: {
        name,
        attributes: parsed.value as StyleBackgroundImageOptions["attributes"],
        ...(encoded === undefined ? {} : { data: toUint8Array(encoded, { encoding: "base64" }) }),
      },
    };
  }
  if (name === "style:tab-stops") {
    const stops: StyleTabStopOptions[] = [];
    for (const stopElement of element.elements ?? []) {
      if (stopElement.type !== "element" || stopElement.name !== "style:tab-stop")
        return {
          ok: false,
          elementName: stopElement.name ?? "",
          reason: "child has no ODF RNG mapping",
        };
      const parsed = parseAttributes(stopElement, TAB_STOP_ATTRIBUTES, scale);
      if (!parsed.ok) return parsed;
      stops.push({ attributes: parsed.value as StyleTabStopOptions["attributes"] });
    }
    return { ok: true, value: { name, stops } };
  }
  if (name === "style:drop-cap") {
    const parsed = parseAttributes(element, DROP_CAP_ATTRIBUTES, scale);
    return parsed.ok
      ? { ok: true, value: { name, attributes: parsed.value as StyleDropCapOptions["attributes"] } }
      : parsed;
  }
  if (name === "style:columns") {
    const parsed = parseAttributes(element, ["fo:column-count", "fo:column-gap"], scale);
    if (!parsed.ok) return parsed;
    const columns: StyleColumnOptions[] = [];
    let separator: StyleColumnSeparatorOptions | undefined;
    for (const child of element.elements ?? []) {
      if (child.type !== "element") continue;
      if (child.name === "style:column") {
        const parsedColumn = parseAttributes(child, COLUMN_ATTRIBUTES, scale);
        if (!parsedColumn.ok) return parsedColumn;
        columns.push({ attributes: parsedColumn.value as StyleColumnOptions["attributes"] });
        continue;
      }
      if (child.name === "style:column-sep") {
        const parsedSeparator = parseAttributes(child, COLUMN_SEPARATOR_ATTRIBUTES, scale);
        if (!parsedSeparator.ok) return parsedSeparator;
        separator = {
          attributes: parsedSeparator.value as StyleColumnSeparatorOptions["attributes"],
        };
        continue;
      }
      return { ok: false, elementName: child.name ?? "", reason: "child has no ODF RNG mapping" };
    }
    return {
      ok: true,
      value: {
        name,
        attributes: parsed.value as StyleColumnsOptions["attributes"],
        columns,
        ...(separator ? { separator } : {}),
      },
    };
  }
  if (name === "text:list-style") {
    const parsed = parseAttributes(element, LIST_STYLE_ATTRIBUTES, scale);
    if (!parsed.ok) return parsed;
    const levels: StyleListLevelOptions[] = [];
    for (const levelElement of element.elements ?? []) {
      if (levelElement.type !== "element") continue;
      const levelName = levelElement.name ?? "";
      const levelAttributes = listLevelAttributes(levelName);
      if (!levelAttributes)
        return { ok: false, elementName: levelName, reason: "child has no ODF RNG mapping" };
      const parsedLevel = parseAttributes(levelElement, levelAttributes, scale);
      if (!parsedLevel.ok) return parsedLevel;
      let listProperties: ListLevelStylePropertiesOptions | undefined;
      let textProperties: TextStylePropertiesOptions | undefined;
      for (const child of levelElement.elements ?? []) {
        if (child.type !== "element") continue;
        if (child.name === "style:list-level-properties") {
          const parsedProperties = parseListLevelProperties(child, scale);
          if (!parsedProperties.ok) return parsedProperties;
          listProperties = parsedProperties.value;
        } else if (child.name === "style:text-properties") {
          const parsedText = parseStyleProperty(child, { ...context, family: "text" });
          if (!parsedText.ok) return parsedText;
          textProperties = parsedText.value as TextStylePropertiesOptions;
        } else {
          return {
            ok: false,
            elementName: child.name ?? "",
            reason: "child has no ODF RNG mapping",
          };
        }
      }
      levels.push({
        name: levelName as StyleListLevelOptions["name"],
        attributes: parsedLevel.value as StyleListLevelOptions["attributes"],
        ...(listProperties ? { listProperties } : {}),
        ...(textProperties ? { textProperties } : {}),
      });
    }
    return {
      ok: true,
      value: {
        name,
        attributes: parsed.value as StyleListStyleOptions["attributes"],
        levels,
      },
    };
  }
  if (name === "style:footnote-sep") {
    const parsed = parseAttributes(element, FOOTNOTE_SEPARATOR_ATTRIBUTES, scale);
    return parsed.ok
      ? {
          ok: true,
          value: {
            name,
            attributes: parsed.value as StyleFootnoteSeparatorOptions["attributes"],
          },
        }
      : parsed;
  }
  if (name === "presentation:sound") {
    const parsed = parseAttributes(element, SOUND_ATTRIBUTES, scale);
    return parsed.ok
      ? { ok: true, value: { name, attributes: parsed.value as StyleSoundOptions["attributes"] } }
      : parsed;
  }
  if (name === "text:notes-configuration") {
    const parsed = parseAttributes(element, NOTES_CONFIGURATION_ATTRIBUTES, scale);
    if (!parsed.ok) return parsed;
    const continuation: Partial<Record<"forward" | "backward", string>> = {};
    for (const child of element.elements ?? []) {
      if (child.type !== "element") continue;
      const text = (child.elements ?? [])
        .filter((item) => item.type === "text")
        .map((item) => String(item.text ?? ""))
        .join("");
      if (child.name === "text:note-continuation-notice-forward") continuation.forward = text;
      else if (child.name === "text:note-continuation-notice-backward")
        continuation.backward = text;
      else
        return { ok: false, elementName: child.name ?? "", reason: "child has no ODF RNG mapping" };
    }
    return {
      ok: true,
      value: {
        name,
        attributes: parsed.value as StyleNotesConfigurationOptions["attributes"],
        ...(continuation.forward !== undefined
          ? { continuationForward: continuation.forward }
          : {}),
        ...(continuation.backward !== undefined
          ? { continuationBackward: continuation.backward }
          : {}),
      },
    };
  }
  return { ok: false, elementName: name, reason: "child has no ODF RNG mapping" };
}

function listLevelAttributes(name: string) {
  if (name === "text:list-level-style-number") return NUMBER_LEVEL_ATTRIBUTES;
  if (name === "text:list-level-style-bullet") return BULLET_LEVEL_ATTRIBUTES;
  if (name === "text:list-level-style-image") return IMAGE_LEVEL_ATTRIBUTES;
  return undefined;
}

function stylePropertyChildXml(
  child: StylePropertyChildOptions,
  scale: StylePropertyScale,
): string {
  if (child.name === "style:background-image") {
    return xmlElement(
      child.name,
      childAttributesXml(child.attributes, scale),
      child.data === undefined
        ? []
        : [xmlElement("office:binary-data", undefined, [encodeBase64(toUint8Array(child.data))])],
    );
  }
  if (child.name === "style:tab-stops")
    return xmlElement(
      child.name,
      undefined,
      child.stops.map((stop) =>
        xmlElement("style:tab-stop", childAttributesXml(stop.attributes, scale)),
      ),
    );
  if (child.name === "style:columns")
    return xmlElement(child.name, childAttributesXml(child.attributes, scale), [
      ...(child.separator
        ? [xmlElement("style:column-sep", childAttributesXml(child.separator.attributes, scale))]
        : []),
      ...child.columns.map((column) =>
        xmlElement("style:column", childAttributesXml(column.attributes, scale)),
      ),
    ]);
  if (child.name === "text:list-style")
    return xmlElement(
      child.name,
      childAttributesXml(child.attributes, scale),
      child.levels.map((level) =>
        xmlElement(level.name, childAttributesXml(level.attributes, scale), [
          ...(level.listProperties ? [listLevelPropertiesXml(level.listProperties, scale)] : []),
          ...(level.textProperties ? [stylePropertyXml(level.textProperties, scale)] : []),
        ]),
      ),
    );
  if (child.name === "text:notes-configuration")
    return xmlElement(child.name, childAttributesXml(child.attributes, scale), [
      ...(child.continuationForward === undefined
        ? []
        : [
            xmlElement("text:note-continuation-notice-forward", undefined, [
              escapeText(child.continuationForward),
            ]),
          ]),
      ...(child.continuationBackward === undefined
        ? []
        : [
            xmlElement("text:note-continuation-notice-backward", undefined, [
              escapeText(child.continuationBackward),
            ]),
          ]),
    ]);
  return xmlElement(child.name, childAttributesXml(child.attributes, scale));
}

function parseStyleChildValue(
  name: string,
  source: string,
  scale: StylePropertyScale,
): StylePropertyValue | undefined {
  if (CHILD_BOOLEAN_ATTRIBUTES.has(name)) {
    if (source !== "true" && source !== "false") return undefined;
    return source === "true";
  }
  if (name === "chart:axis-position") {
    const value = Number(source);
    return Number.isFinite(value) ? value : source;
  }
  if (CHILD_DOUBLE_ATTRIBUTES.has(name)) {
    const value = Number(source);
    return Number.isFinite(value) ? value : undefined;
  }
  if (CHILD_PERCENT_ATTRIBUTES.has(name)) {
    if (!source.endsWith("%")) return source;
    const value = Number.parseFloat(source);
    return Number.isFinite(value) ? value : undefined;
  }
  if (CHILD_INTEGER_ATTRIBUTES.has(name)) {
    const value = Number(source);
    return Number.isInteger(value) ? value : undefined;
  }
  if (CHILD_LENGTH_ATTRIBUTES.has(name)) {
    const emu = lengthToEmu(source);
    return emu === undefined ? source : childLengthValue(emu, scale);
  }
  return source;
}

function childAttributesXml(
  attributes: Record<string, StylePropertyValue>,
  scale: StylePropertyScale,
): XmlAttributes {
  return Object.fromEntries(
    Object.entries(attributes).map(([name, value]) => [
      name,
      childAttributeXml(name, value, scale),
    ]),
  );
}

function childAttributeXml(
  name: string,
  value: StylePropertyValue,
  scale: StylePropertyScale,
): StylePropertyValue {
  if (typeof value !== "number") return value;
  if (CHILD_LENGTH_ATTRIBUTES.has(name)) return childLengthXml(value, scale);
  if (CHILD_PERCENT_ATTRIBUTES.has(name)) return `${value}%`;
  return value;
}

function childLengthValue(emu: number, scale: StylePropertyScale): StylePropertyValue {
  if (scale === "emu") return emu;
  return scale === "spreadsheet" ? Math.round(emu / 9525) : Math.round(emu / 635);
}

function childLengthXml(value: number, scale: StylePropertyScale): string {
  const emu = scale === "emu" ? value : scale === "spreadsheet" ? value * 9525 : value * 635;
  return emuToLength(emu);
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
