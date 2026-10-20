/**
 * Table Style system for DrawingML.
 *
 * Generates CT_TableStyle and CT_TableStyleList XML for custom table styles
 * in presentations (PPTX) and documents.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_TableStyle / CT_TableStyleList
 *
 * @module
 */
import { attr, attrNum, element, findChild, stringify } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { ReadContext, WriteContext } from "../descriptor";
import { parse as parseDesc, stringify as stringifyDesc } from "../descriptor";
import { xsdCompoundLine, xsdLineCap, xsdLineEndSize, xsdPenAlignment } from "../util/mappings";
import type { Percentage } from "../util/values";
import { parseColorChoice, stringifyColorChoice } from "./color/color-descriptors";
import type { SolidFillOptions } from "./color/solid-fill";
import { fillDesc } from "./fill/fill-descriptors";
import type { FillOptions } from "./fill/fill-options";
import { createCustomDash } from "./outline/custom-dash";
import { parsePositivePercentage } from "./outline/custom-dash";
import type { DashStop } from "./outline/custom-dash";
import { createLineEnd } from "./outline/line-end";
import type { LineEndOptions } from "./outline/line-end";
import type { CompoundLine, LineCap, LineJoin, PenAlignment, PresetDash } from "./outline/outline";
import type { StyleMatrixReferenceOptions } from "./style-reference";
import { cell3DDesc } from "./three-d/three-d-descriptors";
import type { Cell3DOptions } from "./three-d/three-d-descriptors";

export type { StyleMatrixReferenceOptions } from "./style-reference";

// cell3DDesc's stringify/parse never touch the context (bevel/lightRig are
// pure attribute emitters), so bare stubs satisfy the descriptor signature.
const EMPTY_WRITE_CTX = {} as WriteContext;
const EMPTY_READ_CTX = {} as ReadContext;

// ── Options ──

export type TableStyleRegion =
  | "tblBg"
  | "wholeTbl"
  | "band1H"
  | "band2H"
  | "band1V"
  | "band2V"
  | "lastCol"
  | "firstCol"
  | "lastRow"
  | "seCell"
  | "swCell"
  | "firstRow"
  | "neCell"
  | "nwCell";

/** Style-part on/off state (ST_OnOffStyleType). */
export type OnOffStyleType = "on" | "off" | "def";

/**
 * Font reference — a:fontRef (CT_FontReference). Unlike lnRef/fillRef/effectRef
 * (CT_StyleMatrixReference, integer idx), its @idx is a font-collection name
 * (ST_FontCollectionIndex: major/minor/none).
 */
export interface TableFontReferenceOptions {
  /** Theme font slot: "major" headings, "minor" body text, "none". */
  collection: "major" | "minor" | "none";
  color?: SolidFillOptions;
}

export interface TablePartStyleOptions {
  /** Cell text style */
  text?: TableTextStyleOptions;
  /** Cell style (fill, borders, 3D) */
  cell?: TableCellStyleOptions;
  /** Table background (a:tblBg only) — themeable fill/effect references. */
  background?: TableBackgroundStyleOptions;
}

/** Table background style (CT_TableBackgroundStyle: fill → effect). */
export interface TableBackgroundStyleOptions {
  /** Theme fill reference (a:fillRef idx). */
  fillReference?: StyleMatrixReferenceOptions;
  /** Explicit fill choice as raw XML (e.g. a:solidFill). */
  fill?: string;
  /** Theme effect reference (a:effectRef idx). */
  effectReference?: StyleMatrixReferenceOptions;
}

export interface TableTextStyleOptions {
  /** Bold style */
  bold?: OnOffStyleType;
  /** Italic style */
  italic?: OnOffStyleType;
  /** Underline style */
  underline?: OnOffStyleType;
  /** Font reference (themeable) */
  fontReference?: TableFontReferenceOptions;
  /** Color element */
  color?: string;
  /** `a:font` explicit font table; alternative to `fontReference`. */
  font?: string;
}

export interface TableCellStyleOptions {
  /** Cell borders */
  borders?: TableCellBorderOptions;
  /** Fill reference (themeable) */
  fillReference?: StyleMatrixReferenceOptions;
  /** Direct fill */
  fill?: string;
  /** Cell bevel (CT_Cell3D) */
  cell3D?: Cell3DOptions;
}

export interface TableCellBorderOptions {
  left?: ThemeableLineStyleOptions;
  right?: ThemeableLineStyleOptions;
  top?: ThemeableLineStyleOptions;
  bottom?: ThemeableLineStyleOptions;
  insideH?: ThemeableLineStyleOptions;
  insideV?: ThemeableLineStyleOptions;
  /** Diagonal border from the top-left to the bottom-right corner. */
  topLeftToBottomRight?: ThemeableLineStyleOptions;
  /** Diagonal border from the top-right to the bottom-left corner. */
  topRightToBottomLeft?: ThemeableLineStyleOptions;
  /** Trailing a:extLst inner XML for border extensions. */
  ext?: string;
}

export interface ThemeableLineStyleOptions {
  /** Line width in EMUs */
  width?: number;
  /** Line end cap (`a:ln/@cap`). */
  cap?: LineCap;
  /** Compound line type (@cmpd) */
  compound?: CompoundLine;
  /** Preset dash pattern (`a:prstDash`). */
  dash?: PresetDash;
  /** Custom dash stops (`a:custDash`); mutually exclusive with `dash`. */
  customDash?: readonly DashStop[];
  /** Line join style (`a:round`, `a:bevel`, or `a:miter`). */
  join?: LineJoin;
  /** Pen alignment (`a:ln/@algn`). */
  alignment?: PenAlignment;
  /** Miter join limit; `800` means 800% and `"800%"` preserves source syntax. */
  miterLimit?: number | Percentage;
  /** Line fill (`a:noFill`, `a:solidFill`, `a:gradFill`, `a:pattFill`, or `a:grpFill`). */
  fill?: FillOptions;
  /** Starting arrowhead (`a:headEnd`). */
  headEnd?: LineEndOptions;
  /** Ending arrowhead (`a:tailEnd`). */
  tailEnd?: LineEndOptions;
  /** Line reference (a:lnRef) into the theme style matrix */
  lineReference?: StyleMatrixReferenceOptions;
}

export interface TableStyleOptions {
  /** Unique GUID for this style */
  styleId: string;
  /** Display name */
  styleName: string;
  /** Style regions */
  regions?: Partial<Record<TableStyleRegion, TablePartStyleOptions>>;
}

/** The table-style gallery (p:tblStyleLst / w:tblStyles — default GUID plus custom styles). */
export interface TableStyleListOptions {
  /** Default style GUID */
  defaultStyleId: string;
  /** Custom table styles */
  styles?: readonly TableStyleOptions[];
  /**
   * Unmodeled trailing child elements (markup-compatibility blocks above
   * all), serialized verbatim — re-emitted as written after the styles.
   */
  ext?: string;
}

// ── Helpers ──

/** Pass through a string value (all color/fill factories now return string). */
function toStr(el: string): string {
  return el;
}

/** Create a:styleMatrixReference (fillRef, lnRef, effectRef, fontRef) */
function createStyleMatrixRef(elementName: string, opts: StyleMatrixReferenceOptions): string {
  const children: string[] = [];
  if (opts.color) children.push(stringifyColorChoice(opts.color, EMPTY_WRITE_CTX));
  return element(
    `a:${elementName}`,
    { idx: String(opts.index) },
    children.length > 0 ? children : undefined,
  );
}

/** Create border line element (a:ln with color) or a:lnRef */
function createThemeableLine(opts: ThemeableLineStyleOptions): string {
  if (opts.lineReference !== undefined) {
    const children: string[] = [];
    if (opts.lineReference.color)
      children.push(stringifyColorChoice(opts.lineReference.color, EMPTY_WRITE_CTX));
    return element(
      "a:lnRef",
      { idx: String(opts.lineReference.index) },
      children.length > 0 ? children : undefined,
    );
  }
  const children: string[] = [];
  const attrs: Record<string, string> = {};
  if (opts.width !== undefined) attrs.w = String(opts.width);
  if (opts.compound) attrs.cmpd = xsdCompoundLine.to(opts.compound);
  if (opts.cap) attrs.cap = xsdLineCap.to(opts.cap);
  if (opts.alignment) attrs.algn = xsdPenAlignment.to(opts.alignment);
  if (opts.fill) children.push(fillDesc.stringify(opts.fill, EMPTY_WRITE_CTX) ?? "");
  if (opts.customDash) children.push(createCustomDash(opts.customDash));
  else if (opts.dash) children.push(`<a:prstDash val="${opts.dash}"/>`);
  if (opts.join === "miter") {
    const limit =
      opts.miterLimit === undefined
        ? undefined
        : typeof opts.miterLimit === "string"
          ? opts.miterLimit
          : opts.miterLimit * 1000;
    children.push(`<a:miter${limit ? ` lim="${limit}"` : ""}/>`);
  } else if (opts.join) {
    children.push(`<a:${opts.join}/>`);
  }
  if (opts.headEnd) children.push(createLineEnd("a:headEnd", opts.headEnd));
  if (opts.tailEnd) children.push(createLineEnd("a:tailEnd", opts.tailEnd));
  return element("a:ln", attrs, children.length > 0 ? children : undefined);
}

// ── Part builders ──

/** Each cell-border side: Options key → a:tcBdr child element name. */
const BORDER_SIDES: ReadonlyArray<[Exclude<keyof TableCellBorderOptions, "ext">, string]> = [
  ["left", "a:left"],
  ["right", "a:right"],
  ["top", "a:top"],
  ["bottom", "a:bottom"],
  ["insideH", "a:insideH"],
  ["insideV", "a:insideV"],
  ["topLeftToBottomRight", "a:tl2br"],
  ["topRightToBottomLeft", "a:tr2bl"],
];

/** CT_TableCellBorderStyle: each side wraps an EG_ThemeableLineStyle choice. */
function buildCellBorders(opts: TableCellBorderOptions): string {
  const children: string[] = [];
  for (const [key, name] of BORDER_SIDES) {
    const line = opts[key];
    if (line) children.push(element(name, undefined, [createThemeableLine(line)]));
  }
  if (opts.ext) children.push(`<a:extLst>${opts.ext}</a:extLst>`);
  return element("a:tcBdr", undefined, children);
}

function buildTextStyle(opts: TableTextStyleOptions): string {
  const children: string[] = [];
  if (opts.fontReference) {
    const color = opts.fontReference.color
      ? stringifyColorChoice(opts.fontReference.color, EMPTY_WRITE_CTX)
      : "";
    children.push(`<a:fontRef idx="${opts.fontReference.collection}">${color}</a:fontRef>`);
  }
  if (opts.font) children.push(toStr(opts.font));
  if (opts.color) children.push(toStr(opts.color));
  const attrs: Record<string, string> = {};
  if (opts.bold) attrs.b = opts.bold;
  if (opts.italic) attrs.i = opts.italic;
  if (opts.underline) attrs.u = opts.underline;
  return element("a:tcTxStyle", attrs, children.length > 0 ? children : undefined);
}

function buildCellStyle(opts: TableCellStyleOptions): string {
  const children: string[] = [];
  if (opts.borders) children.push(buildCellBorders(opts.borders));
  if (opts.fillReference) children.push(createStyleMatrixRef("fillRef", opts.fillReference));
  else if (opts.fill) children.push(toStr(opts.fill));
  if (opts.cell3D) children.push(stringifyDesc(cell3DDesc, opts.cell3D, EMPTY_WRITE_CTX) ?? "");
  return element("a:tcStyle", undefined, children);
}

/**
 * Build one region element. a:tblBg is CT_TableBackgroundStyle (themeable
 * fill/effect only); every other region is CT_TablePartStyle (tcTxStyle +
 * tcStyle only) — the two child sets never mix per XSD.
 */
function buildPartStyle(
  elementName: string,
  opts: TablePartStyleOptions,
  isBackground: boolean,
): string {
  const children: string[] = [];
  if (isBackground) {
    const background = opts.background;
    if (background) {
      if (background.fillReference)
        children.push(createStyleMatrixRef("fillRef", background.fillReference));
      else if (background.fill) children.push(toStr(background.fill));
      if (background.effectReference)
        children.push(createStyleMatrixRef("effectRef", background.effectReference));
    }
  } else {
    if (opts.text) children.push(buildTextStyle(opts.text));
    if (opts.cell) children.push(buildCellStyle(opts.cell));
  }
  return element(elementName, undefined, children);
}

// ── Public API ──

/** XSD ordering for CT_TableStyle children */
const REGION_ORDER: readonly TableStyleRegion[] = [
  "tblBg",
  "wholeTbl",
  "band1H",
  "band2H",
  "band1V",
  "band2V",
  "lastCol",
  "firstCol",
  "lastRow",
  "seCell",
  "swCell",
  "firstRow",
  "neCell",
  "nwCell",
];

/** XML element name for each region */
const REGION_ELEMENTS: Record<TableStyleRegion, string> = {
  tblBg: "a:tblBg",
  wholeTbl: "a:wholeTbl",
  band1H: "a:band1H",
  band2H: "a:band2H",
  band1V: "a:band1V",
  band2V: "a:band2V",
  lastCol: "a:lastCol",
  firstCol: "a:firstCol",
  lastRow: "a:lastRow",
  seCell: "a:seCell",
  swCell: "a:swCell",
  firstRow: "a:firstRow",
  neCell: "a:neCell",
  nwCell: "a:nwCell",
};

/** Create a single table style (CT_TableStyle). `elementName` covers the part-list form (a:tblStyle) and the inline a:tblPr form (a:tableStyle). */
export function createTableStyle(opts: TableStyleOptions, elementName = "a:tblStyle"): string {
  const children: string[] = [];
  if (opts.regions) {
    for (const region of REGION_ORDER) {
      const part = opts.regions[region];
      if (!part) continue;
      children.push(buildPartStyle(REGION_ELEMENTS[region], part, region === "tblBg"));
    }
  }
  return element(
    elementName,
    { styleId: opts.styleId, styleName: opts.styleName },
    children.length > 0 ? children : undefined,
  );
}

/** Create table style list (CT_TableStyleList) */
export function createTableStyleList(opts: TableStyleListOptions): string {
  const children: string[] = [];
  if (opts.styles) {
    for (const style of opts.styles) {
      children.push(createTableStyle(style));
    }
  }
  if (opts.ext) children.push(opts.ext);
  return element(
    "a:tblStyleLst",
    {
      "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      def: opts.defaultStyleId,
    },
    children.length > 0 ? children : undefined,
  );
}

// ── Parse (Element → TableStyleListOptions) ──

/**
 * Serialize a single parsed element. `stringify` treats its argument as a
 * document root and serializes the `elements` array, so wrap the child to
 * serialize the child itself (round-trips raw color/fill child elements).
 */
function serializeChild(child: Element): string {
  return stringify({ elements: [child] });
}

/** Reverse of REGION_ELEMENTS: XML element name → region key. */
const ELEMENT_TO_REGION = Object.fromEntries(
  Object.entries(REGION_ELEMENTS).map(([region, name]) => [name, region]),
) as Record<string, TableStyleRegion>;

const COLOR_CHOICE_ELEMENTS = new Set([
  "a:scrgbClr",
  "a:srgbClr",
  "a:hslClr",
  "a:sysClr",
  "a:schemeClr",
  "a:prstClr",
]);

const FILL_ELEMENTS = new Set(["a:noFill", "a:solidFill", "a:gradFill", "a:pattFill", "a:grpFill"]);

/** Parse a:tblStyleLst → TableStyleListOptions (reverse of createTableStyleList). */
export function parseTableStyleList(el: Element): TableStyleListOptions {
  const defaultStyleId = attr(el, "def") ?? "";
  const styles: TableStyleOptions[] = [];
  const extParts: string[] = [];
  for (const child of el.elements ?? []) {
    if (child.name !== "a:tblStyle") {
      // Markup-compatibility blocks and other unmodeled children stay
      // verbatim — re-serialized as written after the parsed styles.
      extParts.push(serializeChild(child));
      continue;
    }
    const style = parseTableStyle(child);
    if (style) styles.push(style);
  }
  return {
    defaultStyleId,
    ...(styles.length > 0 ? { styles } : {}),
    ...(extParts.length > 0 ? { ext: extParts.join("") } : {}),
  };
}

export function parseTableStyle(el: Element): TableStyleOptions | undefined {
  const styleId = attr(el, "styleId");
  const styleName = attr(el, "styleName");
  if (styleId === undefined || styleName === undefined) return undefined;
  const regions: Partial<Record<TableStyleRegion, TablePartStyleOptions>> = {};
  for (const child of el.elements ?? []) {
    if (!child.name) continue;
    const region = ELEMENT_TO_REGION[child.name];
    if (!region) continue;
    const part = region === "tblBg" ? parseTableBackgroundStyle(child) : parsePartStyle(child);
    if (part) regions[region] = part;
  }
  return { styleId, styleName, ...(Object.keys(regions).length > 0 ? { regions } : {}) };
}

/** Parse a:tblBg (CT_TableBackgroundStyle: fill → effect, no tc* wrappers). */
function parseTableBackgroundStyle(el: Element): TablePartStyleOptions | undefined {
  const background: TableBackgroundStyleOptions = {};
  const fillRefEl = findChild(el, "a:fillRef");
  if (fillRefEl) {
    background.fillReference = parseStyleMatrixRef(fillRefEl);
  } else {
    const explicitFill = (el.elements ?? []).find((c) => c.name !== "a:effectRef");
    if (explicitFill) background.fill = serializeChild(explicitFill);
  }
  const effectRefEl = findChild(el, "a:effectRef");
  if (effectRefEl) background.effectReference = parseStyleMatrixRef(effectRefEl);
  return { background };
}

function parsePartStyle(el: Element): TablePartStyleOptions | undefined {
  const part: TablePartStyleOptions = {};
  const txStyle = findChild(el, "a:tcTxStyle");
  if (txStyle) {
    const text = parseTableTextStyle(txStyle);
    part.text = text ?? {};
  }
  const cellStyle = findChild(el, "a:tcStyle");
  if (cellStyle) {
    const cell = parseTableCellStyle(cellStyle);
    if (cell) part.cell = cell;
  }
  return Object.keys(part).length > 0 ? part : undefined;
}

function parseTableTextStyle(el: Element): TableTextStyleOptions | undefined {
  const opts: TableTextStyleOptions = {};
  const b = attr(el, "b");
  if (b === "on" || b === "off" || b === "def") opts.bold = b;
  const i = attr(el, "i");
  if (i === "on" || i === "off" || i === "def") opts.italic = i;
  const u = attr(el, "u");
  if (u === "on" || u === "off" || u === "def") opts.underline = u;
  const fontRefEl = findChild(el, "a:fontRef");
  if (fontRefEl) {
    const idx = attr(fontRefEl, "idx");
    if (idx === "major" || idx === "minor" || idx === "none") {
      const color = parseColorChoice(fontRefEl, EMPTY_READ_CTX);
      opts.fontReference =
        color && Object.keys(color).length > 0 ? { collection: idx, color } : { collection: idx };
    }
  }
  const fontEl = (el.elements ?? []).find((child) => child.name === "a:font");
  if (fontEl) opts.font = serializeChild(fontEl);
  const colorEl = (el.elements ?? []).find((child) => COLOR_CHOICE_ELEMENTS.has(child.name ?? ""));
  if (colorEl) opts.color = serializeChild(colorEl);
  return Object.keys(opts).length > 0 ? opts : undefined;
}

function parseTableCellStyle(el: Element): TableCellStyleOptions | undefined {
  const opts: TableCellStyleOptions = {};
  const tcBdr = findChild(el, "a:tcBdr");
  if (tcBdr) {
    const borders = parseCellBorders(tcBdr);
    if (borders) opts.borders = borders;
  }
  const fillRefEl = findChild(el, "a:fillRef");
  if (fillRefEl) {
    opts.fillReference = parseStyleMatrixRef(fillRefEl);
  } else {
    // fill: first non-tcBdr/non-cell3D child element, serialized
    // (buildCellStyle pushes opts.fill verbatim).
    for (const child of el.elements ?? []) {
      if (child.type !== "element") continue;
      if (child.name === "a:tcBdr" || child.name === "a:cell3D") continue;
      opts.fill = serializeChild(child);
      break;
    }
  }
  const cell3DEl = findChild(el, "a:cell3D");
  if (cell3DEl) opts.cell3D = parseDesc(cell3DDesc, cell3DEl, EMPTY_READ_CTX);
  return Object.keys(opts).length > 0 ? opts : undefined;
}

function parseCellBorders(el: Element): TableCellBorderOptions | undefined {
  const borders: TableCellBorderOptions = {};
  for (const [key, name] of BORDER_SIDES) {
    const borderEl = findChild(el, name);
    if (!borderEl) continue;
    // Each side wraps an EG_ThemeableLineStyle choice (a:ln or a:lnRef).
    const lineEl = borderEl.elements?.find((c) => c.type === "element");
    if (!lineEl) continue;
    const line = parseThemeableLine(lineEl);
    if (line) borders[key] = line;
  }
  const extLst = findChild(el, "a:extLst");
  if (extLst) borders.ext = serializeChild(extLst).replace(/^<a:extLst>|<\/a:extLst>$/g, "");
  // An empty a:tcBdr is meaningful — it overrides the outer-region borders
  // with none — so presence round-trips as an empty borders object.
  return borders;
}

function parseThemeableLine(el: Element): ThemeableLineStyleOptions | undefined {
  const opts: ThemeableLineStyleOptions = {};
  if (el.name === "a:lnRef") {
    opts.lineReference = parseStyleMatrixRef(el);
  } else {
    const w = attrNum(el, "w");
    if (w !== undefined) opts.width = w;
    const cap = attr(el, "cap");
    if (cap !== undefined)
      opts.cap = xsdLineCap.from(String(cap)) as ThemeableLineStyleOptions["cap"];
    const cmpd = attr(el, "cmpd");
    if (cmpd) opts.compound = xsdCompoundLine.from(cmpd) as CompoundLine;
    const alignment = attr(el, "algn");
    if (alignment)
      opts.alignment = xsdPenAlignment.from(
        String(alignment),
      ) as ThemeableLineStyleOptions["alignment"];
  }
  const fillElement = (el.elements ?? []).find((child) => FILL_ELEMENTS.has(child.name ?? ""));
  if (fillElement) opts.fill = fillDesc.parse(fillElement, EMPTY_READ_CTX);
  const presetDash = findChild(el, "a:prstDash");
  if (presetDash?.attributes?.["val"] !== undefined) {
    opts.dash = String(presetDash.attributes["val"]) as ThemeableLineStyleOptions["dash"];
  }
  const customDash = findChild(el, "a:custDash");
  if (customDash) {
    opts.customDash = (customDash.elements ?? [])
      .filter((stop) => stop.name === "a:ds")
      .map((stop) => ({
        d: parsePositivePercentage(attr(stop, "d") ?? ""),
        sp: parsePositivePercentage(attr(stop, "sp") ?? ""),
      }));
  }
  if (findChild(el, "a:round")) opts.join = "round";
  else if (findChild(el, "a:bevel")) opts.join = "bevel";
  else {
    const miter = findChild(el, "a:miter");
    if (miter) {
      opts.join = "miter";
      const limit = String(attr(miter, "lim") ?? "");
      if (limit)
        opts.miterLimit = limit.endsWith("%") ? (limit as Percentage) : Number(limit) / 1000;
    }
  }
  const headEnd = findChild(el, "a:headEnd");
  if (headEnd) opts.headEnd = parseLineEnd(headEnd);
  const tailEnd = findChild(el, "a:tailEnd");
  if (tailEnd) opts.tailEnd = parseLineEnd(tailEnd);
  return Object.keys(opts).length > 0 ? opts : undefined;
}

function parseLineEnd(el: Element): LineEndOptions {
  const options: LineEndOptions = {};
  const type = attr(el, "type");
  if (type) options.type = type as LineEndOptions["type"];
  const width = attr(el, "w");
  if (width) options.width = xsdLineEndSize.from(width) as LineEndOptions["width"];
  const length = attr(el, "len");
  if (length) options.length = xsdLineEndSize.from(length) as LineEndOptions["length"];
  return options;
}

function parseStyleMatrixRef(el: Element): StyleMatrixReferenceOptions {
  const idx = attrNum(el, "idx") ?? 0;
  const color = parseColorChoice(el, EMPTY_READ_CTX);
  const opts: StyleMatrixReferenceOptions =
    color && Object.keys(color).length > 0 ? { index: idx, color } : { index: idx };
  return opts;
}
