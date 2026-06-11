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
import { element } from "@office-open/xml";

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

export type OnOffStyleType = "on" | "off" | "def";

export interface TablePartStyleOptions {
  /** Cell text style */
  text?: TableTextStyleOptions;
  /** Cell style (fill, borders, 3D) */
  cell?: TableCellStyleOptions;
}

export interface TableTextStyleOptions {
  /** Bold style */
  bold?: OnOffStyleType;
  /** Italic style */
  italic?: OnOffStyleType;
  /** Font reference (themeable) */
  fontRef?: StyleMatrixReferenceOptions;
  /** Color element */
  color?: string;
}

export interface TableCellStyleOptions {
  /** Cell borders */
  borders?: TableCellBorderOptions;
  /** Fill reference (themeable) */
  fillRef?: StyleMatrixReferenceOptions;
  /** Direct fill */
  fill?: string;
}

export interface TableCellBorderOptions {
  left?: ThemeableLineStyleOptions;
  right?: ThemeableLineStyleOptions;
  top?: ThemeableLineStyleOptions;
  bottom?: ThemeableLineStyleOptions;
  insideH?: ThemeableLineStyleOptions;
  insideV?: ThemeableLineStyleOptions;
}

export interface ThemeableLineStyleOptions {
  /** Line width in EMUs */
  width?: number;
  /** Fill color component */
  color?: string;
  /** Line reference index into theme style matrix */
  lineRefIdx?: number;
}

export interface StyleMatrixReferenceOptions {
  /** Index into the theme style matrix */
  idx: number;
  /** Color component */
  color?: string;
}

export interface TableStyleOptions {
  /** Unique GUID for this style */
  styleId: string;
  /** Display name */
  styleName: string;
  /** Style regions */
  regions?: Partial<Record<TableStyleRegion, TablePartStyleOptions>>;
}

export interface TableStyleListOptions {
  /** Default style GUID */
  defaultStyleId: string;
  /** Custom table styles */
  styles?: readonly TableStyleOptions[];
}

// ── Helpers ──

function onOffAttr(val: OnOffStyleType | undefined): string {
  if (val === undefined) return "";
  return val;
}

/** Pass through a string value (all color/fill factories now return string). */
function toStr(el: string): string {
  return el;
}

/** Create a:styleMatrixReference (fillRef, lnRef, effectRef, fontRef) */
function createStyleMatrixRef(elementName: string, opts: StyleMatrixReferenceOptions): string {
  const children: string[] = [];
  if (opts.color) children.push(opts.color);
  return element(
    `a:${elementName}`,
    { idx: String(opts.idx) },
    children.length > 0 ? children : undefined,
  );
}

/** Create border line element (a:ln with color) or a:lnRef */
function createThemeableLine(opts: ThemeableLineStyleOptions): string {
  if (opts.lineRefIdx !== undefined) {
    const children: string[] = [];
    if (opts.color) children.push(toStr(opts.color));
    return element(
      "a:lnRef",
      { idx: String(opts.lineRefIdx) },
      children.length > 0 ? children : undefined,
    );
  }
  const children: string[] = [];
  if (opts.color) children.push(toStr(opts.color));
  const attrs: Record<string, string> = {};
  if (opts.width !== undefined) attrs.w = String(opts.width);
  return element("a:ln", attrs, children.length > 0 ? children : undefined);
}

// ── Part builders ──

function buildCellBorders(opts: TableCellBorderOptions): string {
  const children: string[] = [];
  if (opts.left) children.push(createThemeableLine(opts.left));
  if (opts.right) children.push(createThemeableLine(opts.right));
  if (opts.top) children.push(createThemeableLine(opts.top));
  if (opts.bottom) children.push(createThemeableLine(opts.bottom));
  if (opts.insideH) children.push(createThemeableLine(opts.insideH));
  if (opts.insideV) children.push(createThemeableLine(opts.insideV));
  return element("a:tcBdr", undefined, children);
}

function buildTextStyle(opts: TableTextStyleOptions): string {
  const children: string[] = [];
  if (opts.fontRef) children.push(createStyleMatrixRef("fontRef", opts.fontRef));
  if (opts.color) children.push(toStr(opts.color));
  const attrs: Record<string, string> = {};
  if (opts.bold && opts.bold !== "def") attrs.b = onOffAttr(opts.bold);
  if (opts.italic && opts.italic !== "def") attrs.i = onOffAttr(opts.italic);
  return element("a:tcTxStyle", attrs, children.length > 0 ? children : undefined);
}

function buildCellStyle(opts: TableCellStyleOptions): string {
  const children: string[] = [];
  if (opts.borders) children.push(buildCellBorders(opts.borders));
  if (opts.fillRef) children.push(createStyleMatrixRef("fillRef", opts.fillRef));
  else if (opts.fill) children.push(toStr(opts.fill));
  return element("a:tcStyle", undefined, children);
}

function buildPartStyle(elementName: string, opts: TablePartStyleOptions): string {
  const children: string[] = [];
  if (opts.text) children.push(buildTextStyle(opts.text));
  if (opts.cell) children.push(buildCellStyle(opts.cell));
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

/** Create a single table style (CT_TableStyle) */
export function createTableStyle(opts: TableStyleOptions): string {
  const children: string[] = [];
  if (opts.regions) {
    for (const region of REGION_ORDER) {
      const part = opts.regions[region];
      if (!part) continue;
      children.push(buildPartStyle(REGION_ELEMENTS[region], part));
    }
  }
  return element(
    "a:tblStyle",
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
  return element(
    "a:tblStyleLst",
    {
      "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      def: opts.defaultStyleId,
    },
    children.length > 0 ? children : undefined,
  );
}
