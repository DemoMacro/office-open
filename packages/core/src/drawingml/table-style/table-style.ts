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
import { BuilderElement } from "../../xml-components";
import type { XmlComponent } from "../../xml-components";

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
  readonly text?: TableTextStyleOptions;
  /** Cell style (fill, borders, 3D) */
  readonly cell?: TableCellStyleOptions;
}

export interface TableTextStyleOptions {
  /** Bold style */
  readonly bold?: OnOffStyleType;
  /** Italic style */
  readonly italic?: OnOffStyleType;
  /** Font reference (themeable) */
  readonly fontRef?: StyleMatrixReferenceOptions;
  /** Color element */
  readonly color?: XmlComponent;
}

export interface TableCellStyleOptions {
  /** Cell borders */
  readonly borders?: TableCellBorderOptions;
  /** Fill reference (themeable) */
  readonly fillRef?: StyleMatrixReferenceOptions;
  /** Direct fill */
  readonly fill?: XmlComponent;
}

export interface TableCellBorderOptions {
  readonly left?: ThemeableLineStyleOptions;
  readonly right?: ThemeableLineStyleOptions;
  readonly top?: ThemeableLineStyleOptions;
  readonly bottom?: ThemeableLineStyleOptions;
  readonly insideH?: ThemeableLineStyleOptions;
  readonly insideV?: ThemeableLineStyleOptions;
}

export interface ThemeableLineStyleOptions {
  /** Line width in EMUs */
  readonly width?: number;
  /** Fill color component */
  readonly color?: XmlComponent;
  /** Line reference index into theme style matrix */
  readonly lineRefIdx?: number;
}

export interface StyleMatrixReferenceOptions {
  /** Index into the theme style matrix */
  readonly idx: number;
  /** Color component */
  readonly color?: XmlComponent;
}

export interface TableStyleOptions {
  /** Unique GUID for this style */
  readonly styleId: string;
  /** Display name */
  readonly styleName: string;
  /** Style regions */
  readonly regions?: Partial<Record<TableStyleRegion, TablePartStyleOptions>>;
}

export interface TableStyleListOptions {
  /** Default style GUID */
  readonly defaultStyleId: string;
  /** Custom table styles */
  readonly styles?: readonly TableStyleOptions[];
}

// ── Helpers ──

function onOffAttr(val: OnOffStyleType | undefined): string {
  if (val === undefined) return "";
  return val;
}

/** Create a:a:styleMatrixReference (fillRef, lnRef, effectRef, fontRef) */
function createStyleMatrixRef(
  elementName: string,
  opts: StyleMatrixReferenceOptions,
): BuilderElement {
  const children: XmlComponent[] = [];
  if (opts.color) children.push(opts.color);
  return new BuilderElement({
    name: `a:${elementName}`,
    attributes: [{ key: "idx", value: String(opts.idx) }],
    children: children.length > 0 ? children : undefined,
  });
}

/** Create border line element (a:ln with color) or a:lnRef */
function createThemeableLine(opts: ThemeableLineStyleOptions): BuilderElement {
  if (opts.lineRefIdx !== undefined) {
    const children: XmlComponent[] = [];
    if (opts.color) children.push(opts.color);
    return new BuilderElement({
      name: "a:lnRef",
      attributes: [{ key: "idx", value: String(opts.lineRefIdx) }],
      children: children.length > 0 ? children : undefined,
    });
  }
  const children: XmlComponent[] = [];
  if (opts.color) children.push(opts.color);
  return new BuilderElement({
    name: "a:ln",
    attributes: opts.width !== undefined ? [{ key: "w", value: String(opts.width) }] : undefined,
    children: children.length > 0 ? children : undefined,
  });
}

// ── Part builders ──

function buildCellBorders(opts: TableCellBorderOptions): BuilderElement {
  const children: XmlComponent[] = [];
  if (opts.left) children.push(createThemeableLine(opts.left));
  if (opts.right) children.push(createThemeableLine(opts.right));
  if (opts.top) children.push(createThemeableLine(opts.top));
  if (opts.bottom) children.push(createThemeableLine(opts.bottom));
  if (opts.insideH) children.push(createThemeableLine(opts.insideH));
  if (opts.insideV) children.push(createThemeableLine(opts.insideV));
  return new BuilderElement({
    name: "a:tcBdr",
    children,
  });
}

function buildTextStyle(opts: TableTextStyleOptions): BuilderElement {
  const children: XmlComponent[] = [];
  if (opts.fontRef) children.push(createStyleMatrixRef("fontRef", opts.fontRef));
  if (opts.color) children.push(opts.color);
  const attrs: Array<{ key: string; value: string }> = [];
  if (opts.bold && opts.bold !== "def") attrs.push({ key: "b", value: onOffAttr(opts.bold) });
  if (opts.italic && opts.italic !== "def") attrs.push({ key: "i", value: onOffAttr(opts.italic) });
  return new BuilderElement({
    name: "a:tcTxStyle",
    attributes: attrs.length > 0 ? attrs : undefined,
    children: children.length > 0 ? children : undefined,
  });
}

function buildCellStyle(opts: TableCellStyleOptions): BuilderElement {
  const children: XmlComponent[] = [];
  if (opts.borders) children.push(buildCellBorders(opts.borders));
  if (opts.fillRef) children.push(createStyleMatrixRef("fillRef", opts.fillRef));
  else if (opts.fill) children.push(opts.fill);
  return new BuilderElement({
    name: "a:tcStyle",
    children,
  });
}

function buildPartStyle(opts: TablePartStyleOptions): BuilderElement {
  const children: XmlComponent[] = [];
  if (opts.text) children.push(buildTextStyle(opts.text));
  if (opts.cell) children.push(buildCellStyle(opts.cell));
  return new BuilderElement({
    name: "a:wholeTbl", // placeholder, replaced per region
    children,
  });
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
export function createTableStyle(opts: TableStyleOptions): BuilderElement {
  const children: XmlComponent[] = [];
  if (opts.regions) {
    for (const region of REGION_ORDER) {
      const part = opts.regions[region];
      if (!part) continue;
      const partEl = buildPartStyle(part);
      // Override the element name to the correct region
      (partEl as unknown as { root: { name: string } }).root.name = REGION_ELEMENTS[region];
      children.push(partEl);
    }
  }
  return new BuilderElement({
    name: "a:tblStyle",
    attributes: [
      { key: "styleId", value: opts.styleId },
      { key: "styleName", value: opts.styleName },
    ],
    children: children.length > 0 ? children : undefined,
  });
}

/** Create table style list (CT_TableStyleList) */
export function createTableStyleList(opts: TableStyleListOptions): BuilderElement {
  const children: XmlComponent[] = [];
  if (opts.styles) {
    for (const style of opts.styles) {
      children.push(createTableStyle(style));
    }
  }
  return new BuilderElement({
    name: "a:tblStyleLst",
    attributes: [
      { key: "xmlns:a", value: "http://schemas.openxmlformats.org/drawingml/2006/main" },
      { key: "def", value: opts.defaultStyleId },
    ],
    children: children.length > 0 ? children : undefined,
  });
}
