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
  fontReference?: StyleMatrixReferenceOptions;
  /** Color element */
  color?: string;
}

export interface TableCellStyleOptions {
  /** Cell borders */
  borders?: TableCellBorderOptions;
  /** Fill reference (themeable) */
  fillReference?: StyleMatrixReferenceOptions;
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
  /** Line reference (a:lnRef) into the theme style matrix */
  lineReference?: StyleMatrixReferenceOptions;
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
  if (opts.lineReference !== undefined) {
    const children: string[] = [];
    if (opts.color) children.push(toStr(opts.color));
    return element(
      "a:lnRef",
      { idx: String(opts.lineReference.idx) },
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

/** Each cell-border side: Options key → a:tcBdr child element name. */
const BORDER_SIDES: ReadonlyArray<[keyof TableCellBorderOptions, string]> = [
  ["left", "a:left"],
  ["right", "a:right"],
  ["top", "a:top"],
  ["bottom", "a:bottom"],
  ["insideH", "a:insideH"],
  ["insideV", "a:insideV"],
];

/** CT_TableCellBorderStyle: each side wraps an EG_ThemeableLineStyle choice. */
function buildCellBorders(opts: TableCellBorderOptions): string {
  const children: string[] = [];
  for (const [key, name] of BORDER_SIDES) {
    const line = opts[key];
    if (line) children.push(element(name, undefined, [createThemeableLine(line)]));
  }
  return element("a:tcBdr", undefined, children);
}

function buildTextStyle(opts: TableTextStyleOptions): string {
  const children: string[] = [];
  if (opts.fontReference) children.push(createStyleMatrixRef("fontRef", opts.fontReference));
  if (opts.color) children.push(toStr(opts.color));
  const attrs: Record<string, string> = {};
  if (opts.bold && opts.bold !== "def") attrs.b = onOffAttr(opts.bold);
  if (opts.italic && opts.italic !== "def") attrs.i = onOffAttr(opts.italic);
  return element("a:tcTxStyle", attrs, children.length > 0 ? children : undefined);
}

function buildCellStyle(opts: TableCellStyleOptions): string {
  const children: string[] = [];
  if (opts.borders) children.push(buildCellBorders(opts.borders));
  if (opts.fillReference) children.push(createStyleMatrixRef("fillRef", opts.fillReference));
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

/** Parse a:tblStyleLst → TableStyleListOptions (reverse of createTableStyleList). */
export function parseTableStyleList(el: Element): TableStyleListOptions {
  const defaultStyleId = attr(el, "def") ?? "";
  const styles: TableStyleOptions[] = [];
  for (const child of el.elements ?? []) {
    if (child.name !== "a:tblStyle") continue;
    const style = parseTableStyle(child);
    if (style) styles.push(style);
  }
  return { defaultStyleId, ...(styles.length > 0 ? { styles } : {}) };
}

function parseTableStyle(el: Element): TableStyleOptions | undefined {
  const styleId = attr(el, "styleId");
  const styleName = attr(el, "styleName");
  if (styleId === undefined || styleName === undefined) return undefined;
  const regions: Partial<Record<TableStyleRegion, TablePartStyleOptions>> = {};
  for (const child of el.elements ?? []) {
    if (!child.name) continue;
    const region = ELEMENT_TO_REGION[child.name];
    if (!region) continue;
    const part = parsePartStyle(child);
    if (part) regions[region] = part;
  }
  return { styleId, styleName, ...(Object.keys(regions).length > 0 ? { regions } : {}) };
}

function parsePartStyle(el: Element): TablePartStyleOptions | undefined {
  const part: TablePartStyleOptions = {};
  const txStyle = findChild(el, "a:tcTxStyle");
  if (txStyle) {
    const text = parseTableTextStyle(txStyle);
    if (text) part.text = text;
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
  if (b === "on" || b === "off") opts.bold = b;
  const i = attr(el, "i");
  if (i === "on" || i === "off") opts.italic = i;
  const fontRefEl = findChild(el, "a:fontRef");
  if (fontRefEl) opts.fontReference = parseStyleMatrixRef(fontRefEl);
  // color: first non-fontRef child element, serialized back to the raw string
  // buildTextStyle stores (it pushes opts.color verbatim).
  for (const child of el.elements ?? []) {
    if (child.name === "a:fontRef") continue;
    opts.color = serializeChild(child);
    break;
  }
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
    // fill: first non-tcBdr child element, serialized (buildCellStyle pushes
    // opts.fill verbatim).
    for (const child of el.elements ?? []) {
      if (child.name === "a:tcBdr") continue;
      opts.fill = serializeChild(child);
      break;
    }
  }
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
  return Object.keys(borders).length > 0 ? borders : undefined;
}

function parseThemeableLine(el: Element): ThemeableLineStyleOptions | undefined {
  const opts: ThemeableLineStyleOptions = {};
  if (el.name === "a:lnRef") {
    const idx = attrNum(el, "idx");
    if (idx !== undefined) opts.lineReference = { idx };
  } else {
    const w = attrNum(el, "w");
    if (w !== undefined) opts.width = w;
  }
  for (const child of el.elements ?? []) {
    opts.color = serializeChild(child);
    break;
  }
  return Object.keys(opts).length > 0 ? opts : undefined;
}

function parseStyleMatrixRef(el: Element): StyleMatrixReferenceOptions {
  const idx = attrNum(el, "idx") ?? 0;
  const opts: StyleMatrixReferenceOptions = { idx };
  for (const child of el.elements ?? []) {
    opts.color = serializeChild(child);
    break;
  }
  return opts;
}
