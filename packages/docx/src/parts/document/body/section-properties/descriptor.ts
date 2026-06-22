/**
 * Section properties descriptor for DOCX documents.
 *
 * Produces `<w:sectPr>` XML directly from options, eliminating all
 * intermediate XmlComponent instances (create* + toXml pattern).
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_SectPr
 *
 * @module
 */

import { convertToTwip } from "@office-open/core";
import { twipsMeasureValue } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrBool, attrMeasure, attrNum, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { ColumnAttributes } from "@parts/document/body/section-properties/properties/column";
import type { ColumnsAttributes } from "@parts/document/body/section-properties/properties/columns";
import type {
  EndnotePropertiesOptions,
  FootnotePropertiesOptions,
} from "@parts/document/body/section-properties/properties/footnote-endnote-properties";
import type { PageBordersOptions } from "@parts/document/body/section-properties/properties/page-borders";
import { PageNumberSeparator } from "@parts/document/body/section-properties/properties/page-number";
import type { PageNumberTypeAttributes } from "@parts/document/body/section-properties/properties/page-number";
import type {
  SectionPropertiesChangeOptions,
  SectionPropertiesOptions,
} from "@parts/document/body/section-properties/section-properties";
import {
  sectionMarginDefaults,
  sectionPageSizeDefaults,
} from "@parts/document/body/section-properties/section-properties";
import type { HeaderFooterEntry } from "@parts/header-footer";
import type { BorderOptions } from "@shared/border";
import { NumberFormat } from "@shared/constants";
import type { BodyContext } from "@shared/index";

/** Valid page-number @w:fmt values (ST_NumberFormat). */
const PAGE_NUMBER_FORMATS = Object.values(NumberFormat) as readonly string[];
/** Valid page-number @w:chapSep values (ST_ChapterSep). */
const PAGE_NUMBER_SEPARATORS = Object.values(PageNumberSeparator) as readonly string[];

// ── Border XML helper ──

function stringifyBorderXml(tag: string, opts: BorderOptions): string {
  const attrs: string[] = [];
  if (opts.style !== undefined) attrs.push(`w:val="${opts.style}"`);
  if (opts.color !== undefined) attrs.push(`w:color="${opts.color}"`);
  if (opts.size !== undefined) attrs.push(`w:sz="${opts.size}"`);
  if (opts.space !== undefined) attrs.push(`w:space="${opts.space}"`);
  if (opts.themeColor !== undefined) attrs.push(`w:themeColor="${opts.themeColor}"`);
  if (opts.themeTint !== undefined) attrs.push(`w:themeTint="${opts.themeTint}"`);
  if (opts.themeShade !== undefined) attrs.push(`w:themeShade="${opts.themeShade}"`);
  if (opts.shadow !== undefined) attrs.push(`w:shadow="${opts.shadow ? 1 : 0}"`);
  if (opts.frame !== undefined) attrs.push(`w:frame="${opts.frame ? 1 : 0}"`);
  return `<${tag} ${attrs.join(" ")}/>`;
}

// ── Inline XML builders (replacing create* + toXml) ──

function pageSizeXml(
  w: number | string,
  h: number | string,
  orient?: string,
  code?: number,
): string {
  const attrs: string[] = [`w:w="${w}"`, `w:h="${h}"`];
  if (orient) attrs.push(`w:orient="${orient}"`);
  if (code !== undefined) attrs.push(`w:code="${code}"`);
  return `<w:pgSz ${attrs.join(" ")}/>`;
}

function pageMarginXml(
  top: number | string,
  right: number | string,
  bottom: number | string,
  left: number | string,
  header: number | string,
  footer: number | string,
  gutter: number | string,
): string {
  return `<w:pgMar w:top="${top}" w:right="${right}" w:bottom="${bottom}" w:left="${left}" w:header="${header}" w:footer="${footer}" w:gutter="${gutter}"/>`;
}

function headerFooterRefXml(tag: string, id: number, type: string): string {
  return `<${tag} r:id="rId${id}" w:type="${type}"/>`;
}

function sectionTypeXml(val: string): string {
  return `<w:type w:val="${val}"/>`;
}

function verticalAlignXml(val: string): string {
  return `<w:vAlign w:val="${val}"/>`;
}

function lineNumberXml(opts: NonNullable<SectionPropertiesOptions["lineNumbers"]>): string {
  const attrs: string[] = [];
  if (opts.countBy !== undefined) attrs.push(`w:countBy="${opts.countBy}"`);
  if (opts.start !== undefined) attrs.push(`w:start="${opts.start}"`);
  if (opts.restart !== undefined) attrs.push(`w:restart="${opts.restart}"`);
  if (opts.distance !== undefined) attrs.push(`w:distance="${opts.distance}"`);
  return attrs.length ? `<w:lnNumType ${attrs.join(" ")}/>` : "<w:lnNumType/>";
}

function pageNumberXml(opts: NonNullable<PageNumberTypeAttributes>): string {
  const attrs: string[] = [];
  if (opts.start !== undefined) attrs.push(`w:start="${opts.start}"`);
  if (opts.formatType !== undefined) attrs.push(`w:fmt="${opts.formatType}"`);
  if (opts.separator !== undefined) attrs.push(`w:chapSep="${opts.separator}"`);
  if (opts.chapStyle !== undefined) attrs.push(`w:chapStyle="${opts.chapStyle}"`);
  // No attributes → omit pgNumType (never fabricate an empty element).
  return attrs.length ? `<w:pgNumType ${attrs.join(" ")}/>` : "";
}

function docGridXml(linePitch: number, charSpace?: number, type?: string): string {
  const attrs: string[] = [`w:linePitch="${linePitch}"`];
  if (charSpace !== undefined) attrs.push(`w:charSpace="${charSpace}"`);
  if (type !== undefined) attrs.push(`w:type="${type}"`);
  return `<w:docGrid ${attrs.join(" ")}/>`;
}

function columnsXml(opts: NonNullable<SectionPropertiesOptions["column"]>): string {
  const attrs: string[] = [];
  if (opts.space !== undefined) attrs.push(`w:space="${twipsMeasureValue(opts.space)}"`);
  if (opts.count !== undefined) attrs.push(`w:num="${opts.count}"`);
  if (opts.separate !== undefined) attrs.push(`w:sep="${opts.separate ? 1 : 0}"`);
  if (opts.equalWidth !== undefined) attrs.push(`w:equalWidth="${opts.equalWidth ? 1 : 0}"`);

  const attrStr = attrs.join(" ");

  // Custom width columns — children are ColumnAttributes (Column class implements this interface)
  if (!opts.equalWidth && opts.children) {
    const colParts: string[] = [];
    for (const col of opts.children as readonly ColumnAttributes[]) {
      const colAttrs: string[] = [`w:w="${twipsMeasureValue(col.width)}"`];
      if (col.space !== undefined) colAttrs.push(`w:space="${twipsMeasureValue(col.space)}"`);
      colParts.push(`<w:col ${colAttrs.join(" ")}/>`);
    }
    return `<w:cols ${attrStr}>${colParts.join("")}</w:cols>`;
  }
  return `<w:cols ${attrStr}/>`;
}

function footnotePrXml(
  tag: string,
  opts: FootnotePropertiesOptions | EndnotePropertiesOptions,
): string {
  const parts: string[] = [];
  if (opts.pos !== undefined) parts.push(`<w:pos w:val="${opts.pos}"/>`);
  if (opts.formatType !== undefined || opts.format !== undefined) {
    const fmtAttrs: string[] = [];
    if (opts.formatType !== undefined) fmtAttrs.push(`w:fmt="${opts.formatType}"`);
    if (opts.format !== undefined) fmtAttrs.push(`w:format="${opts.format}"`);
    parts.push(`<w:numFmt ${fmtAttrs.join(" ")}/>`);
  }
  if (opts.numStart !== undefined) parts.push(`<w:numStart w:val="${opts.numStart}"/>`);
  if (opts.numRestart !== undefined) parts.push(`<w:numRestart w:val="${opts.numRestart}"/>`);
  const body = parts.join("");
  return body ? `<${tag}>${body}</${tag}>` : `<${tag}/>`;
}

function pageBordersXml(opts: NonNullable<PageBordersOptions>): string {
  const attrs: string[] = [];
  if (opts.display !== undefined) attrs.push(`w:display="${opts.display}"`);
  if (opts.offsetFrom !== undefined) attrs.push(`w:offsetFrom="${opts.offsetFrom}"`);
  if (opts.zOrder !== undefined) attrs.push(`w:zOrder="${opts.zOrder}"`);

  const parts: string[] = [];
  if (opts.top) parts.push(stringifyBorderXml("w:top", opts.top));
  if (opts.left) parts.push(stringifyBorderXml("w:left", opts.left));
  if (opts.bottom) parts.push(stringifyBorderXml("w:bottom", opts.bottom));
  if (opts.right) parts.push(stringifyBorderXml("w:right", opts.right));

  const attrStr = attrs.join(" ");
  const body = parts.join("");
  if (!body && !attrStr) return "<w:pgBorders/>";
  return body ? `<w:pgBorders ${attrStr}>${body}</w:pgBorders>` : `<w:pgBorders ${attrStr}/>`;
}

// ── Header/footer references ──

function appendHeaderFooterRefs(
  parts: string[],
  type: "w:headerReference" | "w:footerReference",
  group?: { default?: HeaderFooterEntry; first?: HeaderFooterEntry; even?: HeaderFooterEntry },
): void {
  if (!group) return;
  if (group.default) parts.push(headerFooterRefXml(type, group.default.referenceId, "default"));
  if (group.first) parts.push(headerFooterRefXml(type, group.first.referenceId, "first"));
  if (group.even) parts.push(headerFooterRefXml(type, group.even.referenceId, "even"));
}

// ── sectPrChange (recursive) ──

function stringifySectionPropertiesChange(opts: SectionPropertiesChangeOptions): string {
  const { author, date, id, ...inner } = opts;
  const innerXml = stringifySectionPropertiesInner(inner);
  return `<w:sectPrChange w:author="${author}" w:date="${date}" w:id="${id}"><w:sectPr>${innerXml}</w:sectPr></w:sectPrChange>`;
}

// ── Core XML builder ──

function stringifySectionPropertiesInner(opts: SectionPropertiesOptions): string {
  const parts: string[] = [];

  // Header/footer references
  appendHeaderFooterRefs(parts, "w:headerReference", opts.headerWrapperGroup);
  appendHeaderFooterRefs(parts, "w:footerReference", opts.footerWrapperGroup);

  // Destructure page options with defaults
  const {
    size: {
      width = sectionPageSizeDefaults.WIDTH,
      height = sectionPageSizeDefaults.HEIGHT,
      orientation = sectionPageSizeDefaults.ORIENTATION,
      code,
    } = {},
    margin: {
      top = sectionMarginDefaults.TOP,
      right = sectionMarginDefaults.RIGHT,
      bottom = sectionMarginDefaults.BOTTOM,
      left = sectionMarginDefaults.LEFT,
      header = sectionMarginDefaults.HEADER,
      footer = sectionMarginDefaults.FOOTER,
      gutter = sectionMarginDefaults.GUTTER,
    } = {},
    pageNumbers = {},
    borders,
    textDirection,
  } = opts.page ?? {};

  const { linePitch = 312, charSpace = 0, type: gridType = "lines" } = opts.grid ?? {};

  // Footnote/endnote properties
  if (opts.footnotePr) parts.push(footnotePrXml("w:footnotePr", opts.footnotePr));
  if (opts.endnotePr) parts.push(footnotePrXml("w:endnotePr", opts.endnotePr));

  // Section type
  if (opts.type) parts.push(sectionTypeXml(opts.type));

  // Page size — normalize both logical dimensions to twips, then swap w/h when
  // landscape. UniversalMeasure ("210mm") is converted so the emitted w:w/w:h is
  // always a plain twip number that the attrNum-based parse reads back exactly.
  const wTwips = convertToTwip(width);
  const hTwips = convertToTwip(height);
  const pgW = orientation === "landscape" ? hTwips : wTwips;
  const pgH = orientation === "landscape" ? wTwips : hTwips;
  parts.push(pageSizeXml(pgW, pgH, orientation, code));

  // Page margin (always present)
  parts.push(pageMarginXml(top, right, bottom, left, header, footer, gutter));

  // Page borders
  if (borders) parts.push(pageBordersXml(borders));

  // Line numbers
  if (opts.lineNumbers) parts.push(lineNumberXml(opts.lineNumbers));

  // Page numbers
  parts.push(pageNumberXml(pageNumbers));

  // Columns
  if (opts.column) parts.push(columnsXml(opts.column));

  // Vertical alignment
  if (opts.verticalAlign) parts.push(verticalAlignXml(opts.verticalAlign));

  // Boolean on/off elements — direct string output
  if (opts.titlePage !== undefined)
    parts.push(opts.titlePage ? "<w:titlePg/>" : '<w:titlePg w:val="0"/>');
  if (textDirection) parts.push(`<w:textDirection w:val="${textDirection}"/>`);
  if (opts.noEndnote !== undefined)
    parts.push(opts.noEndnote ? "<w:noEndnote/>" : '<w:noEndnote w:val="0"/>');
  if (opts.formProtection !== undefined)
    parts.push(opts.formProtection ? "<w:formProt/>" : '<w:formProt w:val="0"/>');
  if (opts.bidi !== undefined) parts.push(opts.bidi ? "<w:bidi/>" : '<w:bidi w:val="0"/>');
  if (opts.rtlGutter !== undefined)
    parts.push(opts.rtlGutter ? "<w:rtlGutter/>" : '<w:rtlGutter w:val="0"/>');

  // Paper source
  if (opts.paperSrc) {
    const psAttr: string[] = [];
    if (opts.paperSrc.first !== undefined) psAttr.push(`w:first="${opts.paperSrc.first}"`);
    if (opts.paperSrc.other !== undefined) psAttr.push(`w:other="${opts.paperSrc.other}"`);
    parts.push(`<w:paperSrc ${psAttr.join(" ")}/>`);
  }

  // Printer settings
  if (opts.printerSettingsId !== undefined) {
    parts.push(`<w:printerSettings r:id="${opts.printerSettingsId}"/>`);
  }

  // Document grid
  parts.push(docGridXml(linePitch, charSpace, gridType));

  // Revision (sectPrChange)
  if (opts.revision) {
    parts.push(stringifySectionPropertiesChange(opts.revision));
  }

  return parts.join("");
}

// ── Descriptor ──

/**
 * Section properties descriptor for DOCX `<w:sectPr>` elements.
 *
 * Produces complete XML directly from options — zero XmlComponent instances
 * in the hot path. All `create*()` + `.toXml()` calls eliminated in favor
 * of direct string concatenation.
 *
 * @example
 * ```typescript
 * const xml = sectionPropertiesDesc.stringify(sectPrOpts, ctx);
 * ```
 */
export const sectionPropertiesDesc: CustomDescriptor<SectionPropertiesOptions, BodyContext> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return stringifySectionPropertiesXml(opts);
  },

  parse(el, _ctx) {
    return parseSectionPropertiesEl(el);
  },
};

/** Standalone stringify — no context needed, pure options → XML. */
export function stringifySectionPropertiesXml(opts: SectionPropertiesOptions): string {
  const inner = stringifySectionPropertiesInner(opts);

  const attrs: string[] = [];
  if (opts.rsidRPr !== undefined) attrs.push(`w:rsidRPr="${opts.rsidRPr}"`);
  if (opts.rsidDel !== undefined) attrs.push(`w:rsidDel="${opts.rsidDel}"`);
  if (opts.rsidR !== undefined) attrs.push(`w:rsidR="${opts.rsidR}"`);
  if (opts.rsidSect !== undefined) attrs.push(`w:rsidSect="${opts.rsidSect}"`);

  const attrStr = attrs.length ? " " + attrs.join(" ") : "";
  return `<w:sectPr${attrStr}>${inner}</w:sectPr>`;
}

// ── Parse (Element → SectionPropertiesOptions) ──

/** Parse a w:sectPr element into SectionPropertiesOptions. */
export function parseSectionPropertiesEl(el: Element): SectionPropertiesOptions {
  const opts: Record<string, unknown> = {};

  // rsid attributes on w:sectPr element
  for (const [attrName, optKey] of [
    ["w:rsidR", "rsidR"],
    ["w:rsidRPr", "rsidRPr"],
    ["w:rsidDel", "rsidDel"],
    ["w:rsidSect", "rsidSect"],
  ] as const) {
    const val = attr(el, attrName);
    if (val) opts[optKey] = val;
  }

  // Page size
  const pgSz = findChild(el, "w:pgSz");
  if (pgSz) {
    const page: Record<string, unknown> = {};
    const size: Record<string, unknown> = {};
    const w = attrNum(pgSz, "w:w");
    const h = attrNum(pgSz, "w:h");
    const orient = attr(pgSz, "w:orient");
    if (orient === "landscape" && w !== undefined && h !== undefined) {
      size.width = h;
      size.height = w;
    } else {
      if (w !== undefined) size.width = w;
      if (h !== undefined) size.height = h;
    }
    if (orient) size.orientation = orient;
    const code = attrNum(pgSz, "w:code");
    if (code !== undefined) size.code = code;
    if (Object.keys(size).length > 0) page.size = size;

    // Page margins
    const pgMar = findChild(el, "w:pgMar");
    if (pgMar) {
      const margin: Record<string, unknown> = {};
      for (const [a, o] of [
        ["w:top", "top"],
        ["w:right", "right"],
        ["w:bottom", "bottom"],
        ["w:left", "left"],
        ["w:header", "header"],
        ["w:footer", "footer"],
        ["w:gutter", "gutter"],
      ] as const) {
        const val = attrNum(pgMar, a);
        if (val !== undefined) margin[o] = val;
      }
      if (Object.keys(margin).length > 0) page.margin = margin;
    }

    // Page number type
    const pgNumType = findChild(el, "w:pgNumType");
    if (pgNumType) {
      const pageNumbers: PageNumberTypeAttributes = {};
      const start = attrNum(pgNumType, "w:start");
      if (start !== undefined) pageNumbers.start = start;
      const fmt = attr(pgNumType, "w:fmt");
      if (fmt && PAGE_NUMBER_FORMATS.includes(fmt)) {
        pageNumbers.formatType = fmt as PageNumberTypeAttributes["formatType"];
      }
      const chapSep = attr(pgNumType, "w:chapSep");
      if (chapSep && PAGE_NUMBER_SEPARATORS.includes(chapSep)) {
        pageNumbers.separator = chapSep as PageNumberTypeAttributes["separator"];
      }
      const chapStyle = attrNum(pgNumType, "w:chapStyle");
      if (chapStyle !== undefined) pageNumbers.chapStyle = chapStyle;
      if (Object.keys(pageNumbers).length > 0) page.pageNumbers = pageNumbers;
    }

    if (Object.keys(page).length > 0) opts.page = page;
  }

  // Columns
  const cols = findChild(el, "w:cols");
  if (cols) {
    const column: ColumnsAttributes = {};
    const count = attrNum(cols, "w:num");
    if (count !== undefined) column.count = count;
    const space = attrMeasure(cols, "w:space");
    if (space !== undefined) column.space = space as ColumnsAttributes["space"];
    const separate = attrBool(cols, "w:sep");
    if (separate !== undefined) column.separate = separate;
    const equalWidth = attrBool(cols, "w:equalWidth");
    if (equalWidth !== undefined) column.equalWidth = equalWidth;
    const colChildren: ColumnAttributes[] = [];
    for (const colEl of cols.elements ?? []) {
      if (colEl.name !== "w:col") continue;
      const width = attrMeasure(colEl, "w:w");
      if (width === undefined) continue;
      const colAttr: ColumnAttributes = { width: width as ColumnAttributes["width"] };
      const colSpace = attrMeasure(colEl, "w:space");
      if (colSpace !== undefined) colAttr.space = colSpace as ColumnAttributes["space"];
      colChildren.push(colAttr);
    }
    if (colChildren.length > 0) column.children = colChildren;
    if (Object.keys(column).length > 0) opts.column = column;
  }

  // Section type
  const type = findChild(el, "w:type");
  if (type) {
    const val = attr(type, "w:val");
    if (val) opts.type = val;
  }

  // Title page
  const titlePg = findChild(el, "w:titlePg");
  if (titlePg) opts.titlePage = attrBool(titlePg, "w:val") ?? true;

  // On/off properties
  for (const [name, optKey] of [
    ["w:noEndnote", "noEndnote"],
    ["w:formProt", "formProtection"],
    ["w:bidi", "bidi"],
    ["w:rtlGutter", "rtlGutter"],
  ] as const) {
    const child = findChild(el, name);
    if (child) opts[optKey] = attrBool(child, "w:val") ?? true;
  }

  // Document grid
  const docGrid = findChild(el, "w:docGrid");
  if (docGrid) {
    const grid: Record<string, unknown> = {};
    const type = attr(docGrid, "w:type");
    if (type) grid.type = type;
    const linePitch = attrNum(docGrid, "w:linePitch");
    if (linePitch !== undefined) grid.linePitch = linePitch;
    const charSpace = attrNum(docGrid, "w:charSpace");
    if (charSpace !== undefined) grid.charSpace = charSpace;
    if (Object.keys(grid).length > 0) opts.grid = grid;
  }

  // Line numbers
  const lnNumType = findChild(el, "w:lnNumType");
  if (lnNumType) {
    const lineNumbers: Record<string, unknown> = {};
    const countBy = attrNum(lnNumType, "w:countBy");
    if (countBy !== undefined) lineNumbers.countBy = countBy;
    const start = attrNum(lnNumType, "w:start");
    if (start !== undefined) lineNumbers.start = start;
    const restart = attr(lnNumType, "w:restart");
    if (restart) lineNumbers.restart = restart;
    const distance = attrNum(lnNumType, "w:distance");
    if (distance !== undefined) lineNumbers.distance = distance;
    if (Object.keys(lineNumbers).length > 0) opts.lineNumbers = lineNumbers;
  }

  // Page borders
  const pgBorders = findChild(el, "w:pgBorders");
  if (pgBorders) {
    const borders: Record<string, unknown> = {};
    for (const side of ["top", "left", "bottom", "right"] as const) {
      const sideEl = findChild(pgBorders, `w:${side}`);
      if (sideEl) {
        const b: Record<string, unknown> = {};
        const val = attr(sideEl, "w:val");
        if (val) b.style = val;
        const color = attr(sideEl, "w:color");
        if (color) b.color = color;
        const sz = attrNum(sideEl, "w:sz");
        if (sz !== undefined) b.size = sz;
        const space = attrNum(sideEl, "w:space");
        if (space !== undefined) b.space = space;
        borders[side] = b;
      }
    }
    const display = attr(pgBorders, "w:display");
    if (display) borders.display = display;
    const offsetFrom = attr(pgBorders, "w:offsetFrom");
    if (offsetFrom) borders.offsetFrom = offsetFrom;
    const zOrder = attr(pgBorders, "w:zOrder");
    if (zOrder) borders.zOrder = zOrder;
    if (Object.keys(borders).length > 0) {
      const page = (opts.page ?? {}) as Record<string, unknown>;
      page.borders = borders;
      opts.page = page;
    }
  }

  // Vertical align
  const vAlign = findChild(el, "w:vAlign");
  if (vAlign) {
    const val = attr(vAlign, "w:val");
    if (val) opts.verticalAlign = val;
  }

  // Text direction
  const textDirection = findChild(el, "w:textDirection");
  if (textDirection) {
    const val = attr(textDirection, "w:val");
    if (val) {
      const page = (opts.page ?? {}) as Record<string, unknown>;
      page.textDirection = val;
      opts.page = page;
    }
  }

  // Footnote properties
  const footnotePr = findChild(el, "w:footnotePr");
  if (footnotePr) {
    opts.footnotePr = parseNotePropertiesEl(footnotePr);
  }

  // Endnote properties
  const endnotePr = findChild(el, "w:endnotePr");
  if (endnotePr) {
    opts.endnotePr = parseNotePropertiesEl(endnotePr);
  }

  // Paper source
  const paperSrc = findChild(el, "w:paperSrc");
  if (paperSrc) {
    const ps: Record<string, unknown> = {};
    const first = attrNum(paperSrc, "w:first");
    if (first !== undefined) ps.first = first;
    const other = attrNum(paperSrc, "w:other");
    if (other !== undefined) ps.other = other;
    if (Object.keys(ps).length > 0) opts.paperSrc = ps;
  }

  // Printer settings
  const printerSettings = findChild(el, "w:printerSettings");
  if (printerSettings) {
    const rId = attr(printerSettings, "r:id");
    if (rId) opts.printerSettingsId = rId;
  }

  // Header/footer references
  const headerGroup: Record<string, unknown> = {};
  const footerGroup: Record<string, unknown> = {};
  for (const child of el.elements ?? []) {
    if (child.name === "w:headerReference") {
      const type = attr(child, "w:type");
      const rId = attr(child, "r:id");
      if (type && rId) headerGroup[type] = { referenceId: parseInt(rId.replace("rId", ""), 10) };
    } else if (child.name === "w:footerReference") {
      const type = attr(child, "w:type");
      const rId = attr(child, "r:id");
      if (type && rId) footerGroup[type] = { referenceId: parseInt(rId.replace("rId", ""), 10) };
    }
  }
  if (Object.keys(headerGroup).length > 0) opts.headerWrapperGroup = headerGroup;
  if (Object.keys(footerGroup).length > 0) opts.footerWrapperGroup = footerGroup;

  // Revision (w:sectPrChange) — symmetric with stringifySectionPropertiesChange
  const sectPrChange = findChild(el, "w:sectPrChange");
  if (sectPrChange) {
    const rev: Record<string, unknown> = {};
    const author = attr(sectPrChange, "w:author");
    if (author) rev.author = author;
    const revDate = attr(sectPrChange, "w:date");
    if (revDate) rev.date = revDate;
    const revId = attrNum(sectPrChange, "w:id");
    if (revId !== undefined) rev.id = revId;
    const innerSectPr = findChild(sectPrChange, "w:sectPr");
    if (innerSectPr) Object.assign(rev, parseSectionPropertiesEl(innerSectPr));
    if (Object.keys(rev).length > 0) opts.revision = rev;
  }

  return opts as unknown as SectionPropertiesOptions;
}

function parseNotePropertiesEl(el: Element): Record<string, unknown> {
  const opts: Record<string, unknown> = {};

  const posEl = findChild(el, "w:pos");
  if (posEl) {
    const val = attr(posEl, "w:val");
    if (val) opts.pos = val;
  }

  const numFmt = findChild(el, "w:numFmt");
  if (numFmt) {
    const fmt = attr(numFmt, "w:fmt");
    if (fmt) opts.formatType = fmt;
    const format = attr(numFmt, "w:format");
    if (format) opts.format = format;
  }

  const numStart = findChild(el, "w:numStart");
  if (numStart) {
    const val = attrNum(numStart, "w:val");
    if (val !== undefined) opts.numStart = val;
  }

  const numRestart = findChild(el, "w:numRestart");
  if (numRestart) {
    const val = attr(numRestart, "w:val");
    if (val) opts.numRestart = val;
  }

  return opts;
}
