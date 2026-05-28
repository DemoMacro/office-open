import type { ISectionPropertiesOptions } from "@file/document/body/section-properties/section-properties";
import type { SectionOptions } from "@file/file";
import type { SectionChild } from "@file/section-child";
/**
 * Body parser for DOCX documents.
 *
 * Parses w:body → SectionOptions[] by splitting at w:sectPr boundaries.
 *
 * @module
 */
import { RawPassthrough } from "@office-open/core";
import { attr, attrBool, attrNum, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import { parseAltChunk } from "../file/alt-chunk/alt-chunk-parse";
import { parseParagraph } from "../file/paragraph/paragraph-parse";
import { parseSdtBlock, setSectionChildrenParser } from "../file/sdt/sdt-parse";
import { parseSubDoc } from "../file/sub-doc/sub-doc-parse";
import { parseToc } from "../file/table-of-contents/toc-parse";
import { parseTable, setSectionChildParser } from "../file/table/table-parse";
import { setTextboxSectionChildrenParser, parseTextbox } from "../file/textbox/textbox-parse";
import { ParseContext } from "./context";

// ── Section properties parser ────────────────────────────────────────────────

/**
 * Parse w:sectPr element into ISectionPropertiesOptions.
 */
function parseSectionProperties(el: Element, ctx: ParseContext): ISectionPropertiesOptions {
  const opts: Record<string, unknown> = {};

  // Page size
  const pgSz = findChild(el, "w:pgSz");
  if (pgSz) {
    const page: Record<string, unknown> = {};
    const size: Record<string, unknown> = {};
    const w = attrNum(pgSz, "w:w");
    const h = attrNum(pgSz, "w:h");
    const orient = attr(pgSz, "w:orient");
    // createPageSize swaps w/h for landscape, so reverse that here
    if (orient === "landscape" && w !== undefined && h !== undefined) {
      size.width = h;
      size.height = w;
    } else {
      if (w !== undefined) size.width = w;
      if (h !== undefined) size.height = h;
    }
    if (orient) size.orientation = orient;
    if (Object.keys(size).length > 0) page.size = size;

    // Page margins
    const pgMar = findChild(el, "w:pgMar");
    if (pgMar) {
      const margin: Record<string, unknown> = {};
      for (const [attrName, optName] of [
        ["w:top", "top"],
        ["w:right", "right"],
        ["w:bottom", "bottom"],
        ["w:left", "left"],
        ["w:header", "header"],
        ["w:footer", "footer"],
        ["w:gutter", "gutter"],
      ] as const) {
        const val = attrNum(pgMar, attrName);
        if (val !== undefined) margin[optName] = val;
      }
      if (Object.keys(margin).length > 0) page.margin = margin;
    }

    // Page number type
    const pgNumType = findChild(el, "w:pgNumType");
    if (pgNumType) {
      const pageNumbers: Record<string, unknown> = {};
      const start = attrNum(pgNumType, "w:start");
      if (start !== undefined) pageNumbers.start = start;
      const fmt = attr(pgNumType, "w:fmt");
      if (fmt) pageNumbers.formatType = fmt;
      if (Object.keys(pageNumbers).length > 0) page.pageNumbers = pageNumbers;
    }

    if (Object.keys(page).length > 0) opts.page = page;
  }

  // Columns
  const cols = findChild(el, "w:cols");
  if (cols) {
    const column: Record<string, unknown> = {};
    const count = attrNum(cols, "w:num");
    if (count !== undefined) column.count = count;
    const space = attrNum(cols, "w:space");
    if (space !== undefined) column.space = space;
    const sep = attrBool(cols, "w:sep");
    if (sep) column.separator = true;
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
    // Global page border attributes
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

  // Headers/footers - parse from references and store in a separate field
  const headerRefs: Record<string, unknown> = {};
  const footerRefs: Record<string, unknown> = {};

  for (const child of el.elements ?? []) {
    if (child.name === "w:headerReference") {
      const rId = attr(child, "r:id");
      const type = attr(child, "w:type");
      if (rId && type) {
        const headerChildren = parseHeaderFooterRef(rId, ctx);
        if (headerChildren) headerRefs[type] = headerChildren;
      }
    }
    if (child.name === "w:footerReference") {
      const rId = attr(child, "r:id");
      const type = attr(child, "w:type");
      if (rId && type) {
        const footerChildren = parseHeaderFooterRef(rId, ctx);
        if (footerChildren) footerRefs[type] = footerChildren;
      }
    }
  }

  if (Object.keys(headerRefs).length > 0) {
    (opts as Record<string, unknown>).parsedHeaders = headerRefs;
  }
  if (Object.keys(footerRefs).length > 0) {
    (opts as Record<string, unknown>).parsedFooters = footerRefs;
  }

  return opts as ISectionPropertiesOptions;
}

/**
 * Parse a header/footer reference by following the relationship to its XML part.
 */
function parseHeaderFooterRef(rId: string, ctx: ParseContext): SectionChild[] | undefined {
  const path = ctx.docx.partRefs.headers.get(rId) ?? ctx.docx.partRefs.footers.get(rId);
  if (!path) return undefined;

  const partEl = ctx.docx.doc.get(path);
  if (!partEl) return undefined;

  // The header/footer XML root element contains w:p, w:tbl, etc.
  const children: SectionChild[] = [];
  for (const child of partEl.elements ?? []) {
    const sectionChild = parseSectionChild(child, ctx);
    if (sectionChild !== undefined) {
      children.push(sectionChild);
    }
  }

  return children.length > 0 ? children : undefined;
}

// ── Section child dispatch ───────────────────────────────────────────────────

/**
 * Parse a single body child element into a SectionChild.
 */
export function parseSectionChild(el: Element, ctx: ParseContext): SectionChild {
  switch (el.name) {
    case "w:p": {
      // Check for textbox (w:pict containing v:textbox)
      const pict = findChild(el, "w:pict");
      if (pict) {
        const textbox = findDeepElement(pict, "v:textbox");
        if (textbox) {
          const textboxOpts = parseTextbox(pict, ctx);
          return { textbox: textboxOpts as SectionChild extends { textbox: infer T } ? T : never };
        }
      }

      return { paragraph: parseParagraph(el, ctx) };
    }
    case "w:tbl":
      return { table: parseTable(el, ctx) };
    case "w:sdt": {
      // Try TOC first
      const tocResult = parseToc(el, ctx);
      if (tocResult) {
        return { toc: tocResult };
      }
      // Otherwise parse as generic SDT block
      const sdtResult = parseSdtBlock(el, ctx);
      return {
        sdt: {
          properties: sdtResult.properties,
          children: sdtResult.children as SectionChild[] | undefined,
        },
      };
    }
    case "w:altChunk":
      return { altChunk: parseAltChunk(el, ctx) };
    case "w:subDoc":
      return { subDoc: parseSubDoc(el, ctx) };
    default:
      return new RawPassthrough(el);
  }
}

/**
 * Find a deep descendant element by name.
 */
function findDeepElement(parent: Element, name: string): Element | undefined {
  for (const child of parent.elements ?? []) {
    if (child.name === name) return child;
    const found = findDeepElement(child, name);
    if (found) return found;
  }
  return undefined;
}

// ── Body parsing with section splitting ───────────────────────────────────────

/**
 * Parse w:body element into SectionOptions[].
 *
 * Splits body content at w:sectPr boundaries to create sections.
 * The last w:sectPr (child of w:body directly) defines the last section.
 * Previous w:sectPr elements appear inside w:pPr elements.
 */
export function parseBody(body: Element, ctx: ParseContext): SectionOptions[] {
  // Wire up circular dependencies
  setSectionChildParser(parseSectionChild);
  setSectionChildrenParser(parseSectionChildrenElements);
  setTextboxSectionChildrenParser(parseSectionChildrenElements);

  // Collect body children and detect section breaks
  interface SectionBoundary {
    index: number;
    sectPr: Element;
  }

  const bodyChildren: Element[] = [];
  const boundaries: SectionBoundary[] = [];

  for (const child of body.elements ?? []) {
    if (child.name === "w:sectPr") {
      // Final section properties (last section)
      boundaries.push({ index: bodyChildren.length, sectPr: child });
    } else {
      bodyChildren.push(child);

      // Check for inline sectPr in paragraph properties
      if (child.name === "w:p") {
        const pPr = findChild(child, "w:pPr");
        if (pPr) {
          const sectPr = findChild(pPr, "w:sectPr");
          if (sectPr) {
            boundaries.push({ index: bodyChildren.length, sectPr });
          }
        }
      }
    }
  }

  // If no boundaries, the whole body is one section
  if (boundaries.length === 0) {
    return [
      {
        children: bodyChildren.map((el) => parseSectionChild(el, ctx)),
      },
    ];
  }

  // Split into sections
  const sections: SectionOptions[] = [];
  let start = 0;

  for (let i = 0; i < boundaries.length; i++) {
    const boundary = boundaries[i];
    // For inline sectPr (inside w:pPr), the containing paragraph was pushed to
    // bodyChildren. Exclude it — it's a section break marker, not content.
    // The last boundary uses a body-level sectPr, so no paragraph to exclude.
    const isInlineSectPr = i < boundaries.length - 1;
    const endIdx = isInlineSectPr ? Math.max(start, boundary.index - 1) : boundary.index;
    const sectionElements = bodyChildren.slice(start, endIdx);
    const parsedProps = parseSectionProperties(boundary.sectPr, ctx);
    const rawProps = parsedProps as Record<string, unknown>;

    // Extract headers/footers that were stored as parsedHeaders/parsedFooters
    const parsedHeaders = rawProps.parsedHeaders as Record<string, SectionChild[]> | undefined;
    const parsedFooters = rawProps.parsedFooters as Record<string, SectionChild[]> | undefined;

    // Build clean properties without internal fields
    const cleanProps = { ...parsedProps };
    delete (cleanProps as Record<string, unknown>).parsedHeaders;
    delete (cleanProps as Record<string, unknown>).parsedFooters;

    const section = {
      children: sectionElements.map((el) => parseSectionChild(el, ctx)),
      properties: cleanProps,
      ...(parsedHeaders ? { headers: parsedHeaders } : {}),
      ...(parsedFooters ? { footers: parsedFooters } : {}),
    } as SectionOptions;

    sections.push(section);
    start = boundary.index;
  }

  // If there are elements after the last boundary, they form the last section
  // with the body-level w:sectPr (already captured)
  // Actually the body-level sectPr IS the last boundary

  return sections;
}

/**
 * Parse a list of elements into SectionChild[].
 * Used by SDT and textbox parsers for their content.
 */
function parseSectionChildrenElements(elements: Element[], ctx: ParseContext): SectionChild[] {
  return elements.map((el) => parseSectionChild(el, ctx));
}
