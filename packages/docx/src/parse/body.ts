/**
 * Body parser for DOCX documents.
 *
 * Parses w:body → SectionOptions[] by splitting at w:sectPr boundaries.
 *
 * @module
 */
import { attr, findChild, findDeep, findFirst, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { parseAltChunk } from "@parts/alt-chunk/alt-chunk-parse";
import { parseCustomXmlBlock } from "@parts/custom-xml/custom-xml-parse";
import { parseSectionPropertiesEl } from "@parts/document/body/section-properties/descriptor";
import type { SectionPropertiesOptions } from "@parts/document/body/section-properties/section-properties";
import type { MarkupRangeOptions, BookmarkStartOptions } from "@parts/paragraph/links/bookmark";
import { parseSdtBlock } from "@parts/sdt/sdt-parse";
import { parseSubDoc } from "@parts/sub-doc/sub-doc-parse";
import {
  parseToc,
  parseTocFieldFromElements,
  selectTocEntryElements,
} from "@parts/table-of-contents/toc-parse";
import { tableDesc } from "@parts/table/descriptor";
import type { TableOptions } from "@parts/table/table";
import { parseTextbox } from "@parts/textbox/textbox-parse";
import type { SectionOptions } from "@shared/section";
import type { SectionChild } from "@shared/section";

import { parseParagraph } from "../body";
import { DocxReadContext } from "../context";
import { setBodyParseChild } from "../parts";
import { stringifyElement } from "../util/stringify-element";

// ── Section properties parser ────────────────────────────────────────────────

/** Internal parse result: section properties with extracted header/footer refs. */
type ParsedSectionProperties = SectionPropertiesOptions & {
  parsedHeaders?: Record<string, SectionChild[]>;
  parsedFooters?: Record<string, SectionChild[]>;
  parsedHeaderPartNames?: NonNullable<NonNullable<SectionOptions["headers"]>["partNames"]>;
  parsedFooterPartNames?: NonNullable<NonNullable<SectionOptions["footers"]>["partNames"]>;
};

/**
 * Parse w:sectPr element into SectionPropertiesOptions.
 * Delegates to the section properties descriptor's parse method.
 */
function parseSectionProperties(el: Element, ctx: DocxReadContext): ParsedSectionProperties {
  const opts: ParsedSectionProperties = parseSectionPropertiesEl(el);

  // Headers/footers - parse from references and store in a separate field
  const headerRefs: Record<string, SectionChild[]> = {};
  const footerRefs: Record<string, SectionChild[]> = {};
  const headerPartNames: NonNullable<ParsedSectionProperties["parsedHeaderPartNames"]> = {};
  const footerPartNames: NonNullable<ParsedSectionProperties["parsedFooterPartNames"]> = {};

  for (const child of el.elements ?? []) {
    if (child.name === "w:headerReference" || child.name === "w:footerReference") {
      const rId = attr(child, "r:id");
      const type = attr(child, "w:type");
      if (!rId || !type) continue;
      const slot = type as "default" | "first" | "even";
      const parsed = parseHeaderFooterRef(rId, ctx);
      if (!parsed) continue;
      if (child.name === "w:headerReference") {
        headerRefs[slot] = parsed.children;
        headerPartNames[slot] = parsed.partName;
      } else {
        footerRefs[slot] = parsed.children;
        footerPartNames[slot] = parsed.partName;
      }
    }
  }

  if (Object.keys(headerRefs).length > 0) {
    opts.parsedHeaders = headerRefs;
    opts.parsedHeaderPartNames = headerPartNames;
  }
  if (Object.keys(footerRefs).length > 0) {
    opts.parsedFooters = footerRefs;
    opts.parsedFooterPartNames = footerPartNames;
  }

  return opts;
}

/**
 * Parse a header/footer reference by following the relationship to its XML part.
 * Returns the parsed children plus the source part file name (headerN.xml) so
 * generate can pin the part numbering to the source on round-trip.
 */
function parseHeaderFooterRef(
  rId: string,
  ctx: DocxReadContext,
): { children: SectionChild[]; partName: string } | undefined {
  const path = ctx.docx.partRefs.headers.get(rId) ?? ctx.docx.partRefs.footers.get(rId);
  if (!path) return undefined;

  const partEl = ctx.docx.doc.get(path);
  if (!partEl) return undefined;

  // The header/footer XML root element contains w:p, w:tbl, etc. Parse under
  // the part's own relationship scope so its drawings resolve images correctly.
  const children: SectionChild[] = [];
  ctx.withPart(path, () => {
    for (const child of partEl.elements ?? []) {
      const sectionChild = parseSectionChild(child, ctx);
      if (sectionChild !== undefined) {
        children.push(sectionChild);
      }
    }
  });

  if (children.length === 0) return undefined;
  const partName = path.split("/").pop() ?? path;
  return { children, partName };
}

// ── Section child dispatch ───────────────────────────────────────────────────

/**
 * Parse a single body child element into a SectionChild.
 */
export function parseSectionChild(el: Element, ctx: DocxReadContext): SectionChild {
  switch (el.name) {
    case "w:p": {
      // Check for textbox (w:pict containing v:textbox)
      const pict = findChild(el, "w:pict");
      if (pict) {
        const textbox = findFirst(pict, "v:textbox");
        if (textbox) {
          const textboxOpts = parseTextbox(pict, ctx, parseSectionChildrenElements);
          return { textbox: textboxOpts as SectionChild extends { textbox: infer T } ? T : never };
        }
      }

      // Cross-paragraph complex field: a paragraph whose fldChar begin/end
      // markers are unbalanced opens (or continues) a field spanning
      // paragraph boundaries, and one carrying instrText with no fldChar at
      // all is a mid-field continuation. The per-paragraph field accumulator
      // cannot represent either shape — keep the whole paragraph verbatim
      // instead of dropping its runs.
      if (
        isCrossParagraphFieldStart(el) ||
        isFieldContinuation(el) ||
        isFieldSeparatorContinuation(el)
      ) {
        return { rawXml: stringifyElement(el) };
      }

      return { paragraph: parseParagraph(el, ctx) };
    }
    case "w:tbl":
      return { table: tableDesc.parse(el, ctx) as TableOptions };
    case "w:sdt": {
      // Try TOC first
      const tocResult = parseToc(el, ctx, parseSectionChildrenElements);
      if (tocResult) {
        return { toc: tocResult };
      }
      // Otherwise parse as generic SDT block
      const sdtResult = parseSdtBlock(el, ctx, parseSectionChildrenElements);
      return {
        sdt: {
          properties: sdtResult.properties,
          endProperties: sdtResult.endProperties,
          children: sdtResult.children as SectionChild[] | undefined,
        },
      };
    }
    case "w:altChunk":
      return { altChunk: parseAltChunk(el, ctx) };
    case "w:subDoc":
      return { subDoc: parseSubDoc(el, ctx) };
    case "w:customXml":
      return { customXml: parseCustomXmlBlock(el, ctx, parseSectionChild) };
    case "w:bookmarkStart": {
      // Body-level range markers sitting between paragraphs (e.g. _Toc bookmark
      // ends grouped after a heading). Carry them as first-class children so
      // they round-trip even though they are not wrapped in a paragraph.
      const idRaw = attr(el, "w:id");
      const name = attr(el, "w:name");
      if (idRaw !== undefined && name) {
        const bookmarkStart: Partial<BookmarkStartOptions> = { id: Number(idRaw), name };
        const disp = attr(el, "w:displacedByCustomXml");
        if (disp === "before" || disp === "after") bookmarkStart.displacedByCustomXml = disp;
        const colFirstRaw = attr(el, "w:colFirst");
        if (colFirstRaw !== undefined) bookmarkStart.colFirst = Number(colFirstRaw);
        const colLastRaw = attr(el, "w:colLast");
        if (colLastRaw !== undefined) bookmarkStart.colLast = Number(colLastRaw);
        return { bookmarkStart: bookmarkStart as BookmarkStartOptions };
      }
      return { rawXml: stringifyElement(el) };
    }
    case "w:bookmarkEnd": {
      const idRaw = attr(el, "w:id");
      if (idRaw !== undefined) {
        const bookmarkEnd: Partial<MarkupRangeOptions> = { id: Number(idRaw) };
        const disp = attr(el, "w:displacedByCustomXml");
        if (disp === "before" || disp === "after") bookmarkEnd.displacedByCustomXml = disp;
        return { bookmarkEnd: bookmarkEnd as MarkupRangeOptions };
      }
      return { rawXml: stringifyElement(el) };
    }
    default:
      return { rawXml: stringifyElement(el) };
  }
}

// ── Body parsing with section splitting ───────────────────────────────────────

/**
 * Parse w:body element into SectionOptions[].
 *
 * Splits body content at w:sectPr boundaries to create sections.
 * The last w:sectPr (child of w:body directly) defines the last section.
 * Previous w:sectPr elements appear inside w:pPr elements.
 */
export function parseBody(body: Element, ctx: DocxReadContext): SectionOptions[] {
  // Register the body child parser for descriptor parse callbacks
  setBodyParseChild(parseSectionChild);

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
        children: parseBodyChildren(bodyChildren, ctx),
      },
    ];
  }

  // Split into sections
  const sections: SectionOptions[] = [];
  let start = 0;

  for (const boundary of boundaries) {
    // A sectPr inside a paragraph's pPr marks that paragraph as the final
    // content paragraph of its section. Its runs/drawings ARE section content
    // (e.g. an inline image), so include it in the slice — the paragraph parser
    // ignores pPr/w:sectPr, and stringify re-injects the sectPr into this same
    // paragraph's pPr. The last boundary is a body-level sectPr (never in a
    // paragraph), so boundary.index already points past every real child.
    const endIdx = boundary.index;
    const sectionElements = bodyChildren.slice(start, endIdx);
    const parsedProps = parseSectionProperties(boundary.sectPr, ctx);

    // Extract headers/footers that were stored as parsedHeaders/parsedFooters
    const { parsedHeaders, parsedFooters, parsedHeaderPartNames, parsedFooterPartNames } =
      parsedProps;

    // Build clean properties without internal fields
    const cleanProps = { ...parsedProps };
    delete cleanProps.parsedHeaders;
    delete cleanProps.parsedFooters;
    delete cleanProps.parsedHeaderPartNames;
    delete cleanProps.parsedFooterPartNames;

    const section = {
      children: parseBodyChildren(sectionElements, ctx),
      properties: cleanProps,
      ...(parsedHeaders
        ? {
            headers: {
              ...parsedHeaders,
              ...(parsedHeaderPartNames ? { partNames: parsedHeaderPartNames } : {}),
            },
          }
        : {}),
      ...(parsedFooters
        ? {
            footers: {
              ...parsedFooters,
              ...(parsedFooterPartNames ? { partNames: parsedFooterPartNames } : {}),
            },
          }
        : {}),
    } as SectionOptions;

    sections.push(section);
    start = boundary.index;
  }

  // If there are elements after the last boundary, they form the last section
  // with the body-level w:sectPr (already captured)
  // Actually the body-level sectPr IS the last boundary

  return sections;
}

// ── Cross-paragraph TOC field aggregation ───────────────────────────────────

/**
 * Net field-nesting change across all descendant fldChar markers
 * (begin: +1, end: -1). Balances cross-paragraph field boundaries without a
 * stack — the running depth hits 0 exactly when the outermost field closes.
 */
function countFieldDelta(el: Element): number {
  let delta = 0;
  const walk = (node: Element): void => {
    if (node.name === "w:fldChar") {
      const type = attr(node, "w:fldCharType");
      if (type === "begin") delta += 1;
      else if (type === "end") delta -= 1;
    }
    for (const c of node.elements ?? []) {
      if (c.type === "element") walk(c);
    }
  };
  walk(el);
  return delta;
}

/**
 * True when a w:p opens a bare TOC complex field: it carries a fldChar begin
 * whose instrText starts with "TOC". Such fields span multiple paragraphs and
 * defeat the per-paragraph field accumulator, so they are aggregated as rawXml.
 */
function isTocFieldBegin(el: Element): boolean {
  if (el.name !== "w:p") return false;
  let hasBegin = false;
  let instr = "";
  const walk = (node: Element): void => {
    if (node.name === "w:fldChar" && attr(node, "w:fldCharType") === "begin") hasBegin = true;
    if (node.name === "w:instrText") instr += textOf(node);
    for (const c of node.elements ?? []) {
      if (c.type === "element") walk(c);
    }
  };
  walk(el);
  return hasBegin && instr.trim().toUpperCase().startsWith("TOC");
}

/**
 * True when a w:p's fldChar markers don't balance out (net opens or closes)
 * — a complex field crossing the paragraph boundary that the per-paragraph
 * accumulator would silently truncate.
 */
function isCrossParagraphFieldStart(el: Element): boolean {
  return countFieldDelta(el) !== 0;
}

/**
 * True when a w:p carries instrText but no fldChar at all — a mid-field
 * continuation paragraph between a distant begin and end.
 */
function isFieldContinuation(el: Element): boolean {
  let hasInstr = false;
  let hasFldChar = false;
  const walk = (node: Element): void => {
    if (node.name === "w:instrText") hasInstr = true;
    else if (node.name === "w:fldChar") hasFldChar = true;
    for (const c of node.elements ?? []) {
      if (c.type === "element") walk(c);
    }
  };
  walk(el);
  return hasInstr && !hasFldChar;
}

/**
 * True when a w:p carries a separate fldChar but no begin — the separator of a
 * field whose begin sits in an earlier paragraph. The per-paragraph
 * accumulator would silently consume the separate run.
 */
function isFieldSeparatorContinuation(el: Element): boolean {
  let hasBegin = false;
  let hasSeparate = false;
  const walk = (node: Element): void => {
    if (node.name === "w:fldChar") {
      const type = attr(node, "w:fldCharType");
      if (type === "begin") hasBegin = true;
      else if (type === "separate") hasSeparate = true;
    }
    for (const c of node.elements ?? []) {
      if (c.type === "element") walk(c);
    }
  };
  walk(el);
  return hasSeparate && !hasBegin;
}

/**
 * Parse a run of body-level elements into SectionChild[], aggregating any
 * cross-paragraph TOC complex field into a single rawXml child so its nested
 * HYPERLINK/PAGEREF fields and bookmark markers round-trip intact.
 */
function parseBodyChildren(elements: Element[], ctx: DocxReadContext): SectionChild[] {
  const children: SectionChild[] = [];
  let tocBuffer: Element[] | null = null;
  let tocDepth = 0;

  const flushToc = (): void => {
    if (!tocBuffer) return;
    children.push(buildTocChild(tocBuffer, ctx));
    // buildTocChild preserves the rendered entries (paragraphs between the
    // separate and end markers) but not the end-closing paragraph, which often
    // carries a trailing page break (the section break before the first
    // heading). Rescue that page break as a standalone child to avoid silently
    // dropping it on round-trip.
    const lastEl = tocBuffer[tocBuffer.length - 1];
    let pageBreakCount = 0;
    for (const br of findDeep(lastEl, "w:br")) {
      if (attr(br, "w:type") === "page") pageBreakCount++;
    }
    for (let i = 0; i < pageBreakCount; i++) {
      children.push({ paragraph: { children: [{ pageBreak: true }] } });
    }
    tocBuffer = null;
    tocDepth = 0;
  };

  for (const el of elements) {
    if (tocBuffer !== null) {
      tocBuffer.push(el);
      tocDepth += countFieldDelta(el);
      if (tocDepth <= 0) flushToc();
      continue;
    }
    if (isTocFieldBegin(el)) {
      tocBuffer = [el];
      tocDepth = countFieldDelta(el);
      if (tocDepth <= 0) flushToc();
      continue;
    }
    children.push(parseSectionChild(el, ctx));
  }

  // Unclosed TOC field at end of content — flush what we have (best effort).
  flushToc();

  return children;
}

/**
 * Build a structured TOC SectionChild from a captured bare TOC field. Extracts
 * the field instruction (switches → TableOfContentsOptions) and preserves the
 * rendered entries (separate→end paragraphs) structurally so MS Office and WPS
 * both display the existing TOC. The field is emitted clean (no dirty flag).
 */
function buildTocChild(els: Element[], ctx: DocxReadContext): SectionChild {
  const tocOpts = parseTocFieldFromElements(els);
  const entryEls = selectTocEntryElements(els);
  if (entryEls.length > 0) {
    tocOpts.entries = entryEls.map((el) => parseSectionChild(el, ctx));
  }
  return { toc: tocOpts };
}

/**
 * Parse a list of elements into SectionChild[].
 * Used by SDT and textbox parsers for their content.
 */
function parseSectionChildrenElements(elements: Element[], ctx: DocxReadContext): SectionChild[] {
  return parseBodyChildren(elements, ctx);
}
