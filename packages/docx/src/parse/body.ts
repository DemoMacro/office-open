/**
 * Body parser for DOCX documents.
 *
 * Parses w:body → SectionOptions[] by splitting at w:sectPr boundaries.
 *
 * @module
 */
import { attr, findChild, findDeep, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { parseAltChunk } from "@parts/alt-chunk/alt-chunk-parse";
import { parseCustomXmlBlock } from "@parts/custom-xml/custom-xml-parse";
import { parseSectionPropertiesEl } from "@parts/document/body/section-properties/descriptor";
import type { SectionPropertiesOptions } from "@parts/document/body/section-properties/section-properties";
import { parseSdtBlock } from "@parts/sdt/sdt-parse";
import { parseSubDoc } from "@parts/sub-doc/sub-doc-parse";
import { parseToc, parseTocFieldFromElements } from "@parts/table-of-contents/toc-parse";
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
    opts.parsedHeaders = headerRefs;
  }
  if (Object.keys(footerRefs).length > 0) {
    opts.parsedFooters = footerRefs;
  }

  return opts;
}

/**
 * Parse a header/footer reference by following the relationship to its XML part.
 */
function parseHeaderFooterRef(rId: string, ctx: DocxReadContext): SectionChild[] | undefined {
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

  return children.length > 0 ? children : undefined;
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
        const textbox = findDeepElement(pict, "v:textbox");
        if (textbox) {
          const textboxOpts = parseTextbox(pict, ctx, parseSectionChildrenElements);
          return { textbox: textboxOpts as SectionChild extends { textbox: infer T } ? T : never };
        }
      }

      return { paragraph: parseParagraph(el, ctx) };
    }
    case "w:tbl":
      return { table: tableDesc.parse(el, ctx) as TableOptions };
    case "w:sdt": {
      // Try TOC first
      const tocResult = parseToc(el, ctx);
      if (tocResult) {
        return { toc: tocResult };
      }
      // Otherwise parse as generic SDT block
      const sdtResult = parseSdtBlock(el, ctx, parseSectionChildrenElements);
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
    case "w:customXml":
      return { customXml: parseCustomXmlBlock(el, ctx, parseSectionChild) };
    case "w:bookmarkStart": {
      // Body-level range markers sitting between paragraphs (e.g. _Toc bookmark
      // ends grouped after a heading). Carry them as first-class children so
      // they round-trip even though they are not wrapped in a paragraph.
      const idRaw = attr(el, "w:id");
      const name = attr(el, "w:name");
      if (idRaw !== undefined && name) {
        const bookmarkStart: {
          id: number;
          name: string;
          displacedByCustomXml?: "before" | "after";
        } = {
          id: Number(idRaw),
          name,
        };
        const disp = attr(el, "w:displacedByCustomXml");
        if (disp === "before" || disp === "after") bookmarkStart.displacedByCustomXml = disp;
        return { bookmarkStart };
      }
      return { rawXml: stringifyElement(el) };
    }
    case "w:bookmarkEnd": {
      const idRaw = attr(el, "w:id");
      if (idRaw !== undefined) {
        const bookmarkEnd: { id: number; displacedByCustomXml?: "before" | "after" } = {
          id: Number(idRaw),
        };
        const disp = attr(el, "w:displacedByCustomXml");
        if (disp === "before" || disp === "after") bookmarkEnd.displacedByCustomXml = disp;
        return { bookmarkEnd };
      }
      return { rawXml: stringifyElement(el) };
    }
    default:
      return { rawXml: stringifyElement(el) };
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

  for (let i = 0; i < boundaries.length; i++) {
    const boundary = boundaries[i];
    // For inline sectPr (inside w:pPr), the containing paragraph was pushed to
    // bodyChildren. Exclude it — it's a section break marker, not content.
    // The last boundary uses a body-level sectPr, so no paragraph to exclude.
    const isInlineSectPr = i < boundaries.length - 1;
    const endIdx = isInlineSectPr ? Math.max(start, boundary.index - 1) : boundary.index;
    const sectionElements = bodyChildren.slice(start, endIdx);
    const parsedProps = parseSectionProperties(boundary.sectPr, ctx);

    // Extract headers/footers that were stored as parsedHeaders/parsedFooters
    const { parsedHeaders, parsedFooters } = parsedProps;

    // Build clean properties without internal fields
    const cleanProps = { ...parsedProps };
    delete cleanProps.parsedHeaders;
    delete cleanProps.parsedFooters;

    const section = {
      children: parseBodyChildren(sectionElements, ctx),
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
    children.push(buildTocChild(tocBuffer));
    // The TOC field's closing paragraph often also carries a trailing page
    // break (the section break before the first heading). buildTocChild
    // discards the rendered result, so rescue that page break as a standalone
    // child to avoid silently dropping it on round-trip.
    const lastEl = tocBuffer[tocBuffer.length - 1];
    const pageBreakCount = findDeep(lastEl, "w:br").filter(
      (b) => attr(b, "w:type") === "page",
    ).length;
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
 * the field instruction (switches → TableOfContentsOptions) and discards the
 * rendered result (separate→end content) — Word regenerates it from headings on
 * open via the dirty flag emitted by stringifyTableOfContents.
 */
function buildTocChild(els: Element[]): SectionChild {
  return { toc: parseTocFieldFromElements(els) };
}

/**
 * Parse a list of elements into SectionChild[].
 * Used by SDT and textbox parsers for their content.
 */
function parseSectionChildrenElements(elements: Element[], ctx: DocxReadContext): SectionChild[] {
  return parseBodyChildren(elements, ctx);
}
