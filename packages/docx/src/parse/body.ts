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
import { parseSdtBlock } from "@parts/sdt/sdt-parse";
import type { TableOfContentsOptions } from "@parts/table-of-contents/table-of-contents-properties";
import {
  parseToc,
  parseTocFieldFromElements,
  selectTocEntryElements,
} from "@parts/table-of-contents/toc-parse";
import { tableDesc } from "@parts/table/descriptor";
import type { TableOptions } from "@parts/table/table";
import { parseTextbox } from "@parts/textbox/textbox-parse";
import type { SectionOptions } from "@shared/section";
import type { BlockContentChild, SectionChild } from "@shared/section";

import {
  parseBookmarkEndOptions,
  parseBookmarkStartOptions,
  parseParagraph,
  runRPrXml,
} from "../body";
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
      if (child.type !== "element") continue;
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
      // Try TOC first. Entry paragraphs are parsed directly (not through the
      // cross-paragraph TOC aggregator): the emitted field chain shares its
      // head/end runs with the first/last entry paragraphs, so the aggregator
      // would see a begin+instrText inside the first entry and fold the whole
      // span into a nested TOC — re-parsing output would nest one sdt level
      // deeper on every round-trip.
      const parseTocEntries = (els: Element[], entryCtx: DocxReadContext): SectionChild[] =>
        els.map((entryEl) =>
          entryEl.name === "w:p"
            ? { paragraph: parseParagraph(entryEl, entryCtx) }
            : parseSectionChild(entryEl, entryCtx),
        );
      const tocResult = parseToc(el, ctx, parseTocEntries);
      if (tocResult) {
        return { toc: tocResult };
      }
      // Otherwise parse as generic SDT block
      const sdtResult = parseSdtBlock(el, ctx, parseSectionChildrenElements);
      return {
        sdt: {
          properties: sdtResult.properties,
          endProperties: sdtResult.endProperties,
          children: sdtResult.children as BlockContentChild[] | undefined,
        },
      };
    }
    case "w:altChunk":
      return { altChunk: parseAltChunk(el, ctx) };
    case "w:customXml":
      return { customXml: parseCustomXmlBlock(el, ctx, parseSectionChild) };
    case "w:bookmarkStart": {
      // Body-level range markers sitting between paragraphs (e.g. _Toc bookmark
      // ends grouped after a heading). Carry them as first-class children so
      // they round-trip even though they are not wrapped in a paragraph.
      const bookmarkStart = parseBookmarkStartOptions(el);
      return bookmarkStart ? { bookmarkStart } : { rawXml: stringifyElement(el) };
    }
    case "w:bookmarkEnd": {
      const bookmarkEnd = parseBookmarkEndOptions(el);
      return bookmarkEnd ? { bookmarkEnd } : { rawXml: stringifyElement(el) };
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
  // A self-contained field paragraph (begin→end balanced inside one w:p, e.g.
  // pandoc's empty caption TOCs) round-trips through the per-paragraph field
  // accumulator as a single { field } child — aggregating it here would emit
  // it twice (an entry-less TOC plus the paragraph itself).
  if (countFieldDelta(el) === 0) return false;
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
    // The closing paragraph either joins the entries (pure control paragraph —
    // buildTocChild kept it), returns to the body (it carries real content:
    // Word drops the field end into a following heading when it updates the
    // TOC, and that heading must not move inside the TOC), or is dropped (bare
    // control paragraph) — only then does its trailing page break need rescue
    // as a standalone child. An unclosed field's tail keeps the legacy rescue.
    const lastEl = tocBuffer[tocBuffer.length - 1]!;
    const closed = tocDepth <= 0;
    // Same rule selectTocEntryElements applied inside buildTocChild: the
    // closing paragraph is kept iff it ends up as the last collected entry.
    const entryEls = selectTocEntryElements(tocBuffer);
    const keptInEntries = closed && entryEls[entryEls.length - 1] === lastEl;
    const returnsToBody = closed && !keptInEntries && findFirst(lastEl, "w:t") !== undefined;
    if (returnsToBody) {
      // Direct parseParagraph: the paragraph lives inside the field span, its
      // orphan end fldChar run is consumed by the accumulator, and the
      // cross-paragraph pre-check must not flip it to verbatim.
      children.push(
        lastEl.name === "w:p"
          ? { paragraph: parseParagraph(lastEl, ctx) }
          : parseSectionChild(lastEl, ctx),
      );
    } else if (!keptInEntries) {
      let pageBreakCount = 0;
      for (const br of findDeep(lastEl, "w:br")) {
        if (attr(br, "w:type") === "page") pageBreakCount++;
      }
      for (let i = 0; i < pageBreakCount; i++) {
        children.push({ paragraph: { children: [{ pageBreak: true }] } });
      }
    }
    tocBuffer = null;
    tocDepth = 0;
  };

  for (const el of elements) {
    // Whitespace text nodes survive captureSpacesBetweenElements parsing;
    // block content is element-only, so skip anything without an element name.
    if (el.type !== "element") continue;
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
  // The aggregated field span carries no w:sdt wrapper — keep it bare so the
  // re-emitted TOC does not grow a content control the source never had.
  tocOpts.bare = true;
  // The field end drifted into a following body paragraph (carrying text) —
  // that paragraph round-trips through the returnsToBody path, so the emit
  // path must not inject a second end run into the last entry.
  const lastEl = els[els.length - 1]!;
  if (lastEl && countFieldDelta(lastEl) < 0 && findFirst(lastEl, "w:t") !== undefined) {
    tocOpts.endInBody = true;
  }
  const entryEls = selectTocEntryElements(els);
  if (entryEls.length > 0) {
    // Entry paragraphs live INSIDE the field: the per-paragraph accumulator
    // consumes their fldChar control runs and keeps the rest. They must not go
    // through parseSectionChild's cross-paragraph pre-check — the closing
    // paragraph's unbalanced `end` would flip it to verbatim rawXml and lose
    // the structured pPr (pStyle/divId) this capture exists to preserve.
    tocOpts.entries = entryEls.map((el) =>
      el.name === "w:p" ? { paragraph: parseParagraph(el, ctx) } : parseSectionChild(el, ctx),
    );
  }
  captureTocFieldRPr(els, tocOpts);
  return { toc: tocOpts };
}

/**
 * Capture the outer TOC field's control-run rPr (begin run stands in for the
 * begin/instr/separate controls) and the closing end run's rPr when it differs.
 * The re-emitted field chain carries them so Word's explicit style overrides on
 * these invisible runs round-trip.
 */
export function captureTocFieldRPr(els: Element[], tocOpts: TableOfContentsOptions): void {
  let depth = 0;
  let controlRPr: string | undefined;
  let closed = false;
  // Word splits the field instruction across runs (leading space / text /
  // trailing space, each with its own rPr). Keep the begin→separate control
  // chain verbatim so the run split round-trips instead of collapsing to a
  // single re-composed instruction run.
  let headOpen = false;
  const headRuns: string[] = [];
  const walk = (node: Element): void => {
    if (closed || node.name !== "w:r" || !node.elements) {
      for (const c of node.elements ?? []) {
        if (c.type === "element") walk(c);
      }
      return;
    }
    const fldChar = node.elements.find((c) => c.name === "w:fldChar");
    if (fldChar) {
      const type = attr(fldChar, "w:fldCharType");
      if (type === "begin") {
        depth++;
        if (depth === 1) {
          controlRPr = runRPrXml(node);
          headOpen = true;
          headRuns.push(stringifyElement(node));
          return;
        }
      } else if (type === "separate" && headOpen) {
        headRuns.push(stringifyElement(node));
        tocOpts.headRunsXml = headRuns.join("");
        headOpen = false;
        return;
      } else if (type === "end") {
        depth--;
        if (depth === 0) {
          // endRPrXml distinguishes the three end-run shapes: "" = no rPr
          // (must stay bare, not inherit the control rPr), undefined = same
          // as the control rPr (emit falls back to it), otherwise its own.
          const endRPr = runRPrXml(node);
          if (controlRPr) tocOpts.rPrXml = controlRPr;
          if (endRPr !== controlRPr) tocOpts.endRPrXml = endRPr ?? "";
          closed = true;
          return;
        }
      }
    }
    if (headOpen) {
      headRuns.push(stringifyElement(node));
      return;
    }
    for (const c of node.elements ?? []) {
      if (c.type === "element") walk(c);
    }
  };
  for (const el of els) {
    walk(el);
    if (closed) break;
  }
}

/**
 * Parse a list of elements into SectionChild[].
 * Used by SDT and textbox parsers for their content.
 */
function parseSectionChildrenElements(elements: Element[], ctx: DocxReadContext): SectionChild[] {
  return parseBodyChildren(elements, ctx);
}
