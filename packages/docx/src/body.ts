/**
 * Body-level stringification for DOCX documents.
 *
 * Converts pure JSON options to XML strings for document body content.
 * Pure string concatenation — no intermediate object tree.
 *
 * @module
 */

import type { UniversalMeasure } from "@office-open/core";
import { xsdJcAlignment } from "@office-open/core";
import { toUint8Array } from "@office-open/core";
import { hexColorValue, uCharHexNumber } from "@office-open/core";
import { stringifyVmlShape } from "@office-open/core";
import { stringifyVmlBackground } from "@office-open/core";
import { nextVmlShapeId } from "@office-open/core";
import type { VmlShapeStyle } from "@office-open/core";
import {
  attr,
  attrBool,
  attrMeasure,
  attrNum,
  escapeXml,
  findChild,
  textOf,
} from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { findFirst } from "@office-open/xml";
import { sectionPropertiesDesc } from "@parts/document/body/section-properties/descriptor";
import {
  documentNamespaceAttributesInDialect,
  documentNamespaceDialect,
} from "@parts/document/document-attributes";
import type {
  BackgroundRawMediaOptions,
  DocumentBackgroundOptions,
} from "@parts/document/document-background";
import { parseDrawingRun, parseUnsupportedDrawingRun } from "@parts/drawing/drawing-parse";
import { FontWrapper } from "@parts/fonts/font-wrapper";
import type { BordersOptions } from "@parts/paragraph/formatting/border";
import type { IndentProperties } from "@parts/paragraph/formatting/indent";
import { LineRuleType } from "@parts/paragraph/formatting/spacing";
import type { SpacingProperties } from "@parts/paragraph/formatting/spacing";
import { HeadingLevel } from "@parts/paragraph/formatting/style";
import type { TabStopDefinition } from "@parts/paragraph/formatting/tab-stop";
import type { FrameOptions } from "@parts/paragraph/frame/frame-properties";
import type {
  MarkupRangeOptions,
  BookmarkStartOptions,
  MoveRangeStartOptions,
} from "@parts/paragraph/links/bookmark";
import type { DisplacedByCustomXml } from "@parts/paragraph/links/bookmark";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";
import type {
  NumberingInsertionOptions,
  ParagraphPropertiesOptions,
  ParagraphPropertiesChangeOptions,
} from "@parts/paragraph/properties";
import { parseFormFieldData } from "@parts/paragraph/run/form-field";
import type { FormFieldOptions } from "@parts/paragraph/run/form-field";
import type { PositionalTabOptions } from "@parts/paragraph/run/positional-tab";
import type {
  ParagraphRunPropertiesOptions,
  RunPropertiesOptions,
} from "@parts/paragraph/run/properties";
import { ON_OFF_LEXICAL_VALUES } from "@parts/paragraph/run/properties";
import type { OnOffLexicalValue } from "@parts/paragraph/run/properties";
import type { BreakOptions, RunOptions } from "@parts/paragraph/run/run";
import { breakXml } from "@parts/paragraph/run/run";
import { parseRun, parseRunProperties, parsedRunToOptions } from "@parts/paragraph/run/run-parse";
import { parseSdtProperties } from "@parts/sdt/sdt-parse";
import { parseSubDoc } from "@parts/sub-doc/sub-doc-parse";
import { stringifyTableOfContents } from "@parts/table-of-contents/descriptor";
import type { TableOptions } from "@parts/table/table";
import { parseBorderSide } from "@shared/border";
import type { MediaData } from "@shared/media/data";
import type { SectionChild } from "@shared/section";
import { parseShading } from "@shared/shading";

import type { DocxReadContext, DocxWriteContext, BodyContext } from "./context";
import { DocxParseError } from "./errors";
import { tableDesc, altChunkDesc, sdtBlockDesc, customXmlBlockDesc, subDocDesc } from "./parts";
import { parseCustomXmlProperties } from "./parts/bodychildren";
import { stringifyChildDispatch, stringifyRunInline } from "./parts/inline";
import { parseMathChildren, parseMathParagraph } from "./parts/paragraph/math/stringify";
import type {
  ComplexFieldInstructionMember,
  ComplexFieldOptions,
  FieldMarkerOptions,
  ParagraphChild,
  SdtRunOptions,
  TrackChangeChild,
} from "./parts/paragraph/paragraph";
import type { NumberingChangeOptions } from "./parts/paragraph/properties";
import {
  EMPTY_PPR_RESULT,
  paragraphIdentityAttrs,
  stringifyParagraphProperties,
} from "./parts/paragraph/stringify";
import { replaceRelsWithPlaceholders } from "./util/replace-media-placeholders";
import { stringifyElement } from "./util/stringify-element";

export type { BodyContext } from "./context";

type TrackedCommentReferenceRun = {
  commentReference: number;
  properties?: RunPropertiesOptions;
  additionRsid?: RunOptions["additionRsid"];
  runPropertiesRsid?: RunOptions["runPropertiesRsid"];
  deletionRsid?: RunOptions["deletionRsid"];
};

function isTrackedCommentReferenceRun(value: object): value is TrackedCommentReferenceRun {
  return "commentReference" in value && typeof value.commentReference === "number";
}

// ── Run ──

/**
 * Stringify a run (w:r) from pure JSON options.
 *
 * Handles text, children, breaks, and run properties.
 */
export function stringifyRun(opts: RunOptions, ctx: BodyContext): string {
  return stringifyRunInline(opts, ctx);
}

// ── Paragraph ──

/**
 * Stringify a paragraph (w:p) from pure JSON options or string.
 *
 * Handles paragraph properties, numbering registration, and run children.
 */
export function stringifyParagraph(
  opts: string | ParagraphOptions,
  ctx: BodyContext,
  sectionPropertiesXml?: string,
): string {
  const isPlainText = typeof opts === "string";
  const resolved: ParagraphOptions = isPlainText ? { text: opts } : opts;
  if (
    !isPlainText &&
    (resolved.sectionHeaders !== undefined || resolved.sectionFooters !== undefined)
  ) {
    resolved.sectionProperties = {
      ...resolved.sectionProperties,
      ...ctx.fileData.registerInlineSection({
        headers: resolved.sectionHeaders,
        footers: resolved.sectionFooters,
        properties: resolved.sectionProperties,
      }),
    };
  }
  const embeddedSectionXml =
    !isPlainText && !sectionPropertiesXml && resolved.sectionProperties !== undefined
      ? (sectionPropertiesDesc.stringify(resolved.sectionProperties, ctx) ?? "")
      : "";
  const paragraphSectionXml = sectionPropertiesXml ?? embeddedSectionXml;
  let body = "";

  // Build paragraph properties — direct string output, no intermediate object tree.
  // A string paragraph is `{ text }` by construction: no pPr fields, no numbering.
  const props = isPlainText
    ? EMPTY_PPR_RESULT
    : stringifyParagraphProperties(resolved, ctx.reproducible);

  // Register numbering references (length check skips the iterator allocation
  // on the common no-numbering path)
  const numberingRefs = props.numberingReferences;
  if (numberingRefs.length > 0 && !(ctx.viewWrapper instanceof FontWrapper)) {
    for (const ref of numberingRefs) {
      ctx.file.numbering.createConcreteNumberingInstance(ref.reference, ref.instance);
    }
  }

  // Paragraph properties XML
  if (props.xml) {
    if (paragraphSectionXml) {
      // Insert sectPr before closing </w:pPr>; a self-closed empty pPr
      // (emptyProperties round-trip marker) expands to host the section break.
      const changeIndex = props.xml.indexOf("<w:pPrChange");
      body +=
        changeIndex >= 0
          ? props.xml.slice(0, changeIndex) + paragraphSectionXml + props.xml.slice(changeIndex)
          : props.xml.includes("</w:pPr>")
            ? props.xml.replace("</w:pPr>", paragraphSectionXml + "</w:pPr>")
            : props.xml.replace(/\/>$/, `>${paragraphSectionXml}</w:pPr>`);
    } else {
      body += props.xml;
    }
  } else if (paragraphSectionXml) {
    // No pPr but we need sectPr — wrap in pPr
    body += `<w:pPr>${paragraphSectionXml}</w:pPr>`;
  }

  // Text shorthand — the run is `{ text }` by construction (no rPr/break/children/rsid),
  // so build the bare <w:r> directly instead of walking stringifyRun's full dispatch.
  if (resolved.text !== undefined) {
    body += `<w:r>${textElementXml("w:t", String(resolved.text), resolved.preserveSpace)}</w:r>`;
  }

  // Children
  if (resolved.children) {
    for (const child of resolved.children) {
      if (typeof child === "string") {
        body += `<w:r>${textElementXml("w:t", child)}</w:r>`;
      } else if (typeof child === "object" && child !== null) {
        // Try JSON child dispatch first (image, chart, pageBreak, etc.)
        if ("table" in child) {
          body += tableDesc.stringify(child.table, ctx) ?? "";
          continue;
        }
        if ("blockBreak" in child) {
          body += breakXml(child.blockBreak);
          continue;
        }
        const jsonResult = stringifyChildDispatch(child as ParagraphChild, ctx);
        if (jsonResult !== undefined) {
          body += Array.isArray(jsonResult) ? jsonResult.join("") : jsonResult;
        } else {
          // RunOptions-like plain object — may be an empty run carrying only
          // run properties (round-tripped from <w:r><w:rPr>…</w:rPr></w:r>).
          body += stringifyRun(child as RunOptions, ctx);
        }
      }
    }
  }

  const attr = paragraphIdentityAttrs(resolved);
  return body ? `<w:p${attr}>${body}</w:p>` : `<w:p${attr}/>`;
}

// ── Body child dispatch ──

/**
 * Stringify a body-level child element.
 *
 * Dispatches to the appropriate stringifier based on the child type.
 * Pure JSON API — no class instance support.
 */
export function stringifyBodyChild(
  child: SectionChild,
  ctx: BodyContext,
  sectionPropertiesXml?: string,
): string {
  // Plain object dispatch — all via descriptors
  if ("paragraph" in child) {
    return stringifyParagraph(child.paragraph, ctx, sectionPropertiesXml);
  }
  if ("table" in child) {
    return tableDesc.stringify(child.table, ctx) ?? "";
  }
  if ("toc" in child) {
    const { alias, leading, trailing, ...options } = child.toc;
    const entries = options.entries ?? [];
    // A non-final section may host its sectPr in the TOC's closing entry
    // paragraph (the section-break paragraph the field end shares).
    const entriesXml = entries
      .map((entry, i) =>
        i === entries.length - 1
          ? stringifyBodyChild(entry, ctx, sectionPropertiesXml)
          : stringifyBodyChild(entry, ctx),
      )
      .join("");
    const contentAround = (nodes: SectionChild[] | undefined): string =>
      (nodes ?? []).map((node) => stringifyBodyChild(node, ctx)).join("");
    // A freshly generated TOC (no entries) gets the conventional alias; a
    // round-tripped one emits alias only when the source carried it.
    const isFresh = entries.length === 0 && leading === undefined && trailing === undefined;
    return stringifyTableOfContents(
      alias ?? (isFresh ? "Table of Contents" : undefined),
      options,
      entriesXml,
      contentAround(leading),
      contentAround(trailing),
    );
  }
  if ("textbox" in child) {
    return stringifyTextbox(child.textbox, ctx);
  }
  if ("sdt" in child) {
    return sdtBlockDesc.stringify(child.sdt, ctx) ?? "";
  }
  if ("altChunk" in child) {
    return altChunkDesc.stringify(child.altChunk, ctx) ?? "";
  }
  if ("subDoc" in child) {
    return subDocDesc.stringify(child.subDoc, ctx) ?? "";
  }
  if ("customXml" in child) {
    return customXmlBlockDesc.stringify(child.customXml, ctx) ?? "";
  }
  if ("bookmarkStart" in child) {
    const bs = child.bookmarkStart;
    const a: string[] = [`w:id="${bs.id}"`, `w:name="${escapeXml(bs.name)}"`];
    if (bs.displacedByCustomXml) a.push(`w:displacedByCustomXml="${bs.displacedByCustomXml}"`);
    if (bs.colFirst !== undefined) a.push(`w:colFirst="${bs.colFirst}"`);
    if (bs.colLast !== undefined) a.push(`w:colLast="${bs.colLast}"`);
    return `<w:bookmarkStart ${a.join(" ")}/>`;
  }
  if ("bookmarkEnd" in child) {
    const be = child.bookmarkEnd;
    const a: string[] = [`w:id="${be.id}"`];
    if (be.displacedByCustomXml) a.push(`w:displacedByCustomXml="${be.displacedByCustomXml}"`);
    return `<w:bookmarkEnd ${a.join(" ")}/>`;
  }
  if ("rawXml" in child) {
    return child.rawXml;
  }

  throw new Error("Unknown section child type");
}

// ── Document background (pure function, no XmlComponent) ──

function stringifyDocumentBackground(opts: DocumentBackgroundOptions, ctx: BodyContext): string {
  const attrs: string[] = [];
  if (opts.color !== undefined) attrs.push(`w:color="${hexColorValue(opts.color)}"`);
  if (opts.themeColor !== undefined) attrs.push(`w:themeColor="${opts.themeColor}"`);
  if (opts.themeShade !== undefined)
    attrs.push(`w:themeShade="${uCharHexNumber(opts.themeShade)}"`);
  if (opts.themeTint !== undefined) attrs.push(`w:themeTint="${uCharHexNumber(opts.themeTint)}"`);
  const attrStr = attrs.join(" ");

  if (opts.vmlBackground) {
    let vml = opts.vmlBackground;
    if (opts.rawMedia) {
      const fill = vml.fill ? { ...vml.fill } : undefined;
      for (const m of opts.rawMedia) {
        const data = toUint8Array(m.data);
        const entry = ctx.file.media.addMedia(
          data,
          m.type,
          (fileName) =>
            ({
              type: m.type,
              data,
              fileName,
              transformation: { width: 0, height: 0 },
            }) as MediaData,
          m.fileName,
        );
        const placeholder = `{${m.fileName}}`;
        if (fill?.relationshipId === placeholder) fill.relationshipId = `{${entry.fileName}}`;
      }
      vml = fill ? { ...vml, fill } : vml;
    }
    return `<w:background ${attrStr}>${stringifyVmlBackground(vml)}</w:background>`;
  }

  if (opts.image) {
    const image = opts.image;
    const rawData = toUint8Array(image.data) as Uint8Array;
    const { fileName } = ctx.file.media.addMedia(
      rawData,
      image.type,
      (name) =>
        ({
          type: image.type as "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf",
          data: rawData,
          fileName: name,
          transformation: { width: 0, height: 0 },
        }) as MediaData,
    );

    const vmlBg = stringifyVmlBackground({
      // stringifyDocumentXml emits the background before body children, so the
      // allocator hands out 1025 here — matching Word's fixed v:background id.
      id: nextVmlShapeId(ctx.reproducible),
      fill: {
        type: "frame",
        recolor: true,
        relationshipId: `{${fileName}}`,
        officeTitle: fileName,
      },
    });
    return `<w:background ${attrStr}>${vmlBg}</w:background>`;
  }

  return `<w:background ${attrStr}/>`;
}

// ── Textbox (pure function, no XmlComponent) ──

function stringifyTextbox(
  opts: Omit<ParagraphOptions, "style" | "children"> & {
    style?: VmlShapeStyle;
    children?: SectionChild[];
  },
  ctx: BodyContext,
): string {
  // Destructure to separate VML style/children from paragraph properties
  const { style, children, ...paraOpts } = opts;
  const props = stringifyParagraphProperties(paraOpts, ctx.reproducible);
  const pPrXml = props.xml ?? "";

  // Textbox content — serialize children via stringifyBodyChild
  const contentParts: string[] = [];
  if (children) {
    for (const c of children) {
      contentParts.push(stringifyBodyChild(c, ctx));
    }
  }
  const txbxContent = contentParts.join("");

  const vshape = stringifyVmlShape({
    id: nextVmlShapeId(ctx.reproducible),
    type: "#_x0000_t202",
    style,
    textbox: {
      style: { fitShapeToText: true },
      insetmode: "auto",
      txbxContent,
    },
  });

  // w:pict is run-level content (EG_RunInnerContent) — Word rejects it bare
  // under w:p, so the textbox paragraph carries its VML shape inside a run
  return `<w:p>${pPrXml}<w:r><w:pict>${vshape}</w:pict></w:r></w:p>`;
}

// ── Document body ──

/** Document-level namespace keys (MS Word declaration order). */
const DOC_NAMESPACE_KEYS = [
  "wpc",
  "mc",
  "o",
  "r",
  "m",
  "v",
  "wp14",
  "wp",
  "w10",
  "w",
  "w14",
  "w15",
  "wpg",
  "wpi",
  "wne",
  "wps",
  "cx",
  "cx1",
  "cx2",
  "cx3",
  "cx4",
  "cx5",
  "cx6",
  "cx7",
  "cx8",
  "aink",
  "am3d",
  "w16cex",
  "w16cid",
  "w16",
  "w16sdtdh",
  "w16se",
] as const;

/**
 * Stringify the complete document.xml from context data.
 *
 * This is the pure JSON path — iterates raw section children via
 * `stringifyBodyChild()` while reusing existing SectionProperties
 * instances for sectPr XML (which are already created with correct
 * header/footer references).
 *
 * Produces the complete `<w:document>` element including namespaces,
 * background, body content with interleaved section properties.
 */
export function stringifyDocumentXml(ctx: DocxWriteContext, docCtx: BodyContext): string {
  const sections = ctx._options.sections;
  const bodySections = ctx.sectionProperties;
  const parts: string[] = [];

  // <w:document> open tag
  const conformanceAttr = ctx._options.conformance
    ? ` w:conformance="${ctx._options.conformance}"`
    : "";
  const docNs = documentNamespaceAttributesInDialect(
    DOC_NAMESPACE_KEYS,
    documentNamespaceDialect(ctx),
  );
  parts.push(`<w:document ${docNs} mc:Ignorable="w14 w15 wp14"${conformanceAttr}>`);

  // Background (if any)
  if (ctx._options.background) {
    parts.push(stringifyDocumentBackground(ctx._options.background, docCtx));
  }

  // <w:body> — children go straight into `parts` (single join, no intermediate)
  parts.push("<w:body>");

  for (let si = 0; si < sections.length; si++) {
    const section = sections[si]!;
    const children = section.children ?? [];
    const sectPrOpts = bodySections[si];
    const sectPrXml = sectPrOpts ? (sectionPropertiesDesc.stringify(sectPrOpts, docCtx) ?? "") : "";
    const isLast = si === sections.length - 1;
    const lastChild = children.at(-1);
    const rawSectionBreakHosted =
      !isLast &&
      lastChild !== undefined &&
      "rawXml" in lastChild &&
      lastChild.rawXml.includes("<w:sectPr");

    // Per OOXML, a non-final section's sectPr lives in the pPr of its LAST
    // paragraph (the section-break paragraph carries both content and sectPr).
    // Host it there when possible; fall back to a dedicated break paragraph if
    // the section is empty or ends in a non-paragraph child. The last section
    // emits its sectPr at the body level — and only the last section may omit
    // it: a mid-document section with no explicit properties still emits an
    // empty <w:sectPr/> (CT_SectPr children are all optional) so the
    // user-declared boundary survives generate→parse.
    let sectPrHosted = isLast || rawSectionBreakHosted;
    for (let ci = 0; ci < children.length; ci++) {
      const child = children[ci]!;
      const inject =
        !isLast &&
        sectPrXml &&
        ci === children.length - 1 &&
        ("paragraph" in child || "toc" in child);
      if (inject) sectPrHosted = true;
      parts.push(stringifyBodyChild(child, docCtx, inject ? sectPrXml : undefined));
    }
    if (!isLast && !sectPrHosted) {
      parts.push(`<w:p><w:pPr>${sectPrXml || "<w:sectPr/>"}</w:pPr></w:p>`);
    }
    if (isLast && (sectPrXml || sections.length > 1)) {
      // A single-section document without properties stays sectPr-free (Word
      // applies defaults); with multiple sections the final boundary is
      // structural too and emits an empty sectPr to keep the section count.
      parts.push(sectPrXml || "<w:sectPr/>");
    }
  }

  parts.push("</w:body></w:document>");

  return dedupeRevisionIds(parts.join(""));
}

function textElementXml(tag: string, text: string, preserve?: boolean): string {
  const preserveSpace = preserve === true || /^[\t\n\r ]|[\t\n\r ]$/.test(text);
  const attr = preserveSpace ? ' xml:space="preserve"' : "";
  return `<${tag}${attr}>${escapeXml(text)}</${tag}>`;
}

// Revision markers carry @w:id values unique per marker kind (Word pairs a
// move with shared moveFrom/moveTo ids, and producers reuse ids across kinds).
// Hand-written options may repeat an id within one kind; renumber only those
// duplicates so round-tripped sources keep their original ids byte-for-byte.
const REVISION_ID_TAGS =
  /(<w:(?:ins|del|moveFrom|moveTo|cellIns|cellDel|cellMerge|cellSplit|rPrChange|pPrChange|trPrChange|tcPrChange|sectPrChange|tblPrChange|tblGridChange)\b[^>]*w:id=")(\d+)(")/g;

function dedupeRevisionIds(xml: string): string {
  const counts = new Map<string, Map<string, number>>();
  let hasDup = false;
  for (const m of xml.matchAll(REVISION_ID_TAGS)) {
    const tag = /w:(\w+)\b/.exec(m[1]!)?.[1] ?? "";
    const ids = counts.get(tag) ?? new Map<string, number>();
    const seenCount = (ids.get(m[2]!) ?? 0) + 1;
    ids.set(m[2]!, seenCount);
    counts.set(tag, ids);
    if (seenCount > 1) hasDup = true;
  }
  if (!hasDup) return xml;
  const emitted = new Map<string, Set<string>>();
  return xml.replace(REVISION_ID_TAGS, (m, pre: string, id: string, post: string) => {
    const tag = /w:(\w+)\b/.exec(pre)?.[1] ?? "";
    const used = emitted.get(tag) ?? new Set<string>();
    emitted.set(tag, used);
    if (!used.has(id)) {
      used.add(id);
      return m;
    }
    let n = 0;
    while (used.has(String(n))) n++;
    used.add(String(n));
    return `${pre}${n}${post}`;
  });
}

// ────────────────────────────────────────────────────────────────────────────────
// Parse (XML → JSON options)
// ────────────────────────────────────────────────────────────────────────────────

const HEADING_MAP: Record<string, (typeof HeadingLevel)[keyof typeof HeadingLevel]> = {
  Heading1: HeadingLevel.HEADING_1,
  Heading2: HeadingLevel.HEADING_2,
  Heading3: HeadingLevel.HEADING_3,
  Heading4: HeadingLevel.HEADING_4,
  Heading5: HeadingLevel.HEADING_5,
  Heading6: HeadingLevel.HEADING_6,
  Title: HeadingLevel.TITLE,
};

/** Valid w:spacing/`@w:lineRule` values (ST_LineSpacingRule). */
const LINE_RULES = Object.values(LineRuleType) as readonly string[];

// On/off paragraph properties: XML child tag → options key.
const ON_OFF_PARAGRAPH_PROPS = [
  ["w:keepNext", "keepNext"],
  ["w:keepLines", "keepLines"],
  ["w:pageBreakBefore", "pageBreakBefore"],
  ["w:widowControl", "widowControl"],
  ["w:suppressLineNumbers", "suppressLineNumbers"],
  ["w:contextualSpacing", "contextualSpacing"],
  ["w:bidi", "bidirectional"],
  ["w:wordWrap", "wordWrap"],
  ["w:suppressAutoHyphens", "suppressAutoHyphens"],
  ["w:adjustRightInd", "adjustRightInd"],
  ["w:snapToGrid", "snapToGrid"],
  ["w:mirrorIndents", "mirrorIndents"],
  ["w:kinsoku", "kinsoku"],
  ["w:topLinePunct", "topLinePunct"],
  ["w:autoSpaceDE", "autoSpaceDE"],
  ["w:autoSpaceDN", "autoSpaceEastAsianText"],
  ["w:overflowPunct", "overflowPunctuation"],
  ["w:suppressOverlap", "suppressOverlap"],
] as const;

// pBdr side element names (CT_PBdr children order).
const PBDR_SIDES = [
  "top",
  "left",
  "bottom",
  "right",
  "between",
  "bar",
  "insideHorizontal",
  "insideVertical",
] as const;

// Inline element payloads extracted from the ParagraphChild union — used by
// parse to build typed objects instead of Record<string, unknown>.
type SmartTagInlineOptions = Extract<ParagraphChild, { smartTag: unknown }>["smartTag"];
type CustomXmlInlineOptions = Extract<ParagraphChild, { customXml: unknown }>["customXml"];
type DirInlineOptions = Extract<ParagraphChild, { dir: unknown }>["dir"];
type HyperlinkInlineOptions = Extract<ParagraphChild, { hyperlink: unknown }>["hyperlink"];
type PermStartInlineOptions = Extract<ParagraphChild, { permStart: unknown }>["permStart"];

/**
 * Parse w:pPr element into paragraph properties (without children).
 */
export function parseParagraphProperties(
  el: Element,
  ctx: DocxReadContext,
): Partial<ParagraphPropertiesOptions> {
  const opts: Partial<ParagraphPropertiesOptions> = {};

  // Style / heading
  const pStyle = findChild(el, "w:pStyle");
  if (pStyle) {
    const styleVal = attr(pStyle, "w:val");
    if (styleVal) {
      if (HEADING_MAP[styleVal]) {
        opts.heading = HEADING_MAP[styleVal];
      } else {
        opts.style = styleVal;
      }
    }
  }

  // Alignment
  const jc = findChild(el, "w:jc");
  if (jc) {
    const val = attr(jc, "w:val");
    if (val) opts.alignment = xsdJcAlignment.from(val) as ParagraphPropertiesOptions["alignment"];
  }

  // Spacing — before/after/line are ST_TwipsMeasure (number | UniversalMeasure);
  // use attrMeasure so UniversalMeasure round-trips with the stringify side.
  const spacing = findChild(el, "w:spacing");
  if (spacing) {
    const sp: SpacingProperties = {};
    const before = attrMeasure(spacing, "w:before");
    if (before !== undefined && isNegativeMeasure(before)) {
      throw new DocxParseError(
        "Paragraph spacing before cannot be negative",
        ctx.currentPart,
        "/w:spacing/@w:before",
        "invalid-value",
      );
    }
    if (before !== undefined) sp.before = before as SpacingProperties["before"];
    const after = attrMeasure(spacing, "w:after");
    if (after !== undefined && isNegativeMeasure(after)) {
      throw new DocxParseError(
        "Paragraph spacing after cannot be negative",
        ctx.currentPart,
        "/w:spacing/@w:after",
        "invalid-value",
      );
    }
    if (after !== undefined) sp.after = after as SpacingProperties["after"];
    const line = attrMeasure(spacing, "w:line");
    if (line !== undefined) sp.line = line as SpacingProperties["line"];
    const lineRule = attr(spacing, "w:lineRule");
    if (lineRule && (LINE_RULES as readonly string[]).includes(lineRule)) {
      sp.lineRule = lineRule as SpacingProperties["lineRule"];
    }
    const beforeAutoSpacing = attrBool(spacing, "w:beforeAutospacing");
    if (beforeAutoSpacing !== undefined) sp.beforeAutoSpacing = beforeAutoSpacing;
    const afterAutoSpacing = attrBool(spacing, "w:afterAutospacing");
    if (afterAutoSpacing !== undefined) sp.afterAutoSpacing = afterAutoSpacing;
    const beforeLines = attrNum(spacing, "w:beforeLines");
    if (beforeLines !== undefined) sp.beforeLines = beforeLines;
    const afterLines = attrNum(spacing, "w:afterLines");
    if (afterLines !== undefined) sp.afterLines = afterLines;
    opts.spacing = sp;
  }

  // Indent — left/right/start/end are ST_SignedTwipsMeasure, hanging/firstLine
  // are ST_TwipsMeasure (number | UniversalMeasure); use attrMeasure so
  // UniversalMeasure round-trips. *Chars are ST_DecimalNumber (pure number).
  const ind = findChild(el, "w:ind");
  if (ind) {
    const indentObj: IndentProperties = {};
    const left = attrMeasure(ind, "w:left");
    // Lenient producers emit an unprefixed @left on w:ind; the OOXML attribute
    // is namespaced, but Word and LibreOffice both read the tolerated spelling.
    const legacyLeftValue = attr(ind, "left");
    let legacyLeft: number | string | undefined;
    if (legacyLeftValue !== undefined) {
      const legacyNumber = Number(legacyLeftValue);
      legacyLeft = Number.isNaN(legacyNumber) ? legacyLeftValue : legacyNumber;
    }
    if (left !== undefined || legacyLeft !== undefined) {
      indentObj.left = (left ?? legacyLeft) as IndentProperties["left"];
    }
    const leftChars = attrNum(ind, "w:leftChars");
    if (leftChars !== undefined) indentObj.leftChars = leftChars;
    const right = attrMeasure(ind, "w:right");
    if (right !== undefined) indentObj.right = right as IndentProperties["right"];
    const rightChars = attrNum(ind, "w:rightChars");
    if (rightChars !== undefined) indentObj.rightChars = rightChars;
    const start = attrMeasure(ind, "w:start");
    if (start !== undefined) indentObj.start = start as IndentProperties["start"];
    const startChars = attrNum(ind, "w:startChars");
    if (startChars !== undefined) indentObj.startChars = startChars;
    const end = attrMeasure(ind, "w:end");
    if (end !== undefined) indentObj.end = end as IndentProperties["end"];
    const endChars = attrNum(ind, "w:endChars");
    if (endChars !== undefined) indentObj.endChars = endChars;
    const hanging = attrMeasure(ind, "w:hanging");
    const hangingChars = attrNum(ind, "w:hangingChars");
    if (hangingChars !== undefined) indentObj.hangingChars = hangingChars;
    const firstLine = attrMeasure(ind, "w:firstLine");
    const firstLineRaw = attr(ind, "w:firstLine");
    const firstLineChars = attrNum(ind, "w:firstLineChars");
    if (firstLineChars !== undefined) indentObj.firstLineChars = firstLineChars;
    // Preserve generator-authored negative lexical values verbatim; fresh
    // negative convenience values still normalize to the positive twin below.
    if (typeof hanging === "number" && hanging < 0) indentObj.hangingRaw = attr(ind, "w:hanging");
    else if (hanging !== undefined) indentObj.hanging = hanging as IndentProperties["hanging"];
    if (typeof firstLine === "number" && firstLine < 0) indentObj.firstLineRaw = firstLineRaw;
    else if (firstLine !== undefined)
      indentObj.firstLine = firstLine as IndentProperties["firstLine"];
    if (Object.keys(indentObj).length > 0) opts.indent = indentObj;
  }

  // Numbering (w:numPr)
  const numPr = findChild(el, "w:numPr");
  if (numPr) {
    const ilvl = findChild(numPr, "w:ilvl");
    // No w:ilvl → level stays undefined (numId-only numPr inherits the level).
    const level = ilvl ? attrNum(ilvl, "w:val") : undefined;
    const numIdEl = findChild(numPr, "w:numId");
    const numId = numIdEl ? attr(numIdEl, "w:val") : undefined;
    if (numId === "0") {
      // numId=0 is the OOXML reserved value that cancels numbering inherited
      // from the paragraph style — emit numId=0 verbatim instead of falling
      // back to a bullet (which would inject ListParagraph + numId=1). A
      // w:ilvl alongside keeps its level pin.
      opts.numbering = level !== undefined ? { none: true, level } : false;
    } else if (numId === undefined) {
      // A numPr without numId carries nothing resolvable — except tracked
      // revision markers (w:ins / w:numberingChange): the numbering property
      // set itself is a revision inheriting its numbering from the style
      // chain. Keep those, and keep a bare w:ilvl too (level override with
      // the numbering definition still inherited from the style chain); the
      // bullet fallback would fabricate numId=1 + ListParagraph the source
      // never had.
      const insEl = findChild(numPr, "w:ins");
      const numberingChangeEl = findChild(numPr, "w:numberingChange");
      if (level !== undefined) {
        opts.numbering = { levelOnly: true, level };
      } else if (insEl || numberingChangeEl) {
        const rev: {
          revisionOnly: true;
          numberingChange?: { original: string; id: string; author: string; date?: string };
          insertion?: NumberingInsertionOptions;
        } = { revisionOnly: true };
        if (numberingChangeEl) {
          const nc: { original: string; id: string; author: string; date?: string } = {
            original: attr(numberingChangeEl, "w:original") ?? "",
            id: attr(numberingChangeEl, "w:id") ?? "",
            author: attr(numberingChangeEl, "w:author") ?? "",
          };
          const ncDate = attr(numberingChangeEl, "w:date");
          if (ncDate) nc.date = ncDate;
          rev.numberingChange = nc;
        }
        if (insEl) {
          const ins: Partial<NumberingInsertionOptions> = {
            id: attr(insEl, "w:id") ?? "",
            author: attr(insEl, "w:author") ?? "",
          };
          const insDate = attr(insEl, "w:date");
          if (insDate) ins.date = insDate;
          rev.insertion = ins as NumberingInsertionOptions;
        }
        opts.numbering = rev;
      }
    } else {
      // Round-tripped IDs must survive even when the source package has no
      // numbering part or points at a producer-specific unknown definition.
      const numberingOpts: {
        reference: string;
        sourceNumId?: number;
        level?: number;
        autoStyle: boolean;
        numberingChange?: { original: string; id: string; author: string; date?: string };
        insertion?: NumberingInsertionOptions;
      } = {
        reference: `list_${numId}`,
        sourceNumId: Number(numId),
        level,
        autoStyle: false,
      };
      // w:numberingChange (CT_TrackChangeNumbering) — child of w:numPr
      const numberingChangeEl = findChild(numPr, "w:numberingChange");
      if (numberingChangeEl) {
        const nc: { original: string; id: string; author: string; date?: string } = {
          original: attr(numberingChangeEl, "w:original") ?? "",
          id: attr(numberingChangeEl, "w:id") ?? "",
          author: attr(numberingChangeEl, "w:author") ?? "",
        };
        const ncDate = attr(numberingChangeEl, "w:date");
        if (ncDate) nc.date = ncDate;
        numberingOpts.numberingChange = nc;
      }
      // w:ins (CT_TrackChange) — numbering applied as a revision
      const insEl = findChild(numPr, "w:ins");
      if (insEl) {
        const ins: Partial<NumberingInsertionOptions> = {
          id: attr(insEl, "w:id") ?? "",
          author: attr(insEl, "w:author") ?? "",
        };
        const insDate = attr(insEl, "w:date");
        if (insDate) ins.date = insDate;
        numberingOpts.insertion = ins as NumberingInsertionOptions;
      }
      opts.numbering = numberingOpts;
    }
  }

  // Tab stops
  const tabs = findChild(el, "w:tabs");
  if (tabs) {
    const tabStops: TabStopDefinition[] = [];
    for (const tab of tabs.elements ?? []) {
      if (tab.name !== "w:tab") continue;
      const tabObj: Partial<TabStopDefinition> = {};
      const pos = attrMeasure(tab, "w:pos");
      if (pos !== undefined) tabObj.position = pos as TabStopDefinition["position"];
      const val = attr(tab, "w:val");
      if (val) tabObj.type = val as TabStopDefinition["type"];
      const leader = attr(tab, "w:leader");
      if (leader) tabObj.leader = leader as TabStopDefinition["leader"];
      tabStops.push(tabObj as TabStopDefinition);
    }
    opts.tabStops = tabStops;
  }

  // On/off properties
  for (const [name, optKey] of ON_OFF_PARAGRAPH_PROPS) {
    const child = findChild(el, name);
    if (child) opts[optKey] = attrBool(child, "w:val") ?? true;
  }
  const hyphenationLines = findChild(el, "w:hyphenationLines");
  const hyphenationLineCount = hyphenationLines ? attrNum(hyphenationLines, "w:val") : undefined;
  if (hyphenationLineCount !== undefined) opts.hyphenationLines = hyphenationLineCount;
  const contextualSpacing = findChild(el, "w:contextualSpacing");
  const contextualSpacingRaw =
    contextualSpacing === undefined ? undefined : attr(contextualSpacing, "w:val");
  if (
    contextualSpacingRaw !== undefined &&
    (ON_OFF_LEXICAL_VALUES as readonly string[]).includes(contextualSpacingRaw)
  )
    opts.contextualSpacingRaw = contextualSpacingRaw as OnOffLexicalValue;

  // Thematic break
  const pBdr = findChild(el, "w:pBdr");
  if (pBdr) {
    const border: BordersOptions = {};
    for (const side of PBDR_SIDES) {
      const sideName =
        side === "insideHorizontal"
          ? "w:insideH"
          : side === "insideVertical"
            ? "w:insideV"
            : `w:${side}`;
      const sideEl = findChild(pBdr, sideName);
      if (!sideEl) continue;
      const sideOpts = parseBorderSide(sideEl);
      if (!sideOpts) continue;
      border[side] = sideOpts;
    }
    opts.border = border;
  }

  // Shading
  const shd = findChild(el, "w:shd");
  if (shd) {
    const shading = parseShading(shd);
    if (shading) opts.shading = shading;
  }

  // Text alignment
  const textAlignment = findChild(el, "w:textAlignment");
  if (textAlignment) {
    const val = attr(textAlignment, "w:val");
    if (val) opts.textAlignment = val as ParagraphPropertiesOptions["textAlignment"];
  }

  // Outline level
  const outlineLvl = findChild(el, "w:outlineLvl");
  if (outlineLvl) {
    const val = attrNum(outlineLvl, "w:val");
    if (val !== undefined) opts.outlineLevel = val;
  }

  // Text direction (w:textDirection — ST_TextDirection).
  const textDirection = findChild(el, "w:textDirection");
  if (textDirection) {
    const val = attr(textDirection, "w:val");
    if (val) opts.textDirection = val as ParagraphPropertiesOptions["textDirection"];
  }

  // Textbox tight wrap (w:textboxTightWrap).
  const textboxTightWrap = findChild(el, "w:textboxTightWrap");
  if (textboxTightWrap) {
    const val = attr(textboxTightWrap, "w:val");
    if (val) opts.textboxTightWrap = val as ParagraphPropertiesOptions["textboxTightWrap"];
  }

  // HTML div id (w:divId).
  const divIdEl = findChild(el, "w:divId");
  if (divIdEl) {
    const val = attrNum(divIdEl, "w:val");
    if (val !== undefined) opts.divId = val;
  }

  // Conditional formatting style (w:cnfStyle — CT_Cnf w:val or 12 attrs).
  const cnfStyle = findChild(el, "w:cnfStyle");
  if (cnfStyle) {
    const cnf = parseCnfStyle(cnfStyle);
    if (cnf) opts.cnfStyle = cnf;
  }

  // Word 2013+ extension: collapsed paragraph body (w15:collapsed — CT_OnOff).
  const collapsed = findChild(el, "w15:collapsed");
  if (collapsed) opts.collapsed = attrBool(collapsed, "w:val") ?? true;

  // Run properties (paragraph-level defaults) — the paragraph-mark
  // track-change markers (w:ins/w:del inside w:rPr) ride along on them.
  const rPr = findChild(el, "w:rPr");
  if (rPr) {
    const run = parseRunProperties(rPr) as ParagraphRunPropertiesOptions;
    const ins = findChild(rPr, "w:ins");
    if (ins) {
      run.insertion = {
        id: attrNum(ins, "w:id") ?? 0,
        author: attr(ins, "w:author") ?? "",
        date: attr(ins, "w:date"),
      };
    }
    const del = findChild(rPr, "w:del");
    if (del) {
      run.deletion = {
        id: attrNum(del, "w:id") ?? 0,
        author: attr(del, "w:author") ?? "",
        date: attr(del, "w:date"),
      };
    }
    const movedFrom = findChild(rPr, "w:moveFrom");
    if (movedFrom) {
      run.movedFrom = {
        id: attrNum(movedFrom, "w:id") ?? 0,
        author: attr(movedFrom, "w:author") ?? "",
        date: attr(movedFrom, "w:date"),
      };
    }
    const movedTo = findChild(rPr, "w:moveTo");
    if (movedTo) {
      run.movedTo = {
        id: attrNum(movedTo, "w:id") ?? 0,
        author: attr(movedTo, "w:author") ?? "",
        date: attr(movedTo, "w:date"),
      };
    }
    opts.run = run;
  }

  // Frame properties — rebuilt into the public FrameOptions union so parse
  // output re-enters generate unchanged (CT_FramePr attributes are all
  // optional, hence the optional base fields).
  const framePr = findChild(el, "w:framePr");
  if (framePr) {
    // Mutable superset of the FrameOptions union members — cast once at the end.
    const frame: {
      type?: "absolute" | "alignment" | "mixed";
      position?: { x?: number | UniversalMeasure; y?: number | UniversalMeasure };
      alignment?: { x?: string; y?: string };
      anchor?: { horizontal?: string; vertical?: string };
      space?: { horizontal?: number | UniversalMeasure; vertical?: number | UniversalMeasure };
      dropCap?: string;
      lines?: string;
      wrap?: string;
      rule?: string;
      anchorLock?: boolean;
      width?: number | UniversalMeasure;
      height?: number | UniversalMeasure;
      fillColorRaw?: string;
      fillColorAttrRaw?: "fillcolor" | "w:fillcolor";
    } = {};
    for (const [attrName, optName] of [
      ["w:dropCap", "dropCap"],
      ["w:lines", "lines"],
      ["w:wrap", "wrap"],
      ["w:hRule", "rule"],
    ] as const) {
      const val = attr(framePr, attrName);
      if (val !== undefined) frame[optName] = val;
    }
    // Position: absolute coordinates (w:x/w:y) vs alignment (w:xAlign/w:yAlign)
    const x = attrMeasure(framePr, "w:x") as number | UniversalMeasure;
    const y = attrMeasure(framePr, "w:y") as number | UniversalMeasure;
    const xAlign = attr(framePr, "w:xAlign");
    const yAlign = attr(framePr, "w:yAlign");
    const hasPosition = x !== undefined || y !== undefined;
    const hasAlignment = xAlign !== undefined || yAlign !== undefined;
    if (hasPosition && hasAlignment) {
      frame.type = "mixed";
      if (hasPosition) {
        frame.position = {
          ...(x !== undefined ? { x } : {}),
          ...(y !== undefined ? { y } : {}),
        };
      }
      if (hasAlignment) {
        frame.alignment = {
          ...(xAlign !== undefined ? { x: xAlign } : {}),
          ...(yAlign !== undefined ? { y: yAlign } : {}),
        };
      }
    } else if (hasPosition) {
      frame.type = "absolute";
      frame.position = { ...(x !== undefined ? { x } : {}), ...(y !== undefined ? { y } : {}) };
    } else if (hasAlignment) {
      frame.type = "alignment";
      frame.alignment = {
        ...(xAlign ? { x: xAlign } : {}),
        ...(yAlign ? { y: yAlign } : {}),
      };
    }
    // Anchor (hAnchor/vAnchor)
    const hAnchor = attr(framePr, "w:hAnchor");
    const vAnchor = attr(framePr, "w:vAnchor");
    if (hAnchor || vAnchor) {
      frame.anchor = {
        ...(hAnchor ? { horizontal: hAnchor } : {}),
        ...(vAnchor ? { vertical: vAnchor } : {}),
      };
    }
    // Spacing (hSpace/vSpace)
    const hSpace = attrMeasure(framePr, "w:hSpace") as number | UniversalMeasure;
    const vSpace = attrMeasure(framePr, "w:vSpace") as number | UniversalMeasure;
    if (hSpace !== undefined || vSpace !== undefined) {
      frame.space = {
        ...(hSpace !== undefined ? { horizontal: hSpace } : {}),
        ...(vSpace !== undefined ? { vertical: vSpace } : {}),
      };
    }
    const anchorLock = attrBool(framePr, "w:anchorLock");
    if (anchorLock !== undefined) frame.anchorLock = anchorLock;
    const w = attrMeasure(framePr, "w:w") as number | UniversalMeasure;
    if (w !== undefined) frame.width = w;
    const h = attrMeasure(framePr, "w:h") as number | UniversalMeasure;
    if (h !== undefined) frame.height = h;
    const fillColor = attr(framePr, "w:fillcolor") ?? attr(framePr, "fillcolor");
    if (fillColor !== undefined) {
      frame.fillColorRaw = fillColor;
      frame.fillColorAttrRaw =
        attr(framePr, "w:fillcolor") === undefined ? "fillcolor" : "w:fillcolor";
    }
    if (Object.keys(frame).length > 0) opts.frame = frame as FrameOptions;
    else opts.frame = {};
  }

  // Revision (w:pPrChange) — symmetric with stringifyParagraphProperties
  const pPrChange = findChild(el, "w:pPrChange");
  if (pPrChange) {
    const rev: Partial<ParagraphPropertiesChangeOptions> = {};
    const author = attr(pPrChange, "w:author");
    if (author) rev.author = author;
    const revDate = attr(pPrChange, "w:date");
    if (revDate) rev.date = revDate;
    const revId = attrNum(pPrChange, "w:id");
    if (revId !== undefined) rev.id = revId;
    const innerPPr = findChild(pPrChange, "w:pPr");
    if (innerPPr) Object.assign(rev, parseParagraphProperties(innerPPr, ctx));
    if (Object.keys(rev).length > 0) opts.revision = rev as ParagraphPropertiesChangeOptions;
  }

  // A bare <w:pPr/> yields no fields — mark the presence so stringify
  // re-emits the empty element. A pPr whose only child is w:sectPr is not
  // empty: the section model owns that sectPr and re-injects it here, so the
  // marker must not claim the pPr.
  if (Object.keys(opts).length === 0 && !findChild(el, "w:sectPr")) {
    opts.emptyProperties = true;
  }

  return opts;
}

function isNegativeMeasure(value: number | string): boolean {
  return typeof value === "number" ? value < 0 : value.trimStart().startsWith("-");
}

// CT_Cnf attribute order — shared by parseCnfStyle (w:val digits + w: attrs).
const CNF_KEYS = [
  "firstRow",
  "lastRow",
  "firstColumn",
  "lastColumn",
  "oddVBand",
  "evenVBand",
  "oddHBand",
  "evenHBand",
  "firstRowFirstColumn",
  "firstRowLastColumn",
  "lastRowFirstColumn",
  "lastRowLastColumn",
] as const;

/** Read w:cnfStyle (CT_Cnf). Office's canonical w:val is a 12-char [01]*
 *  string (ST_Cnf), one digit per flag in CNF_KEYS order; the XSD also
 *  permits 12 individual w: attributes, parsed as a fallback. */
function parseCnfStyle(el: Element): NonNullable<ParagraphPropertiesOptions["cnfStyle"]> {
  const val = attr(el, "w:val");
  const result: Record<string, boolean | string> = {};
  if (val !== undefined) result.val = val;
  for (const key of CNF_KEYS) {
    const value = attrBool(el, `w:${key}`);
    if (value !== undefined) result[key] = value;
  }
  // An empty object preserves the presence of a bare/all-zero CT_Cnf element;
  // undefined remains reserved for an absent w:cnfStyle child.
  return result as NonNullable<ParagraphPropertiesOptions["cnfStyle"]>;
}

/**
 * Concatenate `<w:t>` text in a run element.
 *
 * Used to capture a textInput form field's current value from the result run
 * (the runs between the `separate` and `end` fldChars). Only `<w:t>` is read —
 * `<w:tab>`/`<w:br>` in results are ignored as rare for user-entered text.
 */
/** Field-chain accumulator state, shared by the paragraph and track-change run loops. */
interface FieldRunState {
  kind: "form" | "complex" | null;
  pendingFormField: FormFieldOptions | null;
  pendingFormFieldFieldLock?: boolean;
  pendingInstruction: string;
  pendingResult: string;
  controlRPr?: string;
  resultRPr?: string;
  /** Field nesting depth (1 = outermost field). */
  depth: number;
  /** The outer field's begin run — checked for a co-located pagination hint. */
  beginRunEl?: Element;
  /** The separator run — its rsid attributes are captured separately. */
  separatorRunEl?: Element;
  collectingResult: boolean;
  /** Source xml:space marker on the first plain result text. */
  resultPreserveSpace?: boolean;
  /** False only inside a TOC aggregator that re-creates orphan stage markers. */
  preserveOrphanMarkers?: boolean;
  /** Instruction-stage run elements (begin → separate/end), buffered for the
   *  plain-shape check at the closing end marker. */
  instrRunEls: Element[];
  /** Ordered fidelity/canonical members when the stage mixes runs and fields. */
  instructionMembers: ComplexFieldInstructionMember[];
  resultMembers: ComplexFieldInstructionMember[];
  /** Source xml:space marker on the plain instruction text. */
  instructionPreserveSpace?: boolean;
  /** Result-stage run elements (separate → end), buffered likewise. */
  resultRunEls: Element[];
  /** Source run when all field stages were packed into a single w:r. */
  compoundRunXml?: string;
  /** True while replaying logical runs from the current compound source run. */
  replayingCompoundRun?: boolean;
  /** End marker run paired with a compound source run. */
  compoundEndRunXml?: string;
  /** Revision metadata embedded in the begin marker. */
  beginNumberingChange?: NumberingChangeOptions;
  /** The currently open nested field; all logical input is delegated to it. */
  nested?: FieldRunState;
}

const initialFieldRunState = (): FieldRunState => ({
  kind: null,
  depth: 0,
  pendingFormField: null,
  pendingFormFieldFieldLock: undefined,
  pendingInstruction: "",
  pendingResult: "",
  collectingResult: false,
  instructionPreserveSpace: false,
  instructionMembers: [],
  resultMembers: [],
  resultPreserveSpace: false,
  instrRunEls: [],
  resultRunEls: [],
  compoundRunXml: undefined,
});

function activeFieldState(state: FieldRunState): FieldRunState {
  let current = state;
  while (current.nested) current = current.nested;
  return current;
}

function isHyperlinkChild(
  child: ParagraphChild,
): child is Extract<ParagraphChild, { hyperlink: unknown }> {
  return "hyperlink" in child;
}

function nestedFieldMember(child: ParagraphChild): ComplexFieldInstructionMember | undefined {
  if ("complexField" in child) return { complexField: child.complexField };
  if ("formField" in child) return child;
  return undefined;
}

/** Collect an unterminated field's buffered control/result runs, including nested fields. */
function incompleteFieldRunElements(state: FieldRunState): Element[] {
  const nestedElements = state.nested ? incompleteFieldRunElements(state.nested) : [];
  return [
    state.beginRunEl,
    ...state.instrRunEls,
    ...(state.collectingResult && state.separatorRunEl ? [state.separatorRunEl] : []),
    ...nestedElements,
    ...state.resultRunEls,
  ].filter((element): element is Element => element !== undefined);
}

/** Serialize an unterminated field's buffered control/result runs verbatim. */
function flushIncompleteFieldRunState(state: FieldRunState): { rawXml: string } | undefined {
  if (!state.kind) return undefined;
  const elements = incompleteFieldRunElements(state);
  if (elements.length === 0) return undefined;
  return { rawXml: elements.map((element) => stringifyElement(element)).join("") };
}

function parseNumberingChange(element: Element): NumberingChangeOptions {
  const date = attr(element, "w:date");
  return {
    original: attr(element, "w:original") ?? "",
    id: attr(element, "w:id") ?? "",
    author: attr(element, "w:author") ?? "",
    ...(date !== undefined ? { date } : {}),
  };
}

function parseFieldMarkerRun(run: Element, type: "separate" | "end"): FieldMarkerOptions {
  const marker: FieldMarkerOptions = { type };
  const rPrXml = runRPrXml(run);
  if (rPrXml !== undefined) marker.rPrXml = rPrXml;
  const additionRsid = attr(run, "w:rsidR");
  if (additionRsid) marker.additionRsid = additionRsid;
  const runPropertiesRsid = attr(run, "w:rsidRPr");
  if (runPropertiesRsid) marker.runPropertiesRsid = runPropertiesRsid;
  const deletionRsid = attr(run, "w:rsidDel");
  if (deletionRsid) marker.deletionRsid = deletionRsid;
  if (findChild(run, "w:lastRenderedPageBreak")) marker.lastRenderedPageBreak = true;
  const fldChar = findChild(run, "w:fldChar");
  const numberingChangeEl = fldChar && findChild(fldChar, "w:numberingChange");
  if (numberingChangeEl) marker.numberingChange = parseNumberingChange(numberingChangeEl);
  return marker;
}

/**
 * Release an unterminated field inside a track-change wrapper. Control runs
 * stay verbatim, while completed result members are already canonical and must
 * not disappear with the unclosed outer field.
 */
function incompleteTrackChangeFieldChildren(state: FieldRunState): TrackChangeChild[] {
  if (!state.kind) return [];

  const headerState: FieldRunState = {
    ...state,
    resultRunEls: state.nested ? state.resultRunEls : [],
  };
  const header = flushIncompleteFieldRunState(headerState);
  const children: TrackChangeChild[] = header ? [header] : [];
  const members = state.collectingResult ? state.resultMembers : state.instructionMembers;

  for (const member of members) {
    if ("run" in member) {
      if (state.collectingResult) children.push(member.run);
      continue;
    }
    if ("complexField" in member || "formField" in member) {
      children.push(member);
    }
  }
  return children;
}

/**
 * True when the field-stage runs are exactly what the plain stringify template
 * reproduces: a single run with the expected rPr whose only child is the
 * stage's text element. Multi-run chains (leading/trailing space runs with
 * their own rStyle, w:br inside the format switch, per-run rFonts on result
 * segments) need the verbatim channel.
 */
function isPlainFieldRuns(
  runs: Element[],
  expectedRPr: string | undefined,
  tags: string[],
): boolean {
  if (runs.length === 0) return true;
  if (runs.length > 1) return false;
  const run = runs[0]!;
  if (runRPrXml(run) !== expectedRPr) return false;
  const content = (run.elements ?? []).filter((c) => c.name !== "w:rPr");
  return content.length === 1 && tags.includes(content[0]!.name!);
}

/**
 * Split a run that packs a complete (often nested) field into one XML run.
 * Word normally emits one field stage per run, but some producers put every
 * marker and instruction in one `w:r`; feeding that run only once would drop
 * every stage after its first marker.
 */
function splitCompoundFieldRun(run: Element): Element[] {
  const elements = run.elements ?? [];
  const fieldControls = elements.filter(
    (child) =>
      child.name === "w:fldChar" ||
      child.name === "w:instrText" ||
      child.name === "w:delInstrText" ||
      child.name === "w:t" ||
      child.name === "w:delText",
  );
  if (!elements.some((child) => child.name === "w:fldChar") || fieldControls.length < 2) {
    return [run];
  }

  const properties = elements.find((child) => child.name === "w:rPr");
  return fieldControls.map((child) => ({
    ...structuredClone(run),
    elements: properties
      ? [structuredClone(properties), structuredClone(child)]
      : [structuredClone(child)],
  }));
}

/**
 * Feed one w:r element into the field accumulator. A field (form field OR
 * plain complex field) spans several runs (begin fldChar → instrText →
 * separate → result → end):
 * - Form fields carry w:ffData on the begin fldChar; their state lives there,
 *   and only a textInput's result is captured (as `value`).
 * - Plain complex fields (PAGE/DATE/TOC/HYPERLINK...) have no ffData; their
 *   instrText + result are captured as a complexField child for round-trip.
 * Control and in-field runs are consumed; the closing run returns the whole
 * field collapsed to a single child. Runs outside any field pass through.
 */
function feedFieldRun(
  run: Element,
  state: FieldRunState,
  ctx: DocxReadContext,
): { consumed: boolean; child?: ParagraphChild } {
  if (state.nested) {
    const fed = feedFieldRun(run, state.nested, ctx);
    const member = fed.child ? nestedFieldMember(fed.child) : undefined;
    if (member) {
      state.depth--;
      (state.collectingResult ? state.resultMembers : state.instructionMembers).push(member);
      state.nested = undefined;
    }
    return { consumed: true };
  }

  const logicalRuns = splitCompoundFieldRun(run);
  if (logicalRuns.length > 1) {
    state.compoundRunXml = stringifyElement(run);
    state.compoundEndRunXml = undefined;
    state.replayingCompoundRun = true;
    let child: ParagraphChild | undefined;
    for (const logicalRun of logicalRuns) {
      const fed = feedFieldRun(logicalRun, state, ctx);
      if (fed.child) child = fed.child;
    }
    state.replayingCompoundRun = false;
    return { consumed: true, child };
  }

  const fldCharEl = findChild(run, "w:fldChar");
  if (fldCharEl) {
    const fctype = attr(fldCharEl, "w:fldCharType");
    if (fctype === "begin") {
      // A nested field opening inside an outer field's instruction/result span
      // (e.g. `t "<seq Appendix>-<seq Figure>"` marker fields): keep the whole
      // nested chain verbatim in the outer field's stage buffer instead of
      // resetting the accumulator — the plain state machine cannot represent
      // interleaved instructions, and the verbatim channel round-trips them.
      if (state.kind && state.depth > 0) {
        state.depth++;
        state.nested = initialFieldRunState();
        feedFieldRun(run, state.nested, ctx);
        return { consumed: true };
      }
      state.pendingFormFieldFieldLock = undefined;
      const ffDataEl = findChild(fldCharEl, "w:ffData");
      if (ffDataEl) {
        state.kind = "form";
        state.pendingFormField = parseFormFieldData(ffDataEl);
        state.pendingFormFieldFieldLock = attrBool(fldCharEl, "w:fldLock");
        state.instructionPreserveSpace = false;
        state.resultPreserveSpace = false;
      } else {
        state.kind = "complex";
        state.pendingInstruction = "";
        state.pendingResult = "";
        state.instructionPreserveSpace = false;
        state.resultPreserveSpace = false;
      }
      // The begin run's rPr stands in for the field's control-run rPr.
      state.controlRPr = runRPrXml(run);
      const numberingChangeEl = findChild(fldCharEl, "w:numberingChange");
      state.beginNumberingChange = numberingChangeEl
        ? parseNumberingChange(numberingChangeEl)
        : undefined;
      state.beginRunEl = run;
      state.separatorRunEl = undefined;
      state.resultRPr = undefined;
      state.collectingResult = false;
      state.depth = 1;
      state.instrRunEls = [];
      state.instructionMembers = [];
      state.resultMembers = [];
      state.resultRunEls = [];
    } else if (fctype === "separate") {
      if (state.kind && state.depth > 1) {
        if (state.nested) feedFieldRun(run, state.nested, ctx);
        else (state.collectingResult ? state.resultRunEls : state.instrRunEls).push(run);
        return { consumed: true };
      }
      state.collectingResult = true;
      state.separatorRunEl = run;
      state.resultRunEls = [];
    } else if (fctype === "end" && state.kind) {
      if (state.kind && state.depth > 1) {
        state.depth--;
        if (state.nested) {
          const fed = feedFieldRun(run, state.nested, ctx);
          const member = fed.child ? nestedFieldMember(fed.child) : undefined;
          if (member) {
            (state.collectingResult ? state.resultMembers : state.instructionMembers).push(member);
            state.nested = undefined;
          }
        } else (state.collectingResult ? state.resultRunEls : state.instrRunEls).push(run);
        return { consumed: true };
      }
      if (state.compoundRunXml && !state.replayingCompoundRun) {
        state.compoundEndRunXml = stringifyElement(run);
      }
      let child: ParagraphChild | undefined;
      if (state.kind === "form" && state.pendingFormField) {
        const formChild: ParagraphChild = { formField: state.pendingFormField };
        if (state.pendingFormFieldFieldLock !== undefined) {
          formChild.fieldLock = state.pendingFormFieldFieldLock;
        }
        const beginRsid = attr(state.beginRunEl ?? run, "w:rsidR");
        if (beginRsid) formChild.additionRsid = beginRsid;
        const beginRunPropertiesRsid = attr(state.beginRunEl ?? run, "w:rsidRPr");
        if (beginRunPropertiesRsid) formChild.runPropertiesRsid = beginRunPropertiesRsid;
        if (state.controlRPr) formChild.rPrXml = state.controlRPr;
        if (findChild(state.beginRunEl ?? run, "w:lastRenderedPageBreak")) {
          formChild.lastRenderedPageBreak = true;
        }
        const separatorRsid = state.separatorRunEl
          ? attr(state.separatorRunEl, "w:rsidR")
          : undefined;
        const separatorRunPropertiesRsid = state.separatorRunEl
          ? attr(state.separatorRunEl, "w:rsidRPr")
          : undefined;
        if (separatorRsid) formChild.separatorAdditionRsid = separatorRsid;
        if (separatorRunPropertiesRsid)
          formChild.separatorRunPropertiesRsid = separatorRunPropertiesRsid;
        formChild.separatorRPrXml = state.separatorRunEl
          ? (runRPrXml(state.separatorRunEl) ?? "")
          : undefined;
        const instructionRun = state.instrRunEls.find(
          (el) => findChild(el, "w:instrText") ?? findChild(el, "w:delInstrText"),
        );
        const instructionRsid = instructionRun ? attr(instructionRun, "w:rsidR") : undefined;
        const instructionRunPropertiesRsid = instructionRun
          ? attr(instructionRun, "w:rsidRPr")
          : undefined;
        if (instructionRsid) formChild.instructionAdditionRsid = instructionRsid;
        if (instructionRunPropertiesRsid)
          formChild.instructionRunPropertiesRsid = instructionRunPropertiesRsid;
        const instructionRPrXml = instructionRun ? runRPrXml(instructionRun) : undefined;
        formChild.instructionRPrXml = instructionRPrXml;
        formChild.hasResult = state.resultRunEls.length > 0 || state.separatorRunEl !== undefined;
        if (formChild.hasResult && state.resultRunEls.length === 0) {
          formChild.resultRunsXml = "";
        }
        if (
          !isPlainFieldRuns(state.instrRunEls, state.controlRPr, ["w:instrText", "w:delInstrText"])
        ) {
          formChild.instrRunsXml = state.instrRunEls.map((el) => stringifyElement(el)).join("");
        }
        const hasFormFieldInstructionMarkers = state.instructionMembers.some(
          (member) => !("run" in member),
        );
        if (hasFormFieldInstructionMarkers) formChild.instructionMembers = state.instructionMembers;
        if (
          !isPlainFieldRuns(state.resultRunEls, state.resultRPr ?? state.controlRPr, [
            "w:t",
            "w:delText",
          ])
        ) {
          formChild.resultRunsXml = state.resultRunEls.map((el) => stringifyElement(el)).join("");
        }
        const hasFormFieldResultMarkers = state.resultMembers.some((member) => !("run" in member));
        if (hasFormFieldResultMarkers) formChild.resultMembers = state.resultMembers;
        const resultRunEl = state.resultRunEls.at(-1);
        const resultRsid = resultRunEl ? attr(resultRunEl, "w:rsidR") : undefined;
        const resultRPrRsid = resultRunEl ? attr(resultRunEl, "w:rsidRPr") : undefined;
        if (resultRsid) formChild.resultAdditionRsid = resultRsid;
        if (resultRPrRsid) formChild.resultRunPropertiesRsid = resultRPrRsid;
        const resultRPrXml = resultRunEl ? runRPrXml(resultRunEl) : undefined;
        formChild.resultRPrXml = resultRPrXml;
        const endRPr = runRPrXml(run);
        formChild.endRPrXml = endRPr ?? "";
        const endRsid = attr(run, "w:rsidR");
        const endRunPropertiesRsid = attr(run, "w:rsidRPr");
        if (endRsid) formChild.endAdditionRsid = endRsid;
        if (endRunPropertiesRsid) formChild.endRunPropertiesRsid = endRunPropertiesRsid;
        child = formChild;
      } else if (state.kind === "complex") {
        const cf: ComplexFieldOptions = { instruction: state.pendingInstruction };
        const fieldDataEl = findChild(findChild(state.beginRunEl ?? run, "w:fldChar"), "w:fldData");
        const fieldData = fieldDataEl ? textOf(fieldDataEl) : undefined;
        if (fieldData !== undefined) cf.fieldData = fieldData;
        const beginFieldLock = attrBool(
          findChild(state.beginRunEl ?? run, "w:fldChar"),
          "w:fldLock",
        );
        if (beginFieldLock !== undefined) cf.fieldLock = beginFieldLock;
        // Pagination hint parked on the begin run — re-emit it on the begin run.
        if (findChild(state.beginRunEl ?? run, "w:lastRenderedPageBreak")) {
          cf.lastRenderedPageBreak = true;
        }
        const beginRsid = attr(state.beginRunEl ?? run, "w:rsidR");
        if (beginRsid) cf.additionRsid = beginRsid;
        const beginRunPropertiesRsid = attr(state.beginRunEl ?? run, "w:rsidRPr");
        if (beginRunPropertiesRsid) cf.runPropertiesRsid = beginRunPropertiesRsid;
        const beginDeletionRsid = attr(state.beginRunEl ?? run, "w:rsidDel");
        if (beginDeletionRsid) cf.deletionRsid = beginDeletionRsid;
        const instructionRun = state.instrRunEls.find(
          (el) => findChild(el, "w:instrText") ?? findChild(el, "w:delInstrText"),
        );
        const instructionRsid = instructionRun ? attr(instructionRun, "w:rsidR") : undefined;
        const instructionRunPropertiesRsid = instructionRun
          ? attr(instructionRun, "w:rsidRPr")
          : undefined;
        if (instructionRsid) cf.instructionAdditionRsid = instructionRsid;
        if (instructionRunPropertiesRsid)
          cf.instructionRunPropertiesRsid = instructionRunPropertiesRsid;
        const instructionDeletionRsid = instructionRun
          ? attr(instructionRun, "w:rsidDel")
          : undefined;
        if (instructionDeletionRsid) cf.instructionDeletionRsid = instructionDeletionRsid;
        cf.instructionRPrXml = instructionRun ? runRPrXml(instructionRun) : undefined;
        const separatorRsid = state.separatorRunEl
          ? attr(state.separatorRunEl, "w:rsidR")
          : undefined;
        const separatorRunPropertiesRsid = state.separatorRunEl
          ? attr(state.separatorRunEl, "w:rsidRPr")
          : undefined;
        if (separatorRsid) cf.separatorAdditionRsid = separatorRsid;
        if (separatorRunPropertiesRsid) cf.separatorRunPropertiesRsid = separatorRunPropertiesRsid;
        const separatorDeletionRsid = state.separatorRunEl
          ? attr(state.separatorRunEl, "w:rsidDel")
          : undefined;
        if (separatorDeletionRsid) cf.separatorDeletionRsid = separatorDeletionRsid;
        cf.separatorRPrXml = state.separatorRunEl
          ? (runRPrXml(state.separatorRunEl) ?? "")
          : undefined;
        // Mark the result present when the field carried any result run — an
        // empty-text result still round-trips its separate marker.
        cf.hasResult = state.separatorRunEl !== undefined;
        if (state.resultRunEls.length > 0) cf.result = state.pendingResult;
        cf.instructionPreserveSpace = state.instructionPreserveSpace;
        cf.resultPreserveSpace = state.resultPreserveSpace;
        if (state.controlRPr) cf.rPrXml = state.controlRPr;
        cf.resultRPrXml = state.resultRPr;
        const plainResultRun = state.resultRunEls.length === 1 ? state.resultRunEls[0] : undefined;
        const resultRsid = plainResultRun ? attr(plainResultRun, "w:rsidR") : undefined;
        const resultRPrRsid = plainResultRun ? attr(plainResultRun, "w:rsidRPr") : undefined;
        if (resultRsid) cf.resultAdditionRsid = resultRsid;
        if (resultRPrRsid) cf.resultRunPropertiesRsid = resultRPrRsid;
        const resultDeletionRsid = plainResultRun ? attr(plainResultRun, "w:rsidDel") : undefined;
        if (resultDeletionRsid) cf.resultDeletionRsid = resultDeletionRsid;
        // Word styles the end run like the result (not like the controls).
        const endRPr = runRPrXml(run);
        cf.endRPrXml = endRPr ?? "";
        const endRsid = attr(run, "w:rsidR");
        const endRunPropertiesRsid = attr(run, "w:rsidRPr");
        if (endRsid) cf.endAdditionRsid = endRsid;
        if (endRunPropertiesRsid) cf.endRunPropertiesRsid = endRunPropertiesRsid;
        const endDeletionRsid = attr(run, "w:rsidDel");
        if (endDeletionRsid) cf.endDeletionRsid = endDeletionRsid;
        const endFldChar = findChild(run, "w:fldChar");
        const endNumberingChangeEl = endFldChar && findChild(endFldChar, "w:numberingChange");
        const endNumberingChange = endNumberingChangeEl
          ? parseNumberingChange(endNumberingChangeEl)
          : undefined;
        const hasInstructionMarkers = state.instructionMembers.some((member) => !("run" in member));
        if (hasInstructionMarkers) {
          cf.instructionMembers = state.instructionMembers;
        } else if (
          !isPlainFieldRuns(state.instrRunEls, state.controlRPr, ["w:instrText", "w:delInstrText"])
        ) {
          cf.instrRunsXml = state.instrRunEls.map((el) => stringifyElement(el)).join("");
        }
        // The plain result template pairs the result rPr (defaulting to the
        // control rPr) with a single text run — anything else goes verbatim.
        const hasResultMarkers = state.resultMembers.some((member) => !("run" in member));
        if (hasResultMarkers) cf.resultMembers = state.resultMembers;
        if (state.compoundRunXml) {
          cf.compoundRunXml = state.compoundRunXml;
          cf.compoundEndRunXml = state.compoundEndRunXml;
        }
        if (state.beginNumberingChange) cf.beginNumberingChange = state.beginNumberingChange;
        if (endNumberingChange) cf.endNumberingChange = endNumberingChange;
        if (
          !isPlainFieldRuns(state.resultRunEls, state.resultRPr ?? state.controlRPr, [
            "w:t",
            "w:delText",
          ])
        ) {
          cf.resultRunsXml = state.resultRunEls.map((el) => stringifyElement(el)).join("");
        }
        child = { complexField: cf };
      }
      state.kind = null;
      state.depth = 0;
      state.pendingFormField = null;
      state.pendingFormFieldFieldLock = undefined;
      state.collectingResult = false;
      state.instrRunEls = [];
      state.resultRunEls = [];
      state.compoundRunXml = undefined;
      state.compoundEndRunXml = undefined;
      state.beginNumberingChange = undefined;
      state.nested = undefined;
      return { consumed: true, child };
    }
    // An end (or separate) marker with no open field is the tail of a
    // cross-paragraph field the accumulator never saw. A bare marker run is
    // consumed silently (the field's own emit re-creates it); one carrying
    // extra content (pagination hint, rPr) round-trips verbatim so nothing
    // is lost.
    if (!state.kind) {
      // Only a run carrying real content beyond the marker (e.g. a pagination
      // hint) needs the verbatim channel. Separate/end markers are also kept:
      // in a non-TOC accumulator they belong to a field opened elsewhere, and
      // dropping the marker would move the field boundary.
      const extras = (run.elements ?? []).filter(
        (c) => c.type === "element" && c.name !== "w:fldChar" && c.name !== "w:rPr",
      );
      const preserveMarker =
        state.preserveOrphanMarkers !== false && (fctype === "separate" || fctype === "end");
      if (extras.length > 0) {
        return { consumed: true, child: { rawXml: stringifyElement(run) } };
      }
      if (preserveMarker) {
        return { consumed: true, child: { fieldMarker: parseFieldMarkerRun(run, fctype) } };
      }
    }
    return { consumed: true };
  }
  if (state.kind) {
    if (state.kind === "complex") {
      if (state.nested) return feedFieldRun(run, state.nested, ctx);
      if (state.collectingResult) {
        // Capture the first result run's rPr for round-trip.
        if (state.resultRPr === undefined) state.resultRPr = runRPrXml(run);
        const resultTextEl = findChild(run, "w:t") ?? findChild(run, "w:delText");
        if (resultTextEl && attr(resultTextEl, "xml:space") === "preserve") {
          state.resultPreserveSpace = true;
        }
        state.pendingResult += collectRunText(run);
        state.resultRunEls.push(run);
        state.resultMembers.push({
          run: parsedRunToOptions(parseRun(run, ctx)) as RunOptions,
        });
      } else {
        // Deleted fields spell the instruction w:delInstrText instead.
        const instrEl = findChild(run, "w:instrText") ?? findChild(run, "w:delInstrText");
        if (instrEl) {
          state.pendingInstruction += textOf(instrEl);
          state.instructionPreserveSpace ||= attr(instrEl, "xml:space") === "preserve";
        }
        // Buffer the run for the verbatim channel when its shape is not what
        // the plain instruction template reproduces (per-run rPr, w:br...).
        state.instrRunEls.push(run);
        state.instructionMembers.push({
          run: parsedRunToOptions(parseRun(run, ctx)) as RunOptions,
        });
      }
    } else if (state.kind === "form") {
      if (state.nested) return feedFieldRun(run, state.nested, ctx);
      if (state.collectingResult) {
        // Capture a textInput's current value; checkbox/dropdown results are
        // discarded (their state is in w:ffData), but all result runs stay in
        // the identity buffer so their rsids survive generation.
        if (state.pendingFormField?.textInput) {
          const text = collectRunText(run);
          if (text) {
            const ti = state.pendingFormField.textInput;
            ti.value = (ti.value ?? "") + text;
          }
        }
        state.resultRunEls.push(run);
        state.resultMembers.push({
          run: parsedRunToOptions(parseRun(run, ctx)) as RunOptions,
        });
      } else {
        state.instrRunEls.push(run);
        state.instructionMembers.push({
          run: parsedRunToOptions(parseRun(run, ctx)) as RunOptions,
        });
      }
    }
    return { consumed: true };
  }
  return { consumed: false };
}

/**
 * Parse every drawing-shaped member of a w:r in source order. DrawingML can
 * occur directly or inside MCE Choice wrappers, and a run may carry several
 * shapes. Mixed or repeated members stay separate instead of collapsing to
 * the first shape.
 */
function parseDrawingRunChildren(child: Element, ctx: DocxReadContext): ParagraphChild[] {
  const innerContent = (child.elements ?? []).filter(
    (element) =>
      element.type === "element" &&
      element.name !== "w:rPr" &&
      element.name !== "w:lastRenderedPageBreak",
  );
  const drawingElements = innerContent.filter((element) => element.name === "w:drawing");
  const alternateElements = innerContent.filter(
    (element) => element.name === "mc:AlternateContent",
  );
  const drawingOnly = innerContent.every(
    (element) => element.name === "w:drawing" || element.name === "mc:AlternateContent",
  );
  if (
    !drawingOnly ||
    drawingElements.length > 1 ||
    alternateElements.length > 1 ||
    (drawingElements.length === 1 && innerContent.length > 1)
  ) {
    return [parsedRunToOptions(parseRun(child, ctx))];
  }
  const single = parseSingleDrawingRunChild(child, ctx);
  return single ? [single] : [];
}

/**
 * Parse one direct w:drawing or one MCE-wrapped drawing. MCE Fallbacks remain
 * attached to the structured Choice so both representations round-trip.
 */
function parseSingleDrawingRunChild(
  child: Element,
  ctx: DocxReadContext,
): ParagraphChild | undefined {
  const innerContent = (child.elements ?? []).filter(
    (element) =>
      element.type === "element" &&
      element.name !== "w:rPr" &&
      element.name !== "w:lastRenderedPageBreak",
  );
  const drawingElements = innerContent.filter((element) => element.name === "w:drawing");
  if (drawingElements.length > 1 || (drawingElements.length === 1 && innerContent.length > 1)) {
    return parsedRunToOptions(parseRun(child, ctx));
  }
  let drawingEl = findChild(child, "w:drawing");
  let altFallback: string | undefined;
  let altFallbackMedia: BackgroundRawMediaOptions[] | undefined;
  let altRequires: string | undefined;
  if (!drawingEl) {
    const alt = findChild(child, "mc:AlternateContent");
    if (alt) {
      const choice = findChild(alt, "mc:Choice");
      if (choice) {
        drawingEl = findChild(choice, "w:drawing");
        altRequires = attr(choice, "Requires");
      }
      const fallback = findChild(alt, "mc:Fallback");
      if (fallback) {
        // Replace the VML fallback's r:id/r:embed/r:link refs with {fileName}
        // placeholders and collect the media; otherwise the carried source
        // rIds dangle (not defined in the generated rels).
        const replaced = replaceRelsWithPlaceholders(stringifyElement(fallback), ctx);
        altFallback = replaced.rawXml;
        altFallbackMedia = replaced.rawMedia.length > 0 ? replaced.rawMedia : undefined;
      }
    }
  }
  if (!drawingEl) return undefined;
  // Parse the wrapping run's rPr into structured fields so round-trip stays
  // editable (drawings/shapes can be wrapped in <w:r><w:rPr>…</w:rPr>…).
  const rPrEl = findChild(child, "w:rPr");
  const runProperties = rPrEl ? parseRunProperties(rPrEl) : undefined;
  const additionRsid = attr(child, "w:rsidR");
  const deletionRsid = attr(child, "w:rsidDel");
  const runPropertiesRsid = attr(child, "w:rsidRPr");
  const lastRenderedPageBreak = findChild(child, "w:lastRenderedPageBreak") !== undefined;
  // Attach the VML fallback + Choice Requires so stringify can rebuild the
  // mc:AlternateContent wrapper (Choice structured + Fallback raw).
  const attachAlternateContent = (
    target:
      | {
          vmlFallback?: string;
          vmlFallbackMedia?: BackgroundRawMediaOptions[];
          mcChoiceRequires?: string;
        }
      | undefined,
  ): void => {
    if (!target || !altFallback) return;
    target.vmlFallback = altFallback;
    target.vmlFallbackMedia = altFallbackMedia;
    if (altRequires) target.mcChoiceRequires = altRequires;
  };
  const drawingChild = parseDrawingRun(drawingEl, ctx);
  if (!drawingChild) {
    const graphicData = findFirst(drawingEl, "a:graphicData");
    if (!graphicData) return undefined;
    const unsupported = parseUnsupportedDrawingRun(drawingEl, graphicData, ctx);
    attachAlternateContent(unsupported.unsupportedDrawing);
    if (runProperties) unsupported.unsupportedDrawing.runProperties = runProperties;
    if (additionRsid) unsupported.unsupportedDrawing.additionRsid = additionRsid;
    if (runPropertiesRsid) unsupported.unsupportedDrawing.runPropertiesRsid = runPropertiesRsid;
    if (lastRenderedPageBreak) unsupported.unsupportedDrawing.lastRenderedPageBreak = true;
    return unsupported;
  }
  if (altFallback) {
    if ("wpsShape" in drawingChild) {
      drawingChild.wpsShape.vmlFallback = altFallback;
      drawingChild.wpsShape.vmlFallbackMedia = altFallbackMedia;
      if (altRequires) drawingChild.wpsShape.mcChoiceRequires = altRequires;
    } else if ("wpgGroup" in drawingChild) {
      drawingChild.wpgGroup.vmlFallback = altFallback;
      drawingChild.wpgGroup.vmlFallbackMedia = altFallbackMedia;
      if (altRequires) drawingChild.wpgGroup.mcChoiceRequires = altRequires;
    } else if ("wordprocessingCanvas" in drawingChild) {
      drawingChild.wordprocessingCanvas.vmlFallback = altFallback;
      drawingChild.wordprocessingCanvas.vmlFallbackMedia = altFallbackMedia;
      if (altRequires) drawingChild.wordprocessingCanvas.mcChoiceRequires = altRequires;
    } else if ("unsupportedDrawing" in drawingChild) {
      drawingChild.unsupportedDrawing.vmlFallback = altFallback;
      drawingChild.unsupportedDrawing.vmlFallbackMedia = altFallbackMedia;
      if (altRequires) drawingChild.unsupportedDrawing.mcChoiceRequires = altRequires;
    }
  }
  if (runProperties) {
    if ("picture" in drawingChild) {
      drawingChild.picture.runProperties = runProperties;
    } else if ("wpsShape" in drawingChild) {
      drawingChild.wpsShape.runProperties = runProperties;
    } else if ("wpgGroup" in drawingChild) {
      drawingChild.wpgGroup.runProperties = runProperties;
    } else if ("wordprocessingCanvas" in drawingChild) {
      drawingChild.wordprocessingCanvas.runProperties = runProperties;
    } else if ("chart" in drawingChild) {
      drawingChild.chart.runProperties = runProperties;
    } else if ("chartEx" in drawingChild) {
      drawingChild.chartEx.runProperties = runProperties;
    } else if ("smartArt" in drawingChild) {
      drawingChild.smartArt.runProperties = runProperties;
    } else if ("unsupportedDrawing" in drawingChild) {
      drawingChild.unsupportedDrawing.runProperties = runProperties;
    }
  }
  if (additionRsid) {
    if ("picture" in drawingChild) drawingChild.picture.additionRsid = additionRsid;
    else if ("wpsShape" in drawingChild) drawingChild.wpsShape.additionRsid = additionRsid;
    else if ("wpgGroup" in drawingChild) drawingChild.wpgGroup.additionRsid = additionRsid;
    else if ("wordprocessingCanvas" in drawingChild)
      drawingChild.wordprocessingCanvas.additionRsid = additionRsid;
    else if ("chart" in drawingChild) drawingChild.chart.additionRsid = additionRsid;
    else if ("chartEx" in drawingChild) drawingChild.chartEx.additionRsid = additionRsid;
    else if ("smartArt" in drawingChild) drawingChild.smartArt.additionRsid = additionRsid;
    else if ("unsupportedDrawing" in drawingChild)
      drawingChild.unsupportedDrawing.additionRsid = additionRsid;
  }
  if (deletionRsid) {
    if ("picture" in drawingChild) drawingChild.picture.deletionRsid = deletionRsid;
    else if ("wpsShape" in drawingChild) drawingChild.wpsShape.deletionRsid = deletionRsid;
    else if ("wpgGroup" in drawingChild) drawingChild.wpgGroup.deletionRsid = deletionRsid;
    else if ("wordprocessingCanvas" in drawingChild)
      drawingChild.wordprocessingCanvas.deletionRsid = deletionRsid;
    else if ("chart" in drawingChild) drawingChild.chart.deletionRsid = deletionRsid;
    else if ("chartEx" in drawingChild) drawingChild.chartEx.deletionRsid = deletionRsid;
    else if ("smartArt" in drawingChild) drawingChild.smartArt.deletionRsid = deletionRsid;
  }
  if (runPropertiesRsid) {
    if ("picture" in drawingChild) drawingChild.picture.runPropertiesRsid = runPropertiesRsid;
    else if ("wpsShape" in drawingChild)
      drawingChild.wpsShape.runPropertiesRsid = runPropertiesRsid;
    else if ("wpgGroup" in drawingChild)
      drawingChild.wpgGroup.runPropertiesRsid = runPropertiesRsid;
    else if ("wordprocessingCanvas" in drawingChild)
      drawingChild.wordprocessingCanvas.runPropertiesRsid = runPropertiesRsid;
    else if ("chart" in drawingChild) drawingChild.chart.runPropertiesRsid = runPropertiesRsid;
    else if ("chartEx" in drawingChild) drawingChild.chartEx.runPropertiesRsid = runPropertiesRsid;
    else if ("smartArt" in drawingChild)
      drawingChild.smartArt.runPropertiesRsid = runPropertiesRsid;
    else if ("unsupportedDrawing" in drawingChild)
      drawingChild.unsupportedDrawing.runPropertiesRsid = runPropertiesRsid;
  }
  // A run-level empty element (Word's pagination hint) sharing the drawing's
  // run — carried on the drawing options and emitted before the drawing.
  if (lastRenderedPageBreak) {
    if ("picture" in drawingChild) {
      drawingChild.picture.lastRenderedPageBreak = true;
    } else if ("wpsShape" in drawingChild) {
      drawingChild.wpsShape.lastRenderedPageBreak = true;
    } else if ("wpgGroup" in drawingChild) {
      drawingChild.wpgGroup.lastRenderedPageBreak = true;
    } else if ("wordprocessingCanvas" in drawingChild) {
      drawingChild.wordprocessingCanvas.lastRenderedPageBreak = true;
    } else if ("chart" in drawingChild) {
      drawingChild.chart.lastRenderedPageBreak = true;
    } else if ("chartEx" in drawingChild) {
      drawingChild.chartEx.lastRenderedPageBreak = true;
    } else if ("smartArt" in drawingChild) {
      drawingChild.smartArt.lastRenderedPageBreak = true;
    } else if ("unsupportedDrawing" in drawingChild) {
      drawingChild.unsupportedDrawing.lastRenderedPageBreak = true;
    }
  }
  return drawingChild;
}

/**
 * Parse the children of a track-change wrapper (w:ins/w:del/w:moveFrom/w:moveTo):
 * runs (reference runs keep their shape as run-children), the comment range
 * markers Word anchors directly inside the wrapper, complete field chains,
 * drawings (wps shapes and text boxes inserted as revisions), and nested
 * same-family wrappers.
 */
function parseTrackChangeRuns(el: Element, ctx: DocxReadContext): TrackChangeChild[] {
  const out: TrackChangeChild[] = [];
  const fieldState = initialFieldRunState();
  for (const sub of el.elements ?? []) {
    if (sub.name === "w:r") {
      const fed = feedFieldRun(sub, fieldState, ctx);
      if (fed.consumed) {
        if (fed.child) out.push(fed.child as TrackChangeChild);
        continue;
      }
      const drawingChildren = parseDrawingRunChildren(sub, ctx);
      if (drawingChildren.length > 0) {
        out.push(...(drawingChildren as TrackChangeChild[]));
        continue;
      }
      const parsed = parseRun(sub, ctx);
      const runOpts = parsedRunToOptions(parsed);
      if (typeof runOpts === "object" && isTrackedCommentReferenceRun(runOpts)) {
        const { commentReference, properties, additionRsid, runPropertiesRsid, deletionRsid } =
          runOpts;
        out.push({
          ...properties,
          ...(additionRsid ? { additionRsid } : {}),
          ...(runPropertiesRsid ? { runPropertiesRsid } : {}),
          ...(deletionRsid ? { deletionRsid } : {}),
          children: [{ commentReference }],
        });
        continue;
      } else if (!("commentReference" in runOpts)) {
        out.push(runOpts);
      }
    } else if (sub.name === "w:commentRangeStart") {
      const m = parseMarkupRangeOptions(sub, ctx.currentPart);
      if (m) out.push({ commentRangeStart: m });
    } else if (sub.name === "w:commentRangeEnd") {
      const m = parseMarkupRangeOptions(sub, ctx.currentPart);
      if (m) out.push({ commentRangeEnd: m });
    } else if (sub.name === "w:bookmarkStart" || sub.name === "w:bookmarkEnd") {
      const marker =
        sub.name === "w:bookmarkStart"
          ? parseBookmarkStartOptions(sub, ctx.currentPart)
          : parseBookmarkEndOptions(sub, ctx.currentPart);
      if (marker)
        out.push(
          sub.name === "w:bookmarkStart"
            ? ({ bookmarkStart: marker } as TrackChangeChild)
            : { bookmarkEnd: marker },
        );
    } else if (sub.name === "w:proofErr") {
      // Spell/grammar markers can sit inside the wrapper (CT_RunTrackChange
      // accepts the full range-markup group).
      const type = attr(sub, "w:type");
      if (
        type === "spellStart" ||
        type === "spellEnd" ||
        type === "gramStart" ||
        type === "gramEnd"
      ) {
        const stage = activeFieldState(fieldState);
        const members = stage.collectingResult ? stage.resultMembers : stage.instructionMembers;
        if (stage.kind) members.push({ proofErr: type });
        else out.push({ proofErr: type });
      }
    } else if (sub.name === "w:hyperlink") {
      const nested = parseHyperlinkChild(sub, ctx);
      if (nested && "hyperlink" in nested) out.push(nested as TrackChangeChild);
    } else if (sub.name === "w:sdt") {
      const sdtPr = findChild(sub, "w:sdtPr");
      const properties = sdtPr
        ? { ...parseSdtProperties(sdtPr, ctx), propertiesPresent: true }
        : { propertiesPresent: false };
      const sdtEndPr = findChild(sub, "w:sdtEndPr");
      const endRPr = sdtEndPr ? findChild(sdtEndPr, "w:rPr") : undefined;
      const sdt: SdtRunOptions = { properties };
      const children = parseRunLevelChildren(findChild(sub, "w:sdtContent")?.elements, ctx, {
        incompleteFields: "raw",
      });
      if (children.length > 0) sdt.children = children;
      if (sdtEndPr) {
        const endProperties = endRPr ? parseRunProperties(endRPr) : {};
        sdt.endProperties = endProperties;
      }
      out.push({ sdt });
    } else if (sub.name === "w:ins" || sub.name === "w:del") {
      // Nested wrappers (w:ins > w:del) — recurse with the same shape.
      const nested = parseTrackChangeRuns(sub, ctx);
      if (nested.length > 0) {
        const meta = {
          id: attrNum(sub, "w:id") ?? 0,
          author: attr(sub, "w:author") ?? "",
          date: attr(sub, "w:date"),
        };
        out.push(
          sub.name === "w:ins"
            ? { insertion: { ...meta, children: nested } }
            : { deletion: { ...meta, children: nested } },
        );
      }
    }
  }
  return [...out, ...incompleteTrackChangeFieldChildren(fieldState)];
}

function collectRunText(el: Element): string {
  let text = "";
  // w:delText inside a deletion wrapper carries the same character data as
  // w:t — a deleted field's cached result is spelled that way.
  for (const c of el.elements ?? []) {
    if (c.name === "w:t" || c.name === "w:delText") text += textOf(c);
  }
  return text;
}

/**
 * Parse the inline children of a smartTag/customXml container (recursive).
 *
 * Handles runs and nested smartTag/customXml. Form fields, hyperlinks and
 * other paragraph-level constructs are rare inside these containers and are
 * not handled here.
 */
function parseContainerChildren(el: Element, ctx: DocxReadContext): ParagraphChild[] {
  const children: ParagraphChild[] = [];
  for (const sub of el.elements ?? []) {
    switch (sub.name) {
      case "w:r": {
        const run = parsedRunToOptions(parseRun(sub, ctx));
        if (Object.keys(run).length === 0) {
          const bareText = textOf(sub).trim();
          if (bareText !== "") {
            children.push({ text: bareText });
            break;
          }
        }
        children.push(run);
        break;
      }
      case "w:smartTag": {
        const smartTag = parseSmartTagInline(sub, ctx);
        if (smartTag) children.push({ smartTag });
        break;
      }
      case "w:customXml": {
        const customXml = parseCustomXmlInline(sub, ctx);
        if (customXml) children.push({ customXml });
        break;
      }
      case "w:dir":
      case "w:bdo": {
        const val = attr(sub, "w:val");
        if (val !== "ltr" && val !== "rtl") break;
        const container: DirInlineOptions = { val };
        const nested = parseContainerChildren(sub, ctx);
        if (nested.length > 0) container.children = nested;
        children.push(sub.name === "w:dir" ? { dir: container } : { bdo: container });
        break;
      }
      case "w:hyperlink": {
        const child = parseHyperlinkChild(sub, ctx);
        if (child && isHyperlinkChild(child)) children.push(child);
        break;
      }
      case "w:proofErr": {
        const type = attr(sub, "w:type");
        if (
          type === "spellStart" ||
          type === "spellEnd" ||
          type === "gramStart" ||
          type === "gramEnd"
        )
          children.push({ proofErr: type });
        break;
      }
      case "w:bookmarkStart": {
        const marker = parseBookmarkStartOptions(sub, ctx.currentPart);
        if (marker) children.push({ bookmarkStart: marker });
        break;
      }
      case "w:bookmarkEnd": {
        const marker = parseBookmarkEndOptions(sub, ctx.currentPart);
        if (marker) children.push({ bookmarkEnd: marker });
        break;
      }
      case "w:commentRangeStart": {
        const marker = parseMarkupRangeOptions(sub, ctx.currentPart);
        if (marker) children.push({ commentRangeStart: marker });
        break;
      }
      case "w:commentRangeEnd": {
        const marker = parseMarkupRangeOptions(sub, ctx.currentPart);
        if (marker) children.push({ commentRangeEnd: marker });
        break;
      }
      case "w:permStart": {
        const id = attrNum(sub, "w:id");
        if (id === undefined) break;
        const start: Extract<ParagraphChild, { permStart: unknown }>["permStart"] = { id };
        const editor = attr(sub, "w:ed");
        if (editor !== undefined) start.editor = editor;
        const editGroup = attr(sub, "w:edGrp");
        if (editGroup !== undefined) {
          start.editGroup = editGroup as typeof start.editGroup;
        }
        const colFirst = attrNum(sub, "w:colFirst");
        if (colFirst !== undefined) start.colFirst = colFirst;
        const colLast = attrNum(sub, "w:colLast");
        if (colLast !== undefined) start.colLast = colLast;
        children.push({ permStart: start });
        break;
      }
      case "w:permEnd": {
        const id = attrNum(sub, "w:id");
        if (id !== undefined) children.push({ permEnd: id });
        break;
      }
      default:
        break;
    }
  }
  return children;
}

/** Parse a w:smartTag element into its ParagraphChild form. */
function parseSmartTagInline(el: Element, ctx: DocxReadContext): SmartTagInlineOptions | undefined {
  const element = attr(el, "w:element");
  if (!element) return undefined;
  const st: SmartTagInlineOptions = { element };
  const uri = attr(el, "w:uri");
  if (uri) st.uri = uri;
  const pr = findChild(el, "w:smartTagPr");
  if (pr) {
    const props: Array<{ uri?: string; name: string; val: string }> = [];
    for (const a of pr.elements ?? []) {
      if (a.name !== "w:attr") continue;
      const prop: { uri?: string; name: string; val: string } = {
        name: attr(a, "w:name") ?? "",
        val: attr(a, "w:val") ?? "",
      };
      const auri = attr(a, "w:uri");
      if (auri) prop.uri = auri;
      props.push(prop);
    }
    if (props.length > 0) st.properties = props;
  }
  const content = parseContainerChildren(el, ctx);
  if (content.length > 0) st.children = content;
  return st;
}

/** Parse an inline w:customXml element into its ParagraphChild form. */
function parseCustomXmlInline(
  el: Element,
  ctx: DocxReadContext,
): CustomXmlInlineOptions | undefined {
  const element = attr(el, "w:element");
  if (!element) return undefined;
  const cx: CustomXmlInlineOptions = { element };
  const uri = attr(el, "w:uri");
  if (uri) cx.uri = uri;
  const pr = findChild(el, "w:customXmlPr");
  if (pr) {
    const parsed = parseCustomXmlProperties(pr);
    if (parsed.placeholder !== undefined || parsed.attributes !== undefined) {
      cx.properties = parsed;
    }
  }
  const content = parseContainerChildren(el, ctx);
  if (content.length > 0) cx.children = content;
  return cx;
}

/** Parse a move-revision range start (w:moveFromRangeStart / w:moveToRangeStart). */
function parseMoveRangeStart(
  el: Element,
  part = "word/document.xml",
): MoveRangeStartOptions | null {
  const id = parseMarkerId(el, el.name ?? "w:moveFromRangeStart", part);
  if (id === undefined) return null;
  const m: Partial<MoveRangeStartOptions> = { id };
  const name = attr(el, "w:name");
  if (name !== undefined) m.name = name;
  const author = attr(el, "w:author");
  if (author !== undefined) m.author = author;
  const date = attr(el, "w:date");
  if (date !== undefined) m.date = date;
  const disp = attr(el, "w:displacedByCustomXml");
  if (disp === "next" || disp === "prev") m.displacedByCustomXml = disp as DisplacedByCustomXml;
  const colFirst = attrNum(el, "w:colFirst");
  if (colFirst !== undefined) m.colFirst = colFirst;
  const colLast = attrNum(el, "w:colLast");
  if (colLast !== undefined) m.colLast = colLast;
  return m as MoveRangeStartOptions;
}

/** Parse a customXml range start (Ins/Del/MoveFrom/MoveTo). */
function parseCustomXmlRangeStart(
  el: Element,
  part = "word/document.xml",
): { id: number; author?: string; date?: string } | null {
  const id = parseMarkerId(el, el.name ?? "w:customXmlRangeStart", part);
  if (id === undefined) return null;
  const m: { id: number; author?: string; date?: string } = { id };
  const author = attr(el, "w:author");
  if (author !== undefined) m.author = author;
  const date = attr(el, "w:date");
  if (date !== undefined) m.date = date;
  return m;
}

/** Parse a CT_MarkupRange end marker (id + displacedByCustomXml). */
function parseMarkupRangeOptions(
  el: Element,
  part = "word/document.xml",
): MarkupRangeOptions | undefined {
  const id = parseMarkerId(el, el.name ?? "w:markupRange", part);
  if (id === undefined) return undefined;
  const m: Partial<MarkupRangeOptions> = { id };
  const disp = attr(el, "w:displacedByCustomXml");
  if (disp === "next" || disp === "prev") m.displacedByCustomXml = disp as DisplacedByCustomXml;
  return m as MarkupRangeOptions;
}

/**
 * Parse a w:p element into ParagraphOptions.
 */
/** Serialize a w:r's w:rPr child verbatim (or undefined when the run has none). */
export function runRPrXml(run: Element): string | undefined {
  const rPr = findChild(run, "w:rPr");
  return rPr ? stringifyElement(rPr) : undefined;
}

/** Parse a w:bookmarkStart element (id + name + table-column extents). */
export function parseBookmarkStartOptions(
  el: Element,
  part = "word/document.xml",
): BookmarkStartOptions | undefined {
  const id = parseMarkerId(el, "w:bookmarkStart", part);
  const name = attr(el, "w:name");
  if (id === undefined || !name) return undefined;
  const bookmarkStart: Partial<BookmarkStartOptions> = { id, name };
  const disp = attr(el, "w:displacedByCustomXml");
  if (disp === "next" || disp === "prev")
    bookmarkStart.displacedByCustomXml = disp as DisplacedByCustomXml;
  const colFirst = attrNum(el, "w:colFirst");
  if (colFirst !== undefined) bookmarkStart.colFirst = colFirst;
  const colLast = attrNum(el, "w:colLast");
  if (colLast !== undefined) bookmarkStart.colLast = colLast;
  return bookmarkStart as BookmarkStartOptions;
}

/** Parse a w:bookmarkEnd element (id + displacedByCustomXml). */
export function parseBookmarkEndOptions(
  el: Element,
  part = "word/document.xml",
): MarkupRangeOptions | undefined {
  const id = parseMarkerId(el, "w:bookmarkEnd", part);
  if (id === undefined) return undefined;
  const bookmarkEnd: Partial<MarkupRangeOptions> = { id };
  const disp = attr(el, "w:displacedByCustomXml");
  if (disp === "next" || disp === "prev")
    bookmarkEnd.displacedByCustomXml = disp as DisplacedByCustomXml;
  return bookmarkEnd as MarkupRangeOptions;
}

function parseMarkerId(el: Element, elementName: string, part: string): number | undefined {
  const raw = attr(el, "w:id");
  if (raw === undefined) return undefined;
  const id = Number(raw);
  if (!Number.isSafeInteger(id) || id < 0) {
    throw new DocxParseError(
      `Invalid ${elementName} marker id`,
      part,
      `/${elementName}/@w:id`,
      "unsafe-marker-id",
    );
  }
  return id;
}

/**
 * Parses one w:hyperlink element into a `{ hyperlink }` child, or null when it
 * carries no content. CT_Hyperlink content is EG_PContent, so a hyperlink may
 * nest another hyperlink (pandoc emits these for nested anchors) — recursion
 * handles both levels through the same shape.
 */
function parseHyperlinkChild(child: Element, ctx: DocxReadContext): ParagraphChild | null {
  const hl: Partial<HyperlinkInlineOptions> = {};
  const rId = attr(child, "r:id");
  if (rId !== undefined) {
    // rId numbering is per-part: a hyperlink inside a footnote/header
    // resolves against that part's own rels first, then document.xml's.
    const target =
      ctx.docx.partRefs.partHyperlinks.get(ctx.currentPart)?.get(rId) ??
      ctx.docx.partRefs.hyperlinks.get(rId);
    if (target) hl.url = target;
    if (target) hl.sourceRelationshipId = rId;
  }
  const anchor = attr(child, "w:anchor");
  if (anchor !== undefined) hl.anchor = anchor;
  const tooltip = attr(child, "w:tooltip");
  if (tooltip !== undefined) hl.tooltip = tooltip;
  const tgtFrame = attr(child, "w:tgtFrame");
  if (tgtFrame !== undefined) hl.targetFrame = tgtFrame;
  const docLocation = attr(child, "w:docLocation");
  if (docLocation !== undefined) hl.docLocation = docLocation;
  const history = attrBool(child, "w:history");
  if (history !== undefined) hl.history = history;

  const linkRuns: (RunOptions | string | ParagraphChild)[] = [];
  // Field accumulator scoped to the hyperlink: Word nests complex fields
  // inside link content (e.g. the PAGEREF page number of a TOC entry),
  // fully closed within the hyperlink element.
  const fieldState = initialFieldRunState();
  for (const sub of child.elements ?? []) {
    if (sub.name === "w:r") {
      const fed = feedFieldRun(sub, fieldState, ctx);
      if (fed.consumed) {
        if (fed.child) linkRuns.push(fed.child);
        continue;
      }
      // Drawing runs inside hyperlinks (image links) resolve through the
      // same paragraph-child extraction — parseRun skips w:drawing.
      const drawingChildren = parseDrawingRunChildren(sub, ctx);
      if (drawingChildren.length > 0) {
        linkRuns.push(...drawingChildren);
        continue;
      }
      const parsed = parseRun(sub, ctx);
      const runOpts = parsedRunToOptions(parsed);
      // parsedRunToOptions returns { commentReference } for pure
      // comment-reference runs; hyperlink children are runs, so skip it.
      if (Object.keys(runOpts).length > 0) linkRuns.push(runOpts);
    } else if (sub.name === "w:commentRangeStart" || sub.name === "w:commentRangeEnd") {
      const marker =
        sub.name === "w:commentRangeStart"
          ? parseMarkupRangeOptions(sub, ctx.currentPart)
          : parseMarkupRangeOptions(sub, ctx.currentPart);
      if (marker)
        linkRuns.push(
          sub.name === "w:commentRangeStart"
            ? { commentRangeStart: marker }
            : { commentRangeEnd: marker },
        );
    } else if (sub.name === "w:ins" || sub.name === "w:del") {
      const nested = parseTrackChangeRuns(sub, ctx);
      if (nested.length > 0) {
        const revision = {
          id: attrNum(sub, "w:id") ?? 0,
          author: attr(sub, "w:author") ?? "",
          date: attr(sub, "w:date"),
        };
        linkRuns.push(
          sub.name === "w:ins"
            ? { insertion: { ...revision, children: nested } }
            : { deletion: { ...revision, children: nested } },
        );
      }
    } else if (sub.name === "w:hyperlink") {
      const nested = parseHyperlinkChild(sub, ctx);
      if (nested) linkRuns.push(nested);
    } else if (sub.name === "w:bookmarkStart") {
      // Bookmarks may open inside link content (Word parks _GoBack on the
      // last edit position, often an image link run).
      const bookmarkStart = parseBookmarkStartOptions(sub, ctx.currentPart);
      if (bookmarkStart) {
        const stage = activeFieldState(fieldState);
        const members = stage.collectingResult ? stage.resultMembers : stage.instructionMembers;
        if (stage.kind) members.push({ bookmarkStart });
        else linkRuns.push({ bookmarkStart });
      }
    } else if (sub.name === "w:bookmarkEnd") {
      const bookmarkEnd = parseBookmarkEndOptions(sub, ctx.currentPart);
      if (bookmarkEnd) {
        const stage = activeFieldState(fieldState);
        const members = stage.collectingResult ? stage.resultMembers : stage.instructionMembers;
        if (stage.kind) members.push({ bookmarkEnd });
        else linkRuns.push({ bookmarkEnd });
      }
    } else if (sub.name === "w:proofErr") {
      const type = attr(sub, "w:type");
      const proofType =
        type === "spellStart" || type === "spellEnd" || type === "gramStart" || type === "gramEnd"
          ? type
          : undefined;
      if (proofType) {
        const stage = activeFieldState(fieldState);
        const member = { proofErr: proofType } as const;
        const members = stage.collectingResult ? stage.resultMembers : stage.instructionMembers;
        if (stage.kind) members.push(member);
        else linkRuns.push({ proofErr: type } as ParagraphChild);
      }
    }
  }
  hl.children = linkRuns;
  return { hyperlink: hl };
}

/**
 * Parse run-level children shared by paragraphs and inline-SDT content.
 * Includes the field accumulator that collapses form/complex fields spanning
 * multiple w:r runs into a single child.
 */
function parseRunLevelChildren(
  elements: Element[] | undefined,
  ctx: DocxReadContext,
  options: { incompleteFields?: "raw" | "defer"; preserveOrphanMarkers?: boolean } = {},
): ParagraphChild[] {
  const childList: ParagraphChild[] = [];

  // Field accumulator — see feedFieldRun. The whole field collapses to a
  // single child.
  const fieldState = initialFieldRunState();
  fieldState.preserveOrphanMarkers = options.preserveOrphanMarkers !== false;

  for (const child of elements ?? []) {
    switch (child.name) {
      case "w:pPr":
        break;
      case "w:r": {
        // Field: fldChar markers + the instrText/result runs between them.
        const fed = feedFieldRun(child, fieldState, ctx);
        if (fed.consumed) {
          if (fed.child) childList.push(fed.child);
          break;
        }

        const drawingChildren = parseDrawingRunChildren(child, ctx);
        if (drawingChildren.length > 0) {
          childList.push(...drawingChildren);
          break;
        }
        const parsed = parseRun(child, ctx);
        childList.push(parsedRunToOptions(parsed));
        break;
      }
      case "w:hyperlink": {
        const hlChild = parseHyperlinkChild(child, ctx);
        if (!hlChild || !isHyperlinkChild(hlChild)) break;
        if (fieldState.kind) {
          const stage = activeFieldState(fieldState);
          (stage.collectingResult ? stage.resultMembers : stage.instructionMembers).push(hlChild);
        } else {
          childList.push(hlChild);
        }
        break;
      }
      case "w:tbl":
        childList.push({ table: tableDesc.parse(child, ctx) as TableOptions });
        break;
      case "w:br": {
        const type = attr(child, "w:type");
        const clear = attr(child, "w:clear");
        childList.push(
          type === "page"
            ? { pageBreak: true }
            : type === "column"
              ? { columnBreak: true }
              : {
                  blockBreak: {
                    count: 1,
                    ...(type ? { type: type as BreakOptions["type"] } : {}),
                    ...(clear ? { clear: clear as BreakOptions["clear"] } : {}),
                  },
                },
        );
        break;
      }
      case "w:subDoc": {
        // EG_PContent member — the sub-document insertion sits inside a
        // paragraph like a hyperlink, never directly under w:body
        childList.push({ subDoc: parseSubDoc(child, ctx) });
        break;
      }
      case "w:bookmarkStart": {
        const bookmarkStart = parseBookmarkStartOptions(child, ctx.currentPart);
        if (bookmarkStart) {
          const stage = activeFieldState(fieldState);
          const members = stage.collectingResult ? stage.resultMembers : stage.instructionMembers;
          if (stage.kind) members.push({ bookmarkStart });
          else childList.push({ bookmarkStart });
        }
        break;
      }
      case "w:bookmarkEnd": {
        const bookmarkEnd = parseBookmarkEndOptions(child, ctx.currentPart);
        if (bookmarkEnd) {
          const stage = activeFieldState(fieldState);
          const members = stage.collectingResult ? stage.resultMembers : stage.instructionMembers;
          if (stage.kind) members.push({ bookmarkEnd });
          else childList.push({ bookmarkEnd });
        }
        break;
      }
      case "w:commentRangeStart": {
        const m = parseMarkupRangeOptions(child, ctx.currentPart);
        if (m) childList.push({ commentRangeStart: m });
        break;
      }
      case "w:commentRangeEnd": {
        const m = parseMarkupRangeOptions(child, ctx.currentPart);
        if (m) childList.push({ commentRangeEnd: m });
        break;
      }
      case "w:commentReference": {
        const id = attrNum(child, "w:id");
        if (id !== undefined) childList.push({ commentReference: id });
        break;
      }
      case "m:oMath": {
        childList.push({ math: { equations: [parseMathChildren(child, ctx)] } });
        break;
      }
      case "m:oMathPara": {
        childList.push({ math: { ...parseMathParagraph(child, ctx), display: true } });
        break;
      }
      case "w:ins": {
        const children = parseTrackChangeRuns(child, ctx);
        const revision = {
          insertion: {
            id: attrNum(child, "w:id") ?? 0,
            author: attr(child, "w:author") ?? "",
            date: attr(child, "w:date"),
            children,
          },
        };
        const stage = activeFieldState(fieldState);
        const members = stage.collectingResult ? stage.resultMembers : stage.instructionMembers;
        if (stage.kind) {
          const memberCount = members.length;
          members.push(revision);
          const wrapperState = initialFieldRunState();
          wrapperState.preserveOrphanMarkers = true;
          for (const run of (child.elements ?? []).filter((item) => item.name === "w:r")) {
            feedFieldRun(run, wrapperState, ctx);
          }
          if (wrapperState.kind) {
            for (const run of (child.elements ?? []).filter((item) => item.name === "w:r")) {
              feedFieldRun(run, stage, ctx);
            }
          }
          const trackedRun = members.at(-1);
          if (members.length > memberCount && trackedRun && "run" in trackedRun) members.pop();
        } else if (children.length > 0) {
          childList.push(revision);
        }
        break;
      }
      case "w:del": {
        const children = parseTrackChangeRuns(child, ctx);
        const revision = {
          deletion: {
            id: attrNum(child, "w:id") ?? 0,
            author: attr(child, "w:author") ?? "",
            date: attr(child, "w:date"),
            children,
          },
        };
        const stage = activeFieldState(fieldState);
        const members = stage.collectingResult ? stage.resultMembers : stage.instructionMembers;
        if (stage.kind) {
          const memberCount = members.length;
          members.push(revision);
          const wrapperState = initialFieldRunState();
          wrapperState.preserveOrphanMarkers = true;
          for (const run of (child.elements ?? []).filter((item) => item.name === "w:r")) {
            feedFieldRun(run, wrapperState, ctx);
          }
          if (wrapperState.kind) {
            for (const run of (child.elements ?? []).filter((item) => item.name === "w:r")) {
              feedFieldRun(run, stage, ctx);
            }
          }
          const trackedRun = members.at(-1);
          if (members.length > memberCount && trackedRun && "run" in trackedRun) members.pop();
        } else if (children.length > 0) {
          childList.push(revision);
        }
        break;
      }
      case "w:moveFrom": {
        const children = parseTrackChangeRuns(child, ctx);
        const revision = {
          movedFrom: {
            id: attrNum(child, "w:id") ?? 0,
            author: attr(child, "w:author") ?? "",
            date: attr(child, "w:date"),
            children,
          },
        };
        const stage = activeFieldState(fieldState);
        const members = stage.collectingResult ? stage.resultMembers : stage.instructionMembers;
        if (stage.kind) {
          const memberCount = members.length;
          members.push(revision);
          const wrapperState = initialFieldRunState();
          wrapperState.preserveOrphanMarkers = true;
          for (const run of (child.elements ?? []).filter((item) => item.name === "w:r")) {
            feedFieldRun(run, wrapperState, ctx);
          }
          if (wrapperState.kind) {
            for (const run of (child.elements ?? []).filter((item) => item.name === "w:r")) {
              feedFieldRun(run, stage, ctx);
            }
          }
          const trackedRun = members.at(-1);
          if (members.length > memberCount && trackedRun && "run" in trackedRun) members.pop();
        } else if (children.length > 0) {
          childList.push(revision);
        }
        break;
      }
      case "w:moveTo": {
        const children = parseTrackChangeRuns(child, ctx);
        const revision = {
          movedTo: {
            id: attrNum(child, "w:id") ?? 0,
            author: attr(child, "w:author") ?? "",
            date: attr(child, "w:date"),
            children,
          },
        };
        const stage = activeFieldState(fieldState);
        const members = stage.collectingResult ? stage.resultMembers : stage.instructionMembers;
        if (stage.kind) {
          const memberCount = members.length;
          members.push(revision);
          const wrapperState = initialFieldRunState();
          wrapperState.preserveOrphanMarkers = true;
          for (const run of (child.elements ?? []).filter((item) => item.name === "w:r")) {
            feedFieldRun(run, wrapperState, ctx);
          }
          if (wrapperState.kind) {
            for (const run of (child.elements ?? []).filter((item) => item.name === "w:r")) {
              feedFieldRun(run, stage, ctx);
            }
          }
          const trackedRun = members.at(-1);
          if (members.length > memberCount && trackedRun && "run" in trackedRun) members.pop();
        } else if (children.length > 0) {
          childList.push(revision);
        }
        break;
      }
      case "w:fldSimple": {
        const instruction = attr(child, "w:instr");
        if (instruction) {
          const sf: {
            instruction: string;
            cachedValue?: string;
            cachedValuePreserveSpace?: boolean;
            cachedInstructionText?: string;
            cachedInstructionTextPreserveSpace?: boolean;
            cachedInstructionRPrXml?: string;
            cachedRuns?: ParagraphChild[];
            fieldLock?: boolean;
            dirty?: boolean;
          } = { instruction };
          // cachedValue: concatenate the result-run <w:t> text (one or more
          // <w:r> children between the fldSimple tags). Some files carry the
          // result as w:instrText runs instead (nested-field expansions) —
          // those count as display text too.
          let cachedValue = "";
          let cachedValuePreserveSpace = false;
          let cachedInstructionText = "";
          let cachedInstructionTextPreserveSpace = false;
          const cachedRunEls: Element[] = [];
          for (const sub of child.elements ?? []) {
            if (sub.name === "w:r") {
              cachedValue += collectRunText(sub);
              const firstText = (sub.elements ?? []).find(
                (rc) => rc.name === "w:t" || rc.name === "w:delText",
              );
              if (firstText && attr(firstText, "xml:space") === "preserve") {
                cachedValuePreserveSpace = true;
              }
              for (const rc of sub.elements ?? []) {
                if (rc.name === "w:instrText") {
                  cachedInstructionText += textOf(rc) ?? "";
                  cachedInstructionTextPreserveSpace ||= attr(rc, "xml:space") === "preserve";
                }
              }
              cachedRunEls.push(sub);
            }
          }
          // Mark the cached value present when any run sat inside the field —
          // empty-text results still round-trip their run.
          if (cachedRunEls.length > 0) sf.cachedValue = cachedValue;
          const instructionRunEl =
            cachedRunEls.length === 1
              ? (findChild(cachedRunEls[0]!, "w:instrText") ??
                findChild(cachedRunEls[0]!, "w:delInstrText"))
              : undefined;
          const plainInstructionRun =
            instructionRunEl !== undefined &&
            attr(cachedRunEls[0]!, "w:rsidR") === undefined &&
            attr(cachedRunEls[0]!, "w:rsidRPr") === undefined &&
            attr(cachedRunEls[0]!, "w:rsidDel") === undefined;
          const cachedRunHasIdentity = cachedRunEls.some(
            (run) =>
              attr(run, "w:rsidR") !== undefined ||
              attr(run, "w:rsidRPr") !== undefined ||
              attr(run, "w:rsidDel") !== undefined,
          );
          if (plainInstructionRun) {
            sf.cachedInstructionText = cachedInstructionText;
            sf.cachedInstructionTextPreserveSpace = cachedInstructionTextPreserveSpace;
            const instructionRPrXml = runRPrXml(cachedRunEls[0]!);
            if (instructionRPrXml) sf.cachedInstructionRPrXml = instructionRPrXml;
          }
          sf.cachedValuePreserveSpace = cachedValuePreserveSpace;
          // A plain template emits one bare text run or one instruction run.
          // Nested fldChar fields, per-run rsids, rPr, and mixed cached runs
          // must remain typed children instead of collapsing to cachedValue.
          if (
            !plainInstructionRun &&
            (cachedRunHasIdentity || !isPlainFieldRuns(cachedRunEls, undefined, ["w:t"]))
          ) {
            sf.cachedRuns = parseRunLevelChildren(child.elements, ctx);
          }
          const sfLock = attrBool(child, "w:fldLock");
          if (sfLock !== undefined) sf.fieldLock = sfLock;
          const sfDirty = attrBool(child, "w:dirty");
          if (sfDirty !== undefined) sf.dirty = sfDirty;
          if (
            fieldState.kind === "complex" &&
            fieldState.depth === 1 &&
            !fieldState.collectingResult
          ) {
            fieldState.instructionMembers.push({ simpleField: sf });
          } else {
            childList.push({ simpleField: sf });
          }
        }
        break;
      }
      case "w:smartTag": {
        const st = parseSmartTagInline(child, ctx);
        if (st) childList.push({ smartTag: st });
        break;
      }
      case "w:customXml": {
        const cx = parseCustomXmlInline(child, ctx);
        if (cx) childList.push({ customXml: cx });
        break;
      }
      // ── Bidirectional containers (reuse the smartTag/customXml child parser) ──
      case "w:dir": {
        const val = attr(child, "w:val");
        if (val) {
          const dir: DirInlineOptions = { val: val as "ltr" | "rtl" };
          const content = parseContainerChildren(child, ctx);
          if (content.length > 0) dir.children = content;
          childList.push({ dir });
        }
        break;
      }
      case "w:bdo": {
        const val = attr(child, "w:val");
        if (val) {
          const bdo: DirInlineOptions = { val: val as "ltr" | "rtl" };
          const content = parseContainerChildren(child, ctx);
          if (content.length > 0) bdo.children = content;
          childList.push({ bdo });
        }
        break;
      }
      // ── Range markers: proof errors, positional tabs, permissions, revisions ──
      case "w:proofErr": {
        const type = attr(child, "w:type");
        if (
          type === "spellStart" ||
          type === "spellEnd" ||
          type === "gramStart" ||
          type === "gramEnd"
        ) {
          const members = fieldState.collectingResult
            ? fieldState.resultMembers
            : fieldState.instructionMembers;
          if (fieldState.kind) members.push({ proofErr: type });
          else childList.push({ proofErr: type });
        }
        break;
      }
      case "w:ptab": {
        const alignment = attr(child, "w:alignment");
        const leader = attr(child, "w:leader");
        const relativeTo = attr(child, "w:relativeTo");
        if (alignment !== undefined && leader !== undefined && relativeTo !== undefined) {
          childList.push({
            positionalTab: {
              alignment: alignment as PositionalTabOptions["alignment"],
              leader: leader as PositionalTabOptions["leader"],
              relativeTo: relativeTo as PositionalTabOptions["relativeTo"],
            },
          });
        }
        break;
      }
      case "w:permStart": {
        const id = attr(child, "w:id");
        if (id !== undefined) {
          const ps: PermStartInlineOptions = { id };
          const ed = attr(child, "w:ed");
          if (ed !== undefined) ps.editor = ed;
          const editGroup = attr(child, "w:edGrp");
          if (editGroup !== undefined) {
            ps.editGroup = editGroup as PermStartInlineOptions["editGroup"];
          }
          const colFirst = attrNum(child, "w:colFirst");
          if (colFirst !== undefined) ps.colFirst = colFirst;
          const colLast = attrNum(child, "w:colLast");
          if (colLast !== undefined) ps.colLast = colLast;
          childList.push({ permStart: ps });
        }
        break;
      }
      case "w:permEnd": {
        const id = attr(child, "w:id");
        if (id !== undefined) childList.push({ permEnd: id });
        break;
      }
      case "w:moveFromRangeStart": {
        const m = parseMoveRangeStart(child, ctx.currentPart);
        if (m) childList.push({ moveFromRangeStart: m });
        break;
      }
      case "w:moveFromRangeEnd": {
        const m = parseMarkupRangeOptions(child, ctx.currentPart);
        if (m) childList.push({ moveFromRangeEnd: m });
        break;
      }
      case "w:moveToRangeStart": {
        const m = parseMoveRangeStart(child, ctx.currentPart);
        if (m) childList.push({ moveToRangeStart: m });
        break;
      }
      case "w:moveToRangeEnd": {
        const m = parseMarkupRangeOptions(child, ctx.currentPart);
        if (m) childList.push({ moveToRangeEnd: m });
        break;
      }
      case "w:customXmlInsRangeStart": {
        const m = parseCustomXmlRangeStart(child, ctx.currentPart);
        if (m) childList.push({ customXmlInsRangeStart: m });
        break;
      }
      case "w:customXmlInsRangeEnd": {
        const id = parseMarkerId(child, "w:customXmlInsRangeEnd", ctx.currentPart);
        if (id !== undefined) childList.push({ customXmlInsRangeEnd: id });
        break;
      }
      case "w:customXmlDelRangeStart": {
        const m = parseCustomXmlRangeStart(child, ctx.currentPart);
        if (m) childList.push({ customXmlDelRangeStart: m });
        break;
      }
      case "w:customXmlDelRangeEnd": {
        const id = parseMarkerId(child, "w:customXmlDelRangeEnd", ctx.currentPart);
        if (id !== undefined) childList.push({ customXmlDelRangeEnd: id });
        break;
      }
      case "w:customXmlMoveFromRangeStart": {
        const m = parseCustomXmlRangeStart(child, ctx.currentPart);
        if (m) childList.push({ customXmlMoveFromRangeStart: m });
        break;
      }
      case "w:customXmlMoveFromRangeEnd": {
        const id = parseMarkerId(child, "w:customXmlMoveFromRangeEnd", ctx.currentPart);
        if (id !== undefined) childList.push({ customXmlMoveFromRangeEnd: id });
        break;
      }
      case "w:customXmlMoveToRangeStart": {
        const m = parseCustomXmlRangeStart(child, ctx.currentPart);
        if (m) childList.push({ customXmlMoveToRangeStart: m });
        break;
      }
      case "w:customXmlMoveToRangeEnd": {
        const id = parseMarkerId(child, "w:customXmlMoveToRangeEnd", ctx.currentPart);
        if (id !== undefined) childList.push({ customXmlMoveToRangeEnd: id });
        break;
      }
      case "w14:customXmlConflictInsRangeStart": {
        const m = parseCustomXmlRangeStart(child, ctx.currentPart);
        if (m) childList.push({ customXmlConflictInsRangeStart: m });
        break;
      }
      case "w14:customXmlConflictInsRangeEnd": {
        const id = parseMarkerId(child, "w14:customXmlConflictInsRangeEnd", ctx.currentPart);
        if (id !== undefined) childList.push({ customXmlConflictInsRangeEnd: id });
        break;
      }
      case "w14:customXmlConflictDelRangeStart": {
        const m = parseCustomXmlRangeStart(child, ctx.currentPart);
        if (m) childList.push({ customXmlConflictDelRangeStart: m });
        break;
      }
      case "w14:customXmlConflictDelRangeEnd": {
        const id = parseMarkerId(child, "w14:customXmlConflictDelRangeEnd", ctx.currentPart);
        if (id !== undefined) childList.push({ customXmlConflictDelRangeEnd: id });
        break;
      }
      case "w14:conflictIns": {
        const children = parseTrackChangeRuns(child, ctx);
        if (children.length > 0) {
          childList.push({
            conflictIns: {
              id: attrNum(child, "w:id") ?? 0,
              author: attr(child, "w:author") ?? "",
              date: attr(child, "w:date"),
              children,
            },
          });
        }
        break;
      }
      case "w14:conflictDel": {
        const children = parseTrackChangeRuns(child, ctx);
        if (children.length > 0) {
          childList.push({
            conflictDel: {
              id: attrNum(child, "w:id") ?? 0,
              author: attr(child, "w:author") ?? "",
              date: attr(child, "w:date"),
              children,
            },
          });
        }
        break;
      }
      case "w:sdt": {
        const sdtPr = findChild(child, "w:sdtPr");
        const properties = sdtPr
          ? { ...parseSdtProperties(sdtPr, ctx), propertiesPresent: true }
          : { propertiesPresent: false };
        // CT_SdtEndPr wraps its run properties in a w:rPr child
        const sdtEndPr = findChild(child, "w:sdtEndPr");
        const endRPr = sdtEndPr ? findChild(sdtEndPr, "w:rPr") : undefined;
        const endProperties = sdtEndPr ? (endRPr ? parseRunProperties(endRPr) : {}) : undefined;
        const sdtContent = findChild(child, "w:sdtContent");
        const sdtChildren = parseRunLevelChildren(sdtContent?.elements, ctx, {
          incompleteFields: "raw",
        });
        const sdt: SdtRunOptions = { properties };
        if (sdtChildren.length > 0) sdt.children = sdtChildren;
        if (endProperties) sdt.endProperties = endProperties;
        const sdtChild = { sdt };
        const stage = activeFieldState(fieldState);
        const members = stage.collectingResult ? stage.resultMembers : stage.instructionMembers;
        if (stage.kind) members.push(sdtChild);
        else childList.push(sdtChild);
        break;
      }
      default:
        break;
    }
  }

  // A complex field still open at the paragraph end spans paragraphs (e.g. a
  // TOC head paragraph carrying the first entry). Release the buffered result
  // runs as regular children — the aggregator rebuilds the outer field's
  // control chain from its own capture, and a field nested inside the result
  // (a PAGEREF page number) re-parses through the sub accumulator. The head
  // runs stay unreleased: re-emitting them here would duplicate the chain the
  // emit path injects.
  const incompleteField =
    options.incompleteFields === "defer" ? undefined : flushIncompleteFieldRunState(fieldState);
  if (
    options.incompleteFields === "defer" &&
    fieldState.kind === "complex" &&
    (fieldState.resultRunEls.length > 0 || fieldState.resultMembers.length > 0)
  ) {
    for (const member of fieldState.resultMembers) {
      childList.push("run" in member ? member.run : member);
    }
  }
  if (incompleteField) childList.push(incompleteField);

  return childList;
}

/** True when a child is a single-field `{ text: string }` run (simple text). */
function isTextOnlyRun(c: unknown): c is { text: string } {
  return typeof c === "object" && c !== null && "text" in c && Object.keys(c).length === 1;
}

export function parseParagraph(
  el: Element,
  ctx: DocxReadContext,
  options: { incompleteFields?: "raw" | "defer"; preserveOrphanMarkers?: boolean } = {},
): ParagraphOptions {
  const opts: Partial<ParagraphOptions> = {};

  // w:p element attributes: rsid family + w14:paraId/textId (hex string verbatim)
  const paraId = attr(el, "w14:paraId");
  if (paraId) opts.paraId = paraId;
  const textId = attr(el, "w14:textId");
  if (textId) opts.textId = textId;
  const editId = attr(el, "w14:editId");
  if (editId) opts.editId = editId;
  const noSpellErr = attr(el, "w14:noSpellErr");
  if (noSpellErr !== undefined) {
    opts.noSpellErr = noSpellErr === "1" || noSpellErr.toLowerCase() === "true";
  }
  const rsid = attr(el, "w:rsidR");
  if (rsid) opts.additionRsid = rsid;
  const defaultRunRsid = attr(el, "w:rsidRDefault");
  if (defaultRunRsid) opts.defaultRunRsid = defaultRunRsid;
  const propertiesRsid = attr(el, "w:rsidP");
  if (propertiesRsid) opts.propertiesRsid = propertiesRsid;
  const runPropertiesRsid = attr(el, "w:rsidRPr");
  if (runPropertiesRsid) opts.runPropertiesRsid = runPropertiesRsid;
  const deletionRsid = attr(el, "w:rsidDel");
  if (deletionRsid) opts.deletionRsid = deletionRsid;

  const pPr = findChild(el, "w:pPr");
  if (pPr) {
    Object.assign(opts, parseParagraphProperties(pPr, ctx));
  }

  const childList = parseRunLevelChildren(el.elements, ctx, options);

  // Simple text optimization: a single text-only run collapses into opts.text
  // (the canonical ParagraphOptions form). Multiple runs stay a children
  // array — run boundaries carry spelling-check and session-edit history and
  // must round-trip instead of merging into one emitted run.
  if (childList.length === 1) {
    const only = childList[0];
    if (isTextOnlyRun(only) && only.text) {
      opts.text = only.text;
      return opts as ParagraphOptions;
    }
  }
  if (childList.length > 0) {
    opts.children = childList;
  }

  return opts as ParagraphOptions;
}
