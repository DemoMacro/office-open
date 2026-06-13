/**
 * Body-level stringification for DOCX documents.
 *
 * Converts pure JSON options to XML strings for document body content.
 * Pure string concatenation — zero IXmlableObject, zero BaseXmlComponent.
 *
 * @module
 */

import { toUint8Array } from "@office-open/core";
import { uniqueId } from "@office-open/core";
import { hexColorValue, uCharHexNumber } from "@office-open/core";
import { attr, attrBool, attrNum, escapeXml, findChild, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { sectionPropertiesDesc } from "@parts/document/body/section-properties/descriptor";
import type { DocumentBackgroundOptions } from "@parts/document/document-background";
import { parseDrawingRun } from "@parts/drawing/drawing-parse";
import { FontWrapper } from "@parts/fonts/font-wrapper";
import { HeadingLevel } from "@parts/paragraph/formatting/style";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";
import { parseFormFieldData } from "@parts/paragraph/run/form-field";
import type { FormFieldOptions } from "@parts/paragraph/run/form-field";
import type { RunOptions } from "@parts/paragraph/run/run";
import { parseRun, parseRunProperties, parsedRunToOptions } from "@parts/paragraph/run/run-parse";
import { stringifyTableOfContents } from "@parts/table-of-contents/descriptor";
import type { VmlShapeStyle } from "@parts/textbox/shape/shape";
import { styleToKeyMap } from "@parts/textbox/shape/shape";
import type { SectionChild } from "@shared/section";

import type { DocxReadContext, DocxWriteContext, BodyContext } from "./context";
import { tableDesc, altChunkDesc, subDocDesc, sdtBlockDesc, customXmlBlockDesc } from "./parts";
import { stringifyChildDispatch } from "./parts/inline";
import { parseMathChildren } from "./parts/paragraph/math/stringify";
import type { ParagraphChild } from "./parts/paragraph/paragraph";
import { stringifyParagraphProperties, stringifyRunProperties } from "./parts/paragraph/stringify";

export type { BodyContext } from "./context";

// ── Run ──

/**
 * Mapping from empty child property name to self-closing XML element.
 *
 * These are single-key objects like `{ noBreakHyphen: true }` that map to
 * self-closing XML elements with no attributes.
 *
 * XSD reference: EG_RunInnerContent group in wml.xsd.
 */
const EMPTY_RUN_ELEMENTS: Record<string, string> = {
  noBreakHyphen: "<w:noBreakHyphen/>",
  softHyphen: "<w:softHyphen/>",
  dayShort: "<w:dayShort/>",
  monthShort: "<w:monthShort/>",
  yearShort: "<w:yearShort/>",
  dayLong: "<w:dayLong/>",
  monthLong: "<w:monthLong/>",
  yearLong: "<w:yearLong/>",
  annotationRef: "<w:annotationRef/>",
  footnoteRef: "<w:footnoteRef/>",
  endnoteRef: "<w:endnoteRef/>",
  separator: "<w:separator/>",
  continuationSeparator: "<w:continuationSeparator/>",
  pgNum: "<w:pgNum/>",
  carriageReturn: "<w:cr/>",
  lastRenderedPageBreak: "<w:lastRenderedPageBreak/>",
};

/**
 * Stringify a run (w:r) from pure JSON options.
 *
 * Handles text, children, breaks, and run properties.
 */
export function stringifyRun(opts: RunOptions, ctx: BodyContext): string {
  const parts: string[] = [];

  // Pre-scan children for commentReference — needs CommentReference style in rPr
  let commentRefStyle = false;
  if (opts.children) {
    for (const child of opts.children) {
      if (typeof child === "object" && child !== null && "commentReference" in child) {
        commentRefStyle = true;
        break;
      }
    }
  }

  // Run properties — inject CommentReference style if needed
  const runOpts = commentRefStyle ? { ...opts, style: "CommentReference" as const } : opts;
  const rPr = stringifyRunProperties(runOpts);
  if (rPr) parts.push(rPr);

  // Breaks
  if (opts.break) {
    for (let i = 0; i < opts.break; i++) {
      parts.push("<w:br/>");
    }
  }

  // Children or text
  if (opts.children) {
    for (const child of opts.children) {
      if (typeof child === "string") {
        // Simple text string — direct output
        parts.push(`<w:t xml:space="preserve">${escapeXml(child)}</w:t>`);
      } else if (typeof child === "object" && child !== null) {
        // Simple run-level elements — bare content, no <w:r> wrapper
        if ("tab" in child) {
          parts.push("<w:tab/>");
          continue;
        }
        if ("pageBreak" in child) {
          parts.push('<w:br w:type="page"/>');
          continue;
        }
        if ("columnBreak" in child) {
          parts.push('<w:br w:type="column"/>');
          continue;
        }
        if ("commentReference" in child) {
          parts.push(`<w:commentReference w:id="${Number(child.commentReference)}"/>`);
          continue;
        }

        // Empty run elements — self-closing XML with no attributes
        // { noBreakHyphen: true } → <w:noBreakHyphen/>, etc.
        const emptyXml = EMPTY_RUN_ELEMENTS[Object.keys(child)[0]];
        if (emptyXml) {
          parts.push(emptyXml);
          continue;
        }

        // JSON child dispatch (images, charts, etc.)
        const jsonResult = stringifyChildDispatch(child as ParagraphChild, ctx);
        if (jsonResult !== undefined) {
          if (Array.isArray(jsonResult)) {
            parts.push(...jsonResult);
          } else {
            parts.push(jsonResult);
          }
        } else {
          // Fallback: treat as IXmlableObject-like — should not happen in JSON path
          throw new Error(`Unsupported run child type: ${Object.keys(child).join(", ")}`);
        }
      }
    }
  } else if (opts.text !== undefined) {
    parts.push(`<w:t xml:space="preserve">${escapeXml(String(opts.text))}</w:t>`);
  }

  // rsid attributes on <w:r>
  const rsidAttrs: string[] = [];
  if (opts.rsidRPr) rsidAttrs.push(` w:rsidRPr="${opts.rsidRPr}"`);
  if (opts.rsidDel) rsidAttrs.push(` w:rsidDel="${opts.rsidDel}"`);
  const attr = rsidAttrs.join("");

  const body = parts.join("");
  return body.length === 0 ? (attr ? `<w:r${attr}/>` : "<w:r/>") : `<w:r${attr}>${body}</w:r>`;
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
  const resolved: ParagraphOptions = typeof opts === "string" ? { text: opts } : opts;
  const parts: string[] = [];

  // Build paragraph properties — direct string output, zero IXmlableObject allocation
  const props = stringifyParagraphProperties(resolved);

  // Register numbering references
  if (!(ctx.viewWrapper instanceof FontWrapper)) {
    for (const ref of props.numberingReferences) {
      ctx.file.numbering.createConcreteNumberingInstance(ref.reference, ref.instance);
    }
  }

  // Paragraph properties XML
  if (props.xml) {
    if (sectionPropertiesXml) {
      // Insert sectPr before closing </w:pPr>
      parts.push(props.xml.replace("</w:pPr>", sectionPropertiesXml + "</w:pPr>"));
    } else {
      parts.push(props.xml);
    }
  } else if (sectionPropertiesXml) {
    // No pPr but we need sectPr — wrap in pPr
    parts.push(`<w:pPr>${sectionPropertiesXml}</w:pPr>`);
  }

  // Text shorthand
  if (resolved.text !== undefined) {
    parts.push(stringifyRun({ text: resolved.text }, ctx));
  }

  // Children
  if (resolved.children) {
    for (const child of resolved.children) {
      if (typeof child === "string") {
        parts.push(stringifyRun({ text: child }, ctx));
      } else if (typeof child === "object" && child !== null) {
        // Try JSON child dispatch first (image, chart, pageBreak, etc.)
        const jsonResult = stringifyChildDispatch(child as ParagraphChild, ctx);
        if (jsonResult !== undefined) {
          if (Array.isArray(jsonResult)) {
            parts.push(...jsonResult);
          } else {
            parts.push(jsonResult);
          }
        } else if ("text" in child || "children" in child || "break" in child) {
          // RunOptions-like plain object
          parts.push(stringifyRun(child, ctx));
        } else {
          throw new Error(`Unsupported paragraph child type: ${Object.keys(child).join(", ")}`);
        }
      }
    }
  }

  const body = parts.join("");
  return body ? `<w:p>${body}</w:p>` : "<w:p/>";
}

// ── Body child dispatch ──

/**
 * Stringify a body-level child element.
 *
 * Dispatches to the appropriate stringifier based on the child type.
 * Pure JSON API — no class instance support.
 */
export function stringifyBodyChild(child: SectionChild, ctx: BodyContext): string {
  // Plain object dispatch — all via descriptors
  if ("paragraph" in child) {
    return stringifyParagraph(child.paragraph, ctx);
  }
  if ("table" in child) {
    return tableDesc.stringify(child.table, ctx) ?? "";
  }
  if ("toc" in child) {
    const { alias, ...options } = child.toc;
    return stringifyTableOfContents(alias, options);
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
  if ("rawXml" in child) {
    return child.rawXml;
  }

  throw new Error("Unknown section child type");
}

// ── Document background (pure function, no XmlComponent) ──

/** Re-export styleToKeyMap for use in stringifyTextboxStyle. */
const vmlStyleMap = styleToKeyMap;

function stringifyDocumentBackground(opts: DocumentBackgroundOptions, ctx: BodyContext): string {
  const attrs: string[] = [];
  if (opts.color !== undefined) attrs.push(`w:color="${hexColorValue(opts.color)}"`);
  if (opts.themeColor !== undefined) attrs.push(`w:themeColor="${opts.themeColor}"`);
  if (opts.themeShade !== undefined)
    attrs.push(`w:themeShade="${uCharHexNumber(opts.themeShade)}"`);
  if (opts.themeTint !== undefined) attrs.push(`w:themeTint="${uCharHexNumber(opts.themeTint)}"`);
  const attrStr = attrs.join(" ");

  if (opts.image) {
    const fileName = `${uniqueId()}.${opts.image.type}`;
    const rawData = toUint8Array(opts.image.data) as Uint8Array;

    // Register media
    ctx.file.media.addImage(fileName, {
      type: opts.image.type as "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf",
      data: rawData,
      fileName,
      transformation: { emus: { x: 0, y: 0 }, pixels: { x: 0, y: 0 } },
    });

    const vmlBg = `<v:background id="_x0000_s1025"><v:fill r:id="{${fileName}}" o:title="${fileName}" recolor="t" type="frame"/></v:background>`;
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
  const props = stringifyParagraphProperties(paraOpts);
  const pPrXml = props.xml ?? "";

  // VML shape style string
  const styleStr = style
    ? Object.entries(style)
        .map(([k, v]) => `${vmlStyleMap[k as keyof VmlShapeStyle]}:${v}`)
        .join(";")
    : undefined;

  // Shape attributes
  const shapeAttrs: string[] = [`id="_x0000_s${uniqueId()}"`, `type="#_x0000_t202"`];
  if (styleStr) shapeAttrs.push(`style="${styleStr}"`);

  // Textbox content — serialize children via stringifyBodyChild
  const contentParts: string[] = [];
  if (children) {
    for (const c of children) {
      contentParts.push(stringifyBodyChild(c, ctx));
    }
  }
  const txbxContent = contentParts.join("");

  const vmlTextbox = `<v:textbox style="mso-fit-shape-to-text:t;" insetmode="auto"><w:txbxContent>${txbxContent}</w:txbxContent></v:textbox>`;
  const vshape = `<v:shape ${shapeAttrs.join(" ")}>${vmlTextbox}</v:shape>`;

  return `<w:p>${pPrXml}<w:pict>${vshape}</w:pict></w:p>`;
}

// ── Document body ──

/** Document-level namespace string (cached). */
const DOC_NS =
  'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ' +
  'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" ' +
  'xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" ' +
  'xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" ' +
  'xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" ' +
  'xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" ' +
  'xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" ' +
  'xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" ' +
  'xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" ' +
  'xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" ' +
  'xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" ' +
  'xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" ' +
  'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" ' +
  'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" ' +
  'xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" ' +
  'xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" ' +
  'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"';

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
  parts.push(`<w:document ${DOC_NS} mc:Ignorable="w14 w15 wp14">`);

  // Background (if any)
  if (ctx._options.background) {
    parts.push(stringifyDocumentBackground(ctx._options.background, docCtx));
  }

  // <w:body>
  const bodyParts: string[] = [];

  for (let si = 0; si < sections.length; si++) {
    const section = sections[si];

    // Serialize section children using stringifyBodyChild()
    if (section.children) {
      for (const child of section.children) {
        bodyParts.push(stringifyBodyChild(child, docCtx));
      }
    }

    // Section properties — pure function, no XmlComponent instance
    const sectPrOpts = bodySections[si];
    if (sectPrOpts) {
      const sectPrXml = sectionPropertiesDesc.stringify(sectPrOpts, docCtx) ?? "";
      if (si < sections.length - 1) {
        // Non-last section: embed sectPr in a paragraph's pPr
        bodyParts.push(`<w:p><w:pPr>${sectPrXml}</w:pPr></w:p>`);
      } else {
        // Last section: body-level sectPr
        bodyParts.push(sectPrXml);
      }
    }
  }

  parts.push(`<w:body>${bodyParts.join("")}</w:body>`);
  parts.push("</w:document>");

  return parts.join("");
}

// ────────────────────────────────────────────────────────────────────────────────
// Parse (XML → JSON options)
// ────────────────────────────────────────────────────────────────────────────────

const HEADING_MAP: Record<string, string> = {
  Heading1: HeadingLevel.HEADING_1,
  Heading2: HeadingLevel.HEADING_2,
  Heading3: HeadingLevel.HEADING_3,
  Heading4: HeadingLevel.HEADING_4,
  Heading5: HeadingLevel.HEADING_5,
  Heading6: HeadingLevel.HEADING_6,
  Title: HeadingLevel.TITLE,
};

/**
 * Parse w:pPr element into paragraph properties (without children).
 */
export function parseParagraphProperties(
  el: Element,
  ctx: DocxReadContext,
): Record<string, unknown> {
  const opts: Record<string, unknown> = {};

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
    if (val) opts.alignment = val;
  }

  // Spacing
  const spacing = findChild(el, "w:spacing");
  if (spacing) {
    const sp: Record<string, unknown> = {};
    const before = attrNum(spacing, "w:before");
    if (before !== undefined) sp.before = before;
    const after = attrNum(spacing, "w:after");
    if (after !== undefined) sp.after = after;
    const line = attrNum(spacing, "w:line");
    if (line !== undefined) sp.line = line;
    const lineRule = attr(spacing, "w:lineRule");
    if (lineRule) sp.lineRule = lineRule;
    const beforeAuto =
      attrBool(spacing, "w:beforeAutospacing") ?? attrBool(spacing, "w:beforeLines");
    if (beforeAuto !== undefined) sp.beforeAutoSpacing = beforeAuto;
    const afterAuto = attrBool(spacing, "w:afterAutospacing") ?? attrBool(spacing, "w:afterLines");
    if (afterAuto !== undefined) sp.afterAutoSpacing = afterAuto;
    if (Object.keys(sp).length > 0) opts.spacing = sp;
  }

  // Indent
  const ind = findChild(el, "w:ind");
  if (ind) {
    const indentObj: Record<string, unknown> = {};
    const left = attrNum(ind, "w:left");
    if (left !== undefined) indentObj.left = left;
    const right = attrNum(ind, "w:right");
    if (right !== undefined) indentObj.right = right;
    const start = attrNum(ind, "w:start");
    if (start !== undefined) indentObj.start = start;
    const end = attrNum(ind, "w:end");
    if (end !== undefined) indentObj.end = end;
    const hanging = attrNum(ind, "w:hanging");
    if (hanging !== undefined) indentObj.hanging = hanging;
    const firstLine = attrNum(ind, "w:firstLine");
    if (firstLine !== undefined) indentObj.firstLine = firstLine;
    if (Object.keys(indentObj).length > 0) opts.indent = indentObj;
  }

  // Numbering (w:numPr)
  const numPr = findChild(el, "w:numPr");
  if (numPr) {
    const ilvl = findChild(numPr, "w:ilvl");
    const level = ilvl ? (attrNum(ilvl, "w:val") ?? 0) : 0;
    const numIdEl = findChild(numPr, "w:numId");
    const numId = numIdEl ? attr(numIdEl, "w:val") : undefined;
    if (numId !== undefined && ctx.numberingCache.size > 0) {
      const numEl = ctx.docx.numbering;
      if (numEl) {
        let abstractNumId: string | undefined;
        for (const child of numEl.elements ?? []) {
          if (child.name !== "w:num") continue;
          if (attr(child, "w:numId") === numId) {
            const absRef = findChild(child, "w:abstractNumId");
            abstractNumId = absRef ? attr(absRef, "w:val") : undefined;
            break;
          }
        }
        if (abstractNumId !== undefined) {
          opts.numbering = { reference: `list_${numId}`, level };
        } else {
          opts.bullet = { level };
        }
      } else {
        opts.bullet = { level };
      }
    } else {
      opts.bullet = { level };
    }
  }

  // Tab stops
  const tabs = findChild(el, "w:tabs");
  if (tabs) {
    const tabStops: Record<string, unknown>[] = [];
    for (const tab of tabs.elements ?? []) {
      if (tab.name !== "w:tab") continue;
      const tabObj: Record<string, unknown> = {};
      const pos = attrNum(tab, "w:pos");
      if (pos !== undefined) tabObj.position = pos;
      const val = attr(tab, "w:val");
      if (val) tabObj.type = val;
      const leader = attr(tab, "w:leader");
      if (leader) tabObj.leader = leader;
      tabStops.push(tabObj);
    }
    if (tabStops.length > 0) opts.tabStops = tabStops;
  }

  // On/off properties
  for (const [name, optKey] of [
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
    ["w:overflowPunct", "overflowPunctuation"],
    ["w:suppressOverlap", "suppressOverlap"],
  ] as const) {
    const child = findChild(el, name);
    if (child) opts[optKey] = attrBool(child, "w:val") ?? true;
  }

  // Thematic break
  const pBdr = findChild(el, "w:pBdr");
  if (pBdr) {
    const border: Record<string, unknown> = {};
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const sideEl = findChild(pBdr, `w:${side}`);
      if (sideEl) {
        const sideOpts: Record<string, unknown> = {};
        const style = attr(sideEl, "w:val");
        if (style) sideOpts.style = style;
        const color = attr(sideEl, "w:color");
        if (color) sideOpts.color = color;
        const size = attrNum(sideEl, "w:sz");
        if (size !== undefined) sideOpts.size = size;
        const space = attrNum(sideEl, "w:space");
        if (space !== undefined) sideOpts.space = space;
        border[side] = sideOpts;
      }
    }
    if (Object.keys(border).length > 0) opts.border = border;
  }

  // Shading
  const shd = findChild(el, "w:shd");
  if (shd) {
    const shdObj: Record<string, unknown> = {};
    const fill = attr(shd, "w:fill");
    if (fill) shdObj.fill = fill;
    const color = attr(shd, "w:color");
    if (color) shdObj.color = color;
    const val = attr(shd, "w:val");
    if (val) shdObj.type = val;
    if (Object.keys(shdObj).length > 0) opts.shading = shdObj;
  }

  // Text alignment
  const textAlignment = findChild(el, "w:textAlignment");
  if (textAlignment) {
    const val = attr(textAlignment, "w:val");
    if (val) opts.textAlignment = val;
  }

  // Outline level
  const outlineLvl = findChild(el, "w:outlineLvl");
  if (outlineLvl) {
    const val = attrNum(outlineLvl, "w:val");
    if (val !== undefined) opts.outlineLevel = val;
  }

  // Run properties (paragraph-level defaults)
  const rPr = findChild(el, "w:rPr");
  if (rPr) {
    opts.run = parseRunProperties(rPr);
  }

  // Frame properties
  const framePr = findChild(el, "w:framePr");
  if (framePr) {
    const frame: Record<string, unknown> = {};
    for (const [attrName, optName] of [
      ["w:dropCap", "dropCap"],
      ["w:lines", "lines"],
      ["w:wrap", "wrap"],
      ["w:vAnchor", "vAnchor"],
      ["w:hAnchor", "hAnchor"],
      ["w:x", "x"],
      ["w:y", "y"],
      ["w:hRule", "hRule"],
      ["w:hSpace", "hSpace"],
      ["w:vSpace", "vSpace"],
    ] as const) {
      const val = attr(framePr, attrName);
      if (val !== undefined) frame[optName] = val;
    }
    const w = attrNum(framePr, "w:w");
    if (w !== undefined) frame.width = w;
    const h = attrNum(framePr, "w:h");
    if (h !== undefined) frame.height = h;
    if (Object.keys(frame).length > 0) opts.frame = frame;
  }

  return opts;
}

/**
 * Concatenate `<w:t>` text in a run element.
 *
 * Used to capture a textInput form field's current value from the result run
 * (the runs between the `separate` and `end` fldChars). Only `<w:t>` is read —
 * `<w:tab>`/`<w:br>` in results are ignored as rare for user-entered text.
 */
function collectRunText(el: Element): string {
  let text = "";
  for (const c of el.elements ?? []) {
    if (c.name === "w:t") text += textOf(c);
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
function parseContainerChildren(el: Element, ctx: DocxReadContext): unknown[] {
  const children: unknown[] = [];
  for (const sub of el.elements ?? []) {
    switch (sub.name) {
      case "w:r": {
        const parsed = parseRun(sub, ctx);
        const runOpts = parsedRunToOptions(parsed);
        if (runOpts !== null) children.push(runOpts);
        break;
      }
      case "w:smartTag":
        children.push({ smartTag: parseSmartTagInline(sub, ctx) });
        break;
      case "w:customXml":
        children.push({ customXml: parseCustomXmlInline(sub, ctx) });
        break;
      default:
        break;
    }
  }
  return children;
}

/** Parse a w:smartTag element into its ParagraphChild form. */
function parseSmartTagInline(el: Element, ctx: DocxReadContext): Record<string, unknown> {
  const st: Record<string, unknown> = {};
  const element = attr(el, "w:element");
  if (element) st.element = element;
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
function parseCustomXmlInline(el: Element, ctx: DocxReadContext): Record<string, unknown> {
  const cx: Record<string, unknown> = {};
  const element = attr(el, "w:element");
  if (element) cx.element = element;
  const uri = attr(el, "w:uri");
  if (uri) cx.uri = uri;
  const pr = findChild(el, "w:customXmlPr");
  if (pr) {
    const cxPr: {
      placeholder?: string;
      attrs?: Array<{ name: string; val: string; uri?: string }>;
    } = {};
    const placeholder = findChild(pr, "w:placeholder");
    if (placeholder) {
      const pval = attr(placeholder, "w:val");
      if (pval) cxPr.placeholder = pval;
    }
    const attrsList: Array<{ name: string; val: string; uri?: string }> = [];
    for (const a of pr.elements ?? []) {
      if (a.name !== "w:attr") continue;
      const name = attr(a, "w:name");
      const val = attr(a, "w:val");
      if (name && val) {
        const ao: { name: string; val: string; uri?: string } = { name, val };
        const auri = attr(a, "w:uri");
        if (auri) ao.uri = auri;
        attrsList.push(ao);
      }
    }
    if (attrsList.length > 0) cxPr.attrs = attrsList;
    if (cxPr.placeholder !== undefined || cxPr.attrs !== undefined) cx.customXmlPr = cxPr;
  }
  const content = parseContainerChildren(el, ctx);
  if (content.length > 0) cx.children = content;
  return cx;
}

/**
 * Parse a w:p element into ParagraphOptions.
 */
export function parseParagraph(el: Element, ctx: DocxReadContext): ParagraphOptions {
  const opts: Record<string, unknown> = {};

  const pPr = findChild(el, "w:pPr");
  if (pPr) {
    Object.assign(opts, parseParagraphProperties(pPr, ctx));
  }

  // Parse children
  const childList: unknown[] = [];

  // Form field accumulator: a w:checkBox/ddList/textInput field spans several
  // w:r elements (begin fldChar → instrText → separate → result → end). The
  // definition lives in w:ffData inside the begin fldChar; the result runs
  // (separate→end) are captured for textInput only — checkbox/dropdown state
  // already lives in w:ffData — and the whole field collapses to one child.
  let pendingFormField: FormFieldOptions | null = null;
  // True once the `separate` fldChar is seen: subsequent runs (up to `end`)
  // are the field's result.
  let collectingResult = false;

  for (const child of el.elements ?? []) {
    switch (child.name) {
      case "w:pPr":
        break;
      case "w:r": {
        // Form field: fldChar markers + the instrText/result runs between them.
        const fldCharEl = findChild(child, "w:fldChar");
        if (fldCharEl) {
          const fctype = attr(fldCharEl, "w:fldCharType");
          if (fctype === "begin") {
            // Only form fields carry w:ffData on the begin fldChar. A plain
            // complex field (PAGE/DATE/TOC) has none and must not be collected
            // as a form field — otherwise it collapses to an empty {formField:{}}.
            const ffDataEl = findChild(fldCharEl, "w:ffData");
            if (ffDataEl) pendingFormField = parseFormFieldData(ffDataEl);
          } else if (fctype === "separate") {
            collectingResult = true;
          } else if (fctype === "end" && pendingFormField) {
            childList.push({ formField: pendingFormField });
            pendingFormField = null;
            collectingResult = false;
          }
          break;
        }
        if (pendingFormField) {
          // separate→end runs are the field result. Capture a textInput's
          // current (user-entered) value for round-trip fidelity; checkbox
          // and dropdown results are discarded (their state is in w:ffData).
          if (collectingResult && pendingFormField.textInput) {
            const text = collectRunText(child);
            if (text) {
              const ti = pendingFormField.textInput;
              ti.value = (ti.value ?? "") + text;
            }
          }
          break; // instrText / result — state lives in ffData
        }

        const drawingEl = findChild(child, "w:drawing");
        if (drawingEl) {
          const drawingChild = parseDrawingRun(drawingEl, ctx);
          if (drawingChild) {
            childList.push(drawingChild);
            break;
          }
        }
        const parsed = parseRun(child, ctx);
        const runOpts = parsedRunToOptions(parsed);
        if (runOpts !== null) childList.push(runOpts);
        break;
      }
      case "w:hyperlink": {
        const hl: Record<string, unknown> = {};
        const rId = attr(child, "r:id");
        if (rId) {
          const target = ctx.docx.partRefs.hyperlinks.get(rId);
          if (target) hl.link = target;
        }
        const anchor = attr(child, "w:anchor");
        if (anchor) hl.anchor = anchor;
        const tooltip = attr(child, "w:tooltip");
        if (tooltip) hl.tooltip = tooltip;

        const linkRuns: unknown[] = [];
        for (const sub of child.elements ?? []) {
          if (sub.name === "w:r") {
            const parsed = parseRun(sub, ctx);
            const runOpts = parsedRunToOptions(parsed);
            linkRuns.push(runOpts);
          }
        }
        if (linkRuns.length > 0) {
          hl.children = linkRuns;
          childList.push({ hyperlink: hl });
        }
        break;
      }
      case "w:bookmarkStart": {
        const id = attrNum(child, "w:id");
        const name = attr(child, "w:name");
        if (id !== undefined && name) {
          childList.push({ bookmarkStart: { id, name } });
        }
        break;
      }
      case "w:bookmarkEnd": {
        const id = attrNum(child, "w:id");
        if (id !== undefined) childList.push({ bookmarkEnd: id });
        break;
      }
      case "w:commentRangeStart": {
        const id = attrNum(child, "w:id");
        if (id !== undefined) childList.push({ commentRangeStart: id });
        break;
      }
      case "w:commentRangeEnd": {
        const id = attrNum(child, "w:id");
        if (id !== undefined) childList.push({ commentRangeEnd: id });
        break;
      }
      case "w:commentReference": {
        const id = attrNum(child, "w:id");
        if (id !== undefined) childList.push({ commentReference: id });
        break;
      }
      case "m:oMath": {
        const mathChildren = parseMathChildren(child);
        childList.push({ math: { children: mathChildren } });
        break;
      }
      case "w:ins": {
        const insRun = findChild(child, "w:r");
        if (insRun) {
          const parsed = parseRun(insRun, ctx);
          const runOpts = parsedRunToOptions(parsed);
          if (runOpts !== null && typeof runOpts === "object" && !("commentReference" in runOpts)) {
            childList.push({
              insertion: {
                id: attrNum(child, "w:id") ?? 0,
                author: attr(child, "w:author") ?? "",
                date: attr(child, "w:date") ?? "",
                ...(runOpts as Record<string, unknown>),
              },
            });
          }
        }
        break;
      }
      case "w:del": {
        const delRun = findChild(child, "w:r");
        if (delRun) {
          const parsed = parseRun(delRun, ctx);
          const runOpts = parsedRunToOptions(parsed);
          if (runOpts !== null && typeof runOpts === "object" && !("commentReference" in runOpts)) {
            childList.push({
              deletion: {
                id: attrNum(child, "w:id") ?? 0,
                author: attr(child, "w:author") ?? "",
                date: attr(child, "w:date") ?? "",
                ...(runOpts as Record<string, unknown>),
              },
            });
          }
        }
        break;
      }
      case "w:fldSimple": {
        const instruction = attr(child, "w:instr");
        if (instruction) {
          const sf: { instruction: string; cachedValue?: string } = { instruction };
          // cachedValue: concatenate the result-run <w:t> text (one or more
          // <w:r> children between the fldSimple tags).
          let cachedValue = "";
          for (const sub of child.elements ?? []) {
            if (sub.name === "w:r") cachedValue += collectRunText(sub);
          }
          if (cachedValue) sf.cachedValue = cachedValue;
          childList.push({ simpleField: sf });
        }
        break;
      }
      case "w:smartTag": {
        const st = parseSmartTagInline(child, ctx);
        if (st.element) childList.push({ smartTag: st });
        break;
      }
      case "w:customXml": {
        const cx = parseCustomXmlInline(child, ctx);
        if (cx.element) childList.push({ customXml: cx });
        break;
      }
      default:
        break;
    }
  }

  // Simple text optimization
  if (childList.length > 0) {
    const allStrings = childList.every(
      (c) =>
        typeof c === "object" &&
        c !== null &&
        "text" in (c as Record<string, unknown>) &&
        Object.keys(c as Record<string, unknown>).length === 1,
    );
    if (allStrings) {
      const combined = childList.map((c) => (c as Record<string, unknown>).text as string).join("");
      if (combined && Object.keys(opts).length === 0) {
        return combined as unknown as ParagraphOptions;
      }
      if (combined) {
        opts.text = combined;
        return opts as ParagraphOptions;
      }
    }

    opts.children = childList as (RunOptions | string)[];
  }

  return opts as ParagraphOptions;
}
