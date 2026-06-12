/**
 * Shared inline run/paragraph stringification for DOCX descriptors.
 *
 * Used by table.ts, comments.ts, body.ts, and other descriptors that need to
 * serialize paragraph/run content. Includes JSON child dispatch for all
 * IParagraphJsonChild variants (image, chart, hyperlink, etc.).
 *
 * Pure string concatenation — zero IXmlableObject, zero BaseXmlComponent.
 *
 * @module
 */

import { toUint8Array } from "@office-open/core";
import { TargetModeType } from "@office-open/core";
import { uniqueId } from "@office-open/core";
import { chartSpaceDesc } from "@office-open/core/chart";
import type { SourceRectangleOptions } from "@office-open/core/drawingml";
import { createDataModel } from "@office-open/core/smartart";
import { escapeXml } from "@office-open/xml";
import type { IParagraphJsonChild, ParagraphOptions } from "@parts/paragraph/paragraph";
import type { IImageOptions } from "@parts/paragraph/run/image-run";
import type { RunOptions } from "@parts/paragraph/run/run";
import type { SmartArtOptions } from "@parts/paragraph/run/smartart-run";
import type { ChartMediaData, IMediaData, SmartArtMediaData, WpsMediaData } from "@shared/media";
import { createTransformation } from "@shared/media";

import type { BodyContext } from "../context";
import { drawingDesc } from "./drawing";
import { stringifyMath } from "./paragraph/math/stringify";
import { stringifyParagraphProperties, stringifyRunProperties } from "./paragraph/stringify";

// ── Run ──

export function stringifyRunInline(opts: RunOptions, ctx: BodyContext): string {
  const parts: string[] = [];

  const rPr = stringifyRunProperties(opts);
  if (rPr) parts.push(rPr);

  if (opts.break) {
    for (let i = 0; i < opts.break; i++) parts.push("<w:br/>");
  }

  if (opts.children) {
    for (const child of opts.children) {
      if (typeof child === "string") {
        parts.push(`<w:t xml:space="preserve">${escapeXml(child)}</w:t>`);
      } else {
        // JSON child dispatch
        const jsonResult = stringifyJsonChild(child as ParagraphChild, ctx);
        if (jsonResult !== undefined) {
          if (Array.isArray(jsonResult)) {
            parts.push(...jsonResult);
          } else {
            parts.push(jsonResult);
          }
        } else if ("text" in child || "children" in child || "break" in child) {
          parts.push(stringifyRunInline(child as RunOptions, ctx));
        }
      }
    }
  } else if (opts.text !== undefined) {
    parts.push(`<w:t xml:space="preserve">${escapeXml(String(opts.text))}</w:t>`);
  }

  const rsidAttrs: string[] = [];
  if (opts.rsidRPr) rsidAttrs.push(` w:rsidRPr="${opts.rsidRPr}"`);
  if (opts.rsidDel) rsidAttrs.push(` w:rsidDel="${opts.rsidDel}"`);
  const attr = rsidAttrs.join("");

  const body = parts.join("");
  return body.length === 0 ? (attr ? `<w:r${attr}/>` : "<w:r/>") : `<w:r${attr}>${body}</w:r>`;
}

// ── Image helpers ──

function createImageData(
  data: Uint8Array,
  transformation: IImageOptions["transformation"],
  key: string,
  srcRect?: SourceRectangleOptions,
): Pick<IMediaData, "data" | "fileName" | "transformation" | "srcRect"> {
  return {
    data,
    fileName: key,
    srcRect,
    transformation: createTransformation(transformation),
  };
}

let nextChartId = 1;

// ── JSON child dispatch ──

/**
 * Stringify an IParagraphJsonChild into one or more XML strings.
 *
 * Handles side effects (media, chart, smartArt, relationship registration)
 * directly without creating temporary class instances.
 *
 * Returns `undefined` if the child is not a recognized JSON wrapper.
 */
/**
 * Type for children that may be a JSON child or a RunOptions.
 * Used by stringifyParagraphInline to attempt JSON dispatch before RunOptions.
 */
export type ParagraphChild = IParagraphJsonChild | RunOptions;

export function stringifyJsonChild(
  child: ParagraphChild,
  ctx: BodyContext,
): string | string[] | undefined {
  // Simple break types — pure XML, no side effects
  if ("pageBreak" in child) return '<w:r><w:br w:type="page"/></w:r>';
  if ("columnBreak" in child) return '<w:r><w:br w:type="column"/></w:r>';
  if ("tab" in child) return "<w:r><w:tab/></w:r>";

  // Reference types — pure XML, no side effects
  if ("footnoteReference" in child)
    return `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${child.footnoteReference}"/></w:r>`;
  if ("endnoteReference" in child)
    return `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${child.endnoteReference}"/></w:r>`;

  // Comment markers — pure XML
  if ("commentRangeStart" in child)
    return `<w:commentRangeStart w:id="${child.commentRangeStart}"/>`;
  if ("commentRangeEnd" in child) return `<w:commentRangeEnd w:id="${child.commentRangeEnd}"/>`;
  if ("commentReference" in child)
    return `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${child.commentReference}"/></w:r>`;

  // Bookmark markers — pure XML
  if ("bookmarkStart" in child)
    return `<w:bookmarkStart w:id="${child.bookmarkStart.id}" w:name="${child.bookmarkStart.name}"/>`;
  if ("bookmarkEnd" in child) return `<w:bookmarkEnd w:id="${child.bookmarkEnd}"/>`;

  // Symbol run — direct XML output
  if ("symbolRun" in child) {
    const opts = child.symbolRun;
    return stringifyRunInline(
      {
        ...opts,
        children: [`<w:sym w:char="${opts.char}" w:font="${opts.symbolfont ?? "Wingdings"}"/>`],
      },
      ctx,
    );
  }

  // Image — side effect: media registration
  if ("image" in child) {
    const opts = child.image;
    const key = `${uniqueId()}.${opts.type}`;
    const rawData = toUint8Array(opts.data) as Uint8Array;

    let mediaData: IMediaData;
    if (opts.type === "svg") {
      const fallbackData = toUint8Array(opts.fallback.data) as Uint8Array;
      mediaData = {
        type: "svg",
        ...createImageData(rawData, opts.transformation, key, opts.srcRect),
        fallback: {
          type: opts.fallback.type,
          ...createImageData(
            fallbackData,
            opts.transformation,
            `${uniqueId()}.${opts.fallback.type}`,
          ),
        },
      } as IMediaData;
    } else {
      mediaData = {
        type: opts.type,
        ...createImageData(rawData, opts.transformation, key, opts.srcRect),
      } as IMediaData;
    }

    // Register media
    ctx.file.media.addImage(mediaData.fileName, mediaData);
    if (mediaData.type === "svg") {
      ctx.file.media.addImage(mediaData.fallback.fileName, mediaData.fallback);
    }

    // Build drawing XML via descriptor (zero XmlComponent instances)
    const drawingXml = drawingDesc.stringify(
      {
        mediaData,
        docProperties: opts.altText,
        floating: opts.floating,
      },
      ctx,
    );
    return `<w:r>${drawingXml}</w:r>`;
  }

  // Chart — side effect: chart registration
  if ("chart" in child) {
    const opts = child.chart;
    const chartKey = `chart_${nextChartId++}`;
    const mediaData: ChartMediaData = {
      chartKey,
      transformation: createTransformation(opts.transformation),
      type: "chart",
    };

    // Register chart
    const chartXml = chartSpaceDesc.stringify(
      {
        categories: opts.categories,
        series: opts.series,
        showLegend: opts.showLegend,
        style: opts.style,
        title: opts.title,
        type: opts.type,
        threeD: opts.threeD,
      },
      ctx.file,
    );
    ctx.file.charts.addChart(chartKey, {
      key: chartKey,
      chartSpaceXml: chartXml ?? "",
    });

    const drawingXml = drawingDesc.stringify(
      {
        mediaData,
        docProperties: opts.altText,
        floating: opts.floating,
      },
      ctx,
    );
    return `<w:r>${drawingXml}</w:r>`;
  }

  // SmartArt — side effect: smartArt registration
  if ("smartArt" in child) {
    const opts = child.smartArt;
    const hash = hashSmartArtData(opts);
    const smartArtKey = `smartart_${hash}`;
    const mediaData: SmartArtMediaData = {
      smartArtKey,
      transformation: createTransformation(opts.transformation),
      type: "smartart",
    };

    // Register SmartArt
    const layoutId = opts.layout ?? "default";
    const styleId = opts.style ?? "simple1";
    const colorId = opts.color ?? "accent1_2";
    const dataModelXml = createDataModel(opts.data.nodes, layoutId, styleId, colorId);

    ctx.file.smartArts.addSmartArt(smartArtKey, {
      dataModelXml,
      key: smartArtKey,
      layout: layoutId,
      style: styleId,
      color: colorId,
    });

    const drawingXml = drawingDesc.stringify(
      {
        mediaData,
        docProperties: opts.altText,
        floating: opts.floating,
      },
      ctx,
    );
    return `<w:r>${drawingXml}</w:r>`;
  }

  // WPS Shape (WordProcessing Shape) — side effect: blip fill media registration
  if ("wpsShape" in child) {
    const opts = child.wpsShape;
    const mediaData: WpsMediaData = {
      data: opts,
      transformation: createTransformation(opts.transformation),
      type: "wps",
    };

    const drawingXml = drawingDesc.stringify(
      {
        mediaData,
        docProperties: opts.altText,
        floating: opts.floating,
        outline: opts.outline,
        fill: opts.fill,
      },
      ctx,
    );
    return `<w:r>${drawingXml}</w:r>`;
  }

  // Ruby annotation — pure string concatenation
  if ("ruby" in child && typeof child.ruby === "object" && child.ruby !== null) {
    const r = child.ruby as import("@parts/paragraph/run/ruby").RubyOptions;
    const align = r.alignment ?? "center";
    const hps = (r.fontSize ?? 10) * 2;
    const hpsRaise = (r.raise ?? 10) * 2;
    const hpsBaseText = (r.baseFontSize ?? 20) * 2;
    const lid = r.languageId ?? "ja-JP";

    const prParts = [
      `<w:rubyAlign w:val="${align}"/>`,
      `<w:hps w:val="${hps}"/>`,
      `<w:hpsRaise w:val="${hpsRaise}"/>`,
      `<w:hpsBaseText w:val="${hpsBaseText}"/>`,
      `<w:lid w:val="${lid}"/>`,
    ];
    if (r.dirty) prParts.push("<w:dirty/>");

    const rt = `<w:rt><w:r><w:t xml:space="preserve">${escapeXml(r.text)}</w:t></w:r></w:rt>`;
    const rubyBase = `<w:rubyBase><w:r><w:t xml:space="preserve">${escapeXml(r.base)}</w:t></w:r></w:rubyBase>`;

    return `<w:ruby><w:rubyPr>${prParts.join("")}</w:rubyPr>${rt}${rubyBase}</w:ruby>`;
  }

  // Math — pure string concatenation
  if ("math" in child && typeof child.math === "object" && child.math !== null) {
    const mathOpts = child.math;
    const children = mathOpts.children ?? [];
    return stringifyMath(children);
  }

  // Inserted text run — pure function
  if ("insertion" in child) {
    const opts = child.insertion;
    const { id, author, date, ...runOpts } = opts;
    const runXml = stringifyRunInline(runOpts, ctx);
    return `<w:ins w:id="${id}" w:author="${escapeXml(String(author))}" w:date="${date}">${runXml}</w:ins>`;
  }

  // Deleted text run — pure function
  if ("deletion" in child) {
    const opts = child.deletion;
    const { id, author, date, ...runOpts } = opts;

    const parts: string[] = [];
    const rPr = stringifyRunProperties(runOpts);
    if (rPr) parts.push(rPr);

    if (runOpts.break) {
      for (let i = 0; i < runOpts.break; i++) parts.push("<w:br/>");
    }

    if (runOpts.children) {
      for (const c of runOpts.children) {
        if (typeof c === "string") {
          // Page number fields use delInstrText instead of instrText
          const fieldMap: Record<string, string> = {
            CURRENT: "PAGE",
            TOTAL_PAGES: "NUMPAGES",
            TOTAL_PAGES_IN_SECTION: "SECTIONPAGES",
          };
          const instrText = fieldMap[c];
          if (instrText) {
            parts.push(
              '<w:fldChar w:fldCharType="begin"/>' +
                `<w:delInstrText xml:space="preserve">${instrText}</w:delInstrText>` +
                '<w:fldChar w:fldCharType="separate"/>' +
                '<w:fldChar w:fldCharType="end"/>',
            );
          } else {
            parts.push(`<w:delText xml:space="preserve">${escapeXml(c)}</w:delText>`);
          }
        }
      }
    } else if (runOpts.text) {
      parts.push(`<w:delText xml:space="preserve">${escapeXml(String(runOpts.text))}</w:delText>`);
    }

    const runBody = parts.join("");
    return `<w:del w:id="${id}" w:author="${escapeXml(String(author))}" w:date="${date}"><w:r>${runBody}</w:r></w:del>`;
  }

  // Hyperlink — side effect: relationship registration
  if ("hyperlink" in child) {
    const hl = child.hyperlink;

    // Serialize children using stringifyRunInline
    const childParts: string[] = [];
    if (hl.children) {
      for (const rc of hl.children) {
        if (typeof rc === "string") {
          childParts.push(stringifyRunInline({ text: rc }, ctx));
        } else {
          childParts.push(stringifyRunInline(rc, ctx));
        }
      }
    }
    const body = childParts.join("");

    if (hl.link) {
      const linkId = uniqueId();
      ctx.viewWrapper.relationships.addRelationship(
        linkId,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        hl.link,
        TargetModeType.EXTERNAL,
      );
      const attrs = [`r:id="rId${linkId}"`, 'w:history="1"'];
      if (hl.tooltip) attrs.push(`w:tooltip="${escapeXml(hl.tooltip)}"`);
      return `<w:hyperlink ${attrs.join(" ")}>${body}</w:hyperlink>`;
    }
    if (hl.anchor) {
      const attrs = [`w:anchor="${escapeXml(hl.anchor)}"`, 'w:history="1"'];
      if (hl.tooltip) attrs.push(`w:tooltip="${escapeXml(hl.tooltip)}"`);
      return `<w:hyperlink ${attrs.join(" ")}>${body}</w:hyperlink>`;
    }
    return "";
  }

  // ── Proof error markers ──
  if ("proofErr" in child) return `<w:proofErr w:type="${child.proofErr}"/>`;

  // ── Positional tab ──
  if ("positionalTab" in child) {
    const pt = child.positionalTab;
    return `<w:ptab w:alignment="${pt.alignment}" w:leader="${pt.leader}" w:relativeTo="${pt.relativeTo}"/>`;
  }

  // ── Permission range markers ──
  if ("permStart" in child) {
    const ps = child.permStart;
    const a: string[] = [`w:id="${ps.id}"`];
    if (ps.ed !== undefined) a.push(`w:ed="${escapeXml(String(ps.ed))}"`);
    if (ps.editGroup !== undefined) a.push(`w:edGrp="${ps.editGroup}"`);
    if (ps.colFirst !== undefined) a.push(`w:colFirst="${ps.colFirst}"`);
    if (ps.colLast !== undefined) a.push(`w:colLast="${ps.colLast}"`);
    return `<w:permStart ${a.join(" ")}/>`;
  }
  if ("permEnd" in child) return `<w:permEnd w:id="${child.permEnd}"/>`;

  // ── Move revision range markers ──
  if ("moveFromRangeStart" in child) {
    const m = child.moveFromRangeStart;
    const a: string[] = [`w:id="${m.id}"`];
    if (m.name) a.push(`w:name="${escapeXml(m.name)}"`);
    if (m.author) a.push(`w:author="${escapeXml(m.author)}"`);
    if (m.date) a.push(`w:date="${m.date}"`);
    return `<w:moveFromRangeStart ${a.join(" ")}/>`;
  }
  if ("moveFromRangeEnd" in child) return `<w:moveFromRangeEnd w:id="${child.moveFromRangeEnd}"/>`;
  if ("moveToRangeStart" in child) {
    const m = child.moveToRangeStart;
    const a: string[] = [`w:id="${m.id}"`];
    if (m.name) a.push(`w:name="${escapeXml(m.name)}"`);
    if (m.author) a.push(`w:author="${escapeXml(m.author)}"`);
    if (m.date) a.push(`w:date="${m.date}"`);
    return `<w:moveToRangeStart ${a.join(" ")}/>`;
  }
  if ("moveToRangeEnd" in child) return `<w:moveToRangeEnd w:id="${child.moveToRangeEnd}"/>`;

  // ── Move revision text runs ──
  if ("movedFrom" in child) {
    const opts = child.movedFrom;
    const { id, author, date, ...runOpts } = opts;
    const runXml = stringifyRunInline(runOpts, ctx);
    return `<w:moveFrom w:id="${id}" w:author="${escapeXml(String(author))}" w:date="${date}">${runXml}</w:moveFrom>`;
  }
  if ("movedTo" in child) {
    const opts = child.movedTo;
    const { id, author, date, ...runOpts } = opts;
    const runXml = stringifyRunInline(runOpts, ctx);
    return `<w:moveTo w:id="${id}" w:author="${escapeXml(String(author))}" w:date="${date}">${runXml}</w:moveTo>`;
  }

  // ── Custom XML range markers (track changes) ──
  if ("customXmlInsRangeStart" in child) {
    const o = child.customXmlInsRangeStart;
    return `<w:customXmlInsRangeStart w:id="${o.id}" w:author="${escapeXml(o.author)}"${o.date ? ` w:date="${o.date}"` : ""}/>`;
  }
  if ("customXmlInsRangeEnd" in child)
    return `<w:customXmlInsRangeEnd w:id="${child.customXmlInsRangeEnd}"/>`;
  if ("customXmlDelRangeStart" in child) {
    const o = child.customXmlDelRangeStart;
    return `<w:customXmlDelRangeStart w:id="${o.id}" w:author="${escapeXml(o.author)}"${o.date ? ` w:date="${o.date}"` : ""}/>`;
  }
  if ("customXmlDelRangeEnd" in child)
    return `<w:customXmlDelRangeEnd w:id="${child.customXmlDelRangeEnd}"/>`;
  if ("customXmlMoveFromRangeStart" in child) {
    const o = child.customXmlMoveFromRangeStart;
    return `<w:customXmlMoveFromRangeStart w:id="${o.id}" w:author="${escapeXml(o.author)}"${o.date ? ` w:date="${o.date}"` : ""}/>`;
  }
  if ("customXmlMoveFromRangeEnd" in child)
    return `<w:customXmlMoveFromRangeEnd w:id="${child.customXmlMoveFromRangeEnd}"/>`;
  if ("customXmlMoveToRangeStart" in child) {
    const o = child.customXmlMoveToRangeStart;
    return `<w:customXmlMoveToRangeStart w:id="${o.id}" w:author="${escapeXml(o.author)}"${o.date ? ` w:date="${o.date}"` : ""}/>`;
  }
  if ("customXmlMoveToRangeEnd" in child)
    return `<w:customXmlMoveToRangeEnd w:id="${child.customXmlMoveToRangeEnd}"/>`;

  // ── Simple field ──
  if ("simpleField" in child) {
    const sf = child.simpleField;
    if (sf.cachedValue !== undefined) {
      return `<w:fldSimple w:instr="${escapeXml(sf.instruction)}"><w:r><w:t>${escapeXml(sf.cachedValue)}</w:t></w:r></w:fldSimple>`;
    }
    return `<w:fldSimple w:instr="${escapeXml(sf.instruction)}"/>`;
  }

  // ── Sequential identifier (SEQ field) ──
  if ("seqIdentifier" in child) {
    const id = child.seqIdentifier;
    return (
      "<w:r>" +
      '<w:fldChar w:fldCharType="begin"/>' +
      `<w:instrText xml:space="preserve"> SEQ ${escapeXml(id)} </w:instrText>` +
      '<w:fldChar w:fldCharType="separate"/>' +
      '<w:fldChar w:fldCharType="end"/>' +
      "</w:r>"
    );
  }

  // ── Page reference (PAGEREF field) ──
  if ("pageReference" in child) {
    const pr = child.pageReference;
    let instr = ` PAGEREF ${escapeXml(pr.bookmarkId)} `;
    if (pr.hyperlink) instr += "\\h ";
    if (pr.useRelativePosition) instr += "\\p ";
    return (
      "<w:r>" +
      '<w:fldChar w:fldCharType="begin"/>' +
      `<w:instrText xml:space="preserve">${instr}</w:instrText>` +
      '<w:fldChar w:fldCharType="end"/>' +
      "</w:r>"
    );
  }

  // ── Bidirectional text containers ──
  if ("dir" in child) {
    const d = child.dir;
    const childXml = serializeDirChildren(d.children, ctx);
    return `<w:dir w:val="${d.val}">${childXml}</w:dir>`;
  }
  if ("bdo" in child) {
    const b = child.bdo;
    const childXml = serializeDirChildren(b.children, ctx);
    return `<w:bdo w:val="${b.val}">${childXml}</w:bdo>`;
  }

  // ── Smart tag ──
  if ("smartTag" in child) {
    const st = child.smartTag;
    const attrs: string[] = [];
    if (st.uri) attrs.push(`w:uri="${escapeXml(st.uri)}"`);
    attrs.push(`w:element="${escapeXml(st.element)}"`);

    const parts: string[] = [];
    if (st.properties?.length) {
      const propParts: string[] = [];
      for (const p of st.properties) {
        const pa: string[] = [];
        if (p.uri) pa.push(`w:uri="${escapeXml(p.uri)}"`);
        pa.push(`w:name="${escapeXml(p.name)}"`, `w:val="${escapeXml(p.val)}"`);
        propParts.push(`<w:attr ${pa.join(" ")}/>`);
      }
      parts.push(`<w:smartTagPr>${propParts.join("")}</w:smartTagPr>`);
    }
    if (st.children) {
      for (const c of st.children) {
        parts.push(
          typeof c === "string" ? stringifyRunInline({ text: c }, ctx) : stringifyRunInline(c, ctx),
        );
      }
    }
    return `<w:smartTag ${attrs.join(" ")}>${parts.join("")}</w:smartTag>`;
  }

  // ── Custom XML run ──
  if ("customXml" in child) {
    const cx = child.customXml;
    const attrs: string[] = [`w:element="${escapeXml(cx.element)}"`];
    if (cx.uri) attrs.push(`w:uri="${escapeXml(cx.uri)}"`);

    const parts: string[] = [];
    if (cx.customXmlPr) {
      const prParts: string[] = [];
      if (cx.customXmlPr.placeholder)
        prParts.push(`<w:placeholder w:val="${escapeXml(cx.customXmlPr.placeholder)}"/>`);
      if (cx.customXmlPr.attrs?.length) {
        for (const a of cx.customXmlPr.attrs) {
          const aa: string[] = [`w:name="${escapeXml(a.name)}"`, `w:val="${escapeXml(a.val)}"`];
          if (a.uri) aa.push(`w:uri="${escapeXml(a.uri)}"`);
          prParts.push(`<w:attr ${aa.join(" ")}/>`);
        }
      }
      if (prParts.length) parts.push(`<w:customXmlPr>${prParts.join("")}</w:customXmlPr>`);
    }
    if (cx.children) {
      for (const c of cx.children) {
        if (typeof c === "string") {
          parts.push(stringifyRunInline({ text: c }, ctx));
        } else {
          const jr = stringifyJsonChild(c as ParagraphChild, ctx);
          if (jr !== undefined) {
            parts.push(Array.isArray(jr) ? jr.join("") : jr);
          } else {
            parts.push(stringifyRunInline(c as RunOptions, ctx));
          }
        }
      }
    }
    return `<w:customXml ${attrs.join(" ")}>${parts.join("")}</w:customXml>`;
  }

  return undefined;
}

/** Serialize children of Dir/Bdo containers. */
function serializeDirChildren(
  children: (RunOptions | string)[] | undefined,
  ctx: BodyContext,
): string {
  if (!children) return "";
  const parts: string[] = [];
  for (const c of children) {
    if (typeof c === "string") {
      parts.push(stringifyRunInline({ text: c }, ctx));
    } else {
      parts.push(stringifyRunInline(c, ctx));
    }
  }
  return parts.join("");
}

/** Hash SmartArt data for unique key generation (duplicated from SmartArtRun). */
function hashSmartArtData(options: SmartArtOptions): number {
  const data = JSON.stringify(options.data);
  let hash = 0;
  for (let i = 0; i < data.length; i++) {
    const char = data.charCodeAt(i);
    hash = ((hash << 5) - hash + char) | 0;
  }
  return Math.abs(hash);
}

// ── Paragraph ──

export function stringifyParagraphInline(
  opts: string | ParagraphOptions,
  ctx: BodyContext,
): string {
  const resolved: ParagraphOptions = typeof opts === "string" ? { text: opts } : opts;
  const parts: string[] = [];

  const props = stringifyParagraphProperties(resolved);
  if (props.xml) parts.push(props.xml);

  // Register numbering references from inline paragraphs (footnotes, endnotes, etc.)
  // so that concrete numbering instances are created and placeholders get resolved.
  if (props.numberingReferences.length > 0) {
    for (const ref of props.numberingReferences) {
      ctx.file.numbering.createConcreteNumberingInstance(ref.reference, ref.instance);
    }
  }

  if (resolved.text !== undefined) {
    parts.push(stringifyRunInline({ text: resolved.text }, ctx));
  }

  if (resolved.children) {
    for (const child of resolved.children) {
      if (typeof child === "string") {
        parts.push(stringifyRunInline({ text: child }, ctx));
      } else if (typeof child === "object" && child !== null) {
        // Try JSON child dispatch first (image, chart, hyperlink, etc.)
        const jsonResult = stringifyJsonChild(child as ParagraphChild, ctx);
        if (jsonResult !== undefined) {
          if (Array.isArray(jsonResult)) {
            parts.push(...jsonResult);
          } else {
            parts.push(jsonResult);
          }
        } else if ("text" in child || "children" in child || "break" in child) {
          parts.push(stringifyRunInline(child as RunOptions, ctx));
        }
      }
    }
  }

  const body = parts.join("");
  return body ? `<w:p>${body}</w:p>` : "<w:p/>";
}
