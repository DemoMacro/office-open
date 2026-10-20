/**
 * Shared inline run/paragraph stringification for DOCX descriptors.
 *
 * Used by table.ts, comments.ts, body.ts, and other descriptors that need to
 * serialize paragraph/run content. Includes JSON child dispatch for all
 * ParagraphChild variants (picture, chart, hyperlink, etc.).
 *
 * Pure string concatenation — no intermediate object tree.
 *
 * @module
 */

import { hpsMeasureValue } from "@office-open/core";
import { RELATIONSHIP_TYPES } from "@office-open/core/opc";
import { TargetModeType } from "@office-open/core/opc";
import { element, escapeXml } from "@office-open/xml";
import { subDocDesc } from "@parts/bodychildren";
import { objectDesc, type ObjectElementOptions } from "@parts/object";
import type {
  BookmarkOptions,
  BookmarkStartOptions,
  MarkupRangeOptions,
  MoveRangeOptions,
  MoveRangeStartOptions,
} from "@parts/paragraph/links/bookmark";
import type {
  ComplexFieldOptions,
  ComplexFieldInstructionMember,
  FieldMarkerOptions,
  FootnoteEndnoteReferenceOptions,
  ParagraphChild,
  ParagraphOptions,
  TrackChangeChild,
} from "@parts/paragraph/paragraph";
import type { NumberingChangeOptions } from "@parts/paragraph/properties";
import type { CommentChildOptions } from "@parts/paragraph/run/comment-run";
import type { RubyOptions } from "@parts/paragraph/run/ruby";
import {
  breakXml,
  EMPTY_RUN_ELEMENTS,
  type BreakOptions,
  type RunOptions,
} from "@parts/paragraph/run/run";
import type { SymbolRunOptions } from "@parts/paragraph/run/symbol-run";
import { pictRunAttrs, stringifyPict, type PictOptions } from "@parts/pict";
import type { ChangedProperties } from "@shared/track-revision/track-revision";
import { autoRevisionId } from "@shared/track-revision/track-revision";

import type { BodyContext } from "../context";
import { checkboxSymbolRunInner, stringifyCustomXmlShell, stringifySdtShell } from "./bodychildren";
import { stringifyDrawingChild } from "./inline-drawings";
import { stringifyMath, stringifyMathParagraph } from "./paragraph/math/stringify";
import { createBegin, createSeparate, createEnd } from "./paragraph/run/field";
import {
  onOff,
  paragraphIdentityAttrs,
  stringifyParagraphProperties,
  stringifyRunProperties,
} from "./paragraph/stringify";

// ── Run ──

function customMarkFollowsXml(
  reference: number | FootnoteEndnoteReferenceOptions | undefined,
): string {
  if (typeof reference !== "object" || reference.customMarkFollows === undefined) return "";
  return ` w:customMarkFollows="${reference.customMarkFollows ? "1" : "0"}"`;
}

function stringifySymbolRunInner(opts: SymbolRunOptions): string {
  const tag = opts.kind === "office2016" ? "w16se:sym" : "w:sym";
  return `<${tag} w:char="${escapeXml(opts.char)}" w:font="${escapeXml(opts.symbolFont ?? "Wingdings")}"/>`;
}

/** Serialize a complex field's run chain. Inside a w:del wrapper the
 *  instruction is spelled w:delInstrText and the cached result w:delText. */
function stringifySimpleField(
  sf: NonNullable<Extract<ParagraphChild, { simpleField: unknown }>["simpleField"]>,
  ctx: BodyContext,
): string {
  const sfAttrs = [`w:instr="${escapeXml(sf.instruction)}"`];
  if (sf.fieldLock !== undefined) sfAttrs.push(`w:fldLock="${sf.fieldLock ? 1 : 0}"`);
  if (sf.dirty !== undefined) sfAttrs.push(`w:dirty="${sf.dirty ? 1 : 0}"`);
  if (sf.cachedInstructionText !== undefined) {
    const space = sf.cachedInstructionTextPreserveSpace ? ' xml:space="preserve"' : "";
    return `<w:fldSimple ${sfAttrs.join(" ")}><w:r>${sf.cachedInstructionRPrXml ?? ""}<w:instrText${space}>${escapeXml(sf.cachedInstructionText)}</w:instrText></w:r></w:fldSimple>`;
  }
  if (sf.cachedRuns !== undefined) {
    return `<w:fldSimple ${sfAttrs.join(" ")}>${serializeDispatchChildren(sf.cachedRuns, ctx)}</w:fldSimple>`;
  }
  if (sf.cachedValue !== undefined) {
    const space = sf.cachedValuePreserveSpace ? ' xml:space="preserve"' : "";
    return `<w:fldSimple ${sfAttrs.join(" ")}><w:r><w:t${space}>${escapeXml(sf.cachedValue)}</w:t></w:r></w:fldSimple>`;
  }
  return `<w:fldSimple ${sfAttrs.join(" ")}/>`;
}

function stringifyComplexFieldMember(
  member: ComplexFieldInstructionMember,
  ctx: BodyContext,
  isDelete = false,
): string {
  if ("run" in member) {
    if (!isDelete) return stringifyRunInline(member.run, ctx);
    const children = member.run.children?.map((child) => {
      if (
        typeof child === "object" &&
        child !== null &&
        "instructionText" in child &&
        !("deletedInstructionText" in child)
      ) {
        const segment = child as { instructionText: string; preserveSpace?: boolean };
        const { instructionText, preserveSpace } = segment;
        return {
          deletedInstructionText: instructionText,
          ...(preserveSpace !== undefined ? { preserveSpace } : {}),
        };
      }
      return child;
    });
    return stringifyRunInline({ ...member.run, children }, ctx, true);
  }
  if ("simpleField" in member) return stringifySimpleField(member.simpleField, ctx);
  if ("bookmarkStart" in member) {
    return `<w:bookmarkStart ${buildBookmarkStartAttrs(member.bookmarkStart)}/>`;
  }
  if ("bookmarkEnd" in member) {
    return `<w:bookmarkEnd ${buildMarkupRangeAttrs(member.bookmarkEnd)}/>`;
  }
  if ("proofErr" in member) {
    return `<w:proofErr w:type="${member.proofErr}"/>`;
  }
  if ("hyperlink" in member) {
    const xml = stringifyChildDispatch(member, ctx);
    if (Array.isArray(xml)) return xml.join("");
    return xml ?? "";
  }
  if ("complexField" in member) {
    return stringifyComplexFieldRuns(member.complexField, isDelete, ctx);
  }
  if ("fieldMarker" in member) return stringifyFieldMarker(member.fieldMarker);
  if ("sdt" in member) {
    const xml = stringifyChildDispatch(member, ctx);
    if (Array.isArray(xml)) return xml.join("");
    return xml ?? "";
  }
  if (
    "insertion" in member ||
    "deletion" in member ||
    "movedFrom" in member ||
    "movedTo" in member
  ) {
    const xml = stringifyChildDispatch(member, ctx);
    if (Array.isArray(xml)) return xml.join("");
    return xml ?? "";
  }
  const xml = stringifyChildDispatch(member, ctx);
  if (xml === undefined) return "";
  return Array.isArray(xml) ? xml.join("") : xml;
}

function stringifyFieldMarker(marker: FieldMarkerOptions): string {
  const attrs = runAttrs([marker.additionRsid, marker.runPropertiesRsid, marker.deletionRsid]);
  const lrpb = marker.lastRenderedPageBreak ? "<w:lastRenderedPageBreak/>" : "";
  const numberingChangeXml = marker.numberingChange
    ? element("w:numberingChange", {
        "w:original": marker.numberingChange.original,
        "w:id": marker.numberingChange.id,
        "w:author": marker.numberingChange.author,
        "w:date": marker.numberingChange.date,
      })
    : "";
  const fldChar = numberingChangeXml
    ? `<w:fldChar w:fldCharType="${marker.type}">${numberingChangeXml}</w:fldChar>`
    : `<w:fldChar w:fldCharType="${marker.type}"/>`;
  return `<w:r${attrs}>${marker.rPrXml ?? ""}${lrpb}${fldChar}</w:r>`;
}

function stringifyComplexFieldRuns(
  cf: ComplexFieldOptions,
  isDelete = false,
  ctx: BodyContext,
): string {
  if (cf.compoundRunXml) return cf.compoundRunXml + (cf.compoundEndRunXml ?? "");
  const instrTag = isDelete ? "w:delInstrText" : "w:instrText";
  const textTag = isDelete ? "w:delText" : "w:t";
  // Run-properties: Word writes identical rPr across a field's runs. Apply
  // the captured control-run rPr to begin/instrText/separate/end and the
  // result-run rPr to the result (defaults to the control rPr when the
  // result had none, matching Word's uniform behavior).
  const ctrl = cf.rPrXml ?? "";
  const res = cf.resultRPrXml ?? ctrl;
  const beginAttrs = runAttrs([cf.additionRsid, cf.runPropertiesRsid, cf.deletionRsid]);
  const resultAttrs = [
    cf.resultAdditionRsid ? ` w:rsidR="${cf.resultAdditionRsid}"` : "",
    cf.resultRunPropertiesRsid ? ` w:rsidRPr="${cf.resultRunPropertiesRsid}"` : "",
    cf.resultDeletionRsid ? ` w:rsidDel="${cf.resultDeletionRsid}"` : "",
  ].join("");
  // Instruction: verbatim when the source split it across non-plain runs;
  // plain template otherwise; no instruction run at all for an empty
  // instruction (a bare begin→end field round-trips without one).
  const mixedInstructionXml = cf.instructionMembers
    ?.map((member) => stringifyComplexFieldMember(member, ctx, isDelete))
    .join("");
  const instrXml =
    mixedInstructionXml ??
    cf.instrRunsXml ??
    (cf.instruction !== ""
      ? `<w:r${runAttrs([cf.instructionAdditionRsid, cf.instructionRunPropertiesRsid, cf.instructionDeletionRsid])}>${cf.instructionRPrXml ?? ctrl}<${instrTag}${cf.instructionPreserveSpace ? ' xml:space="preserve"' : ""}>${escapeXml(cf.instruction)}</${instrTag}></w:r>`
      : "");
  // `separate` + the result run are emitted only when there is a cached
  // result or ordered result members; a result-less field round-trips as
  // begin/instrText/end. Result runs go verbatim when the source split them
  // beyond the plain template.
  const separatorXml =
    cf.hasResult === true ||
    cf.resultRunsXml !== undefined ||
    cf.resultMembers !== undefined ||
    cf.result !== undefined
      ? `<w:r${runAttrs([cf.separatorAdditionRsid, cf.separatorRunPropertiesRsid, cf.separatorDeletionRsid])}>${cf.separatorRPrXml ?? ctrl}<w:fldChar w:fldCharType="separate"/></w:r>`
      : "";
  const resultXml =
    cf.resultMembers !== undefined
      ? cf.resultMembers
          .map((member) => stringifyComplexFieldMember(member, ctx, isDelete))
          .join("")
      : cf.resultRunsXml !== undefined
        ? cf.resultRunsXml
        : cf.result !== undefined
          ? `<w:r${resultAttrs}>${res}<${textTag}${cf.resultPreserveSpace ? ' xml:space="preserve"' : ""}>${escapeXml(cf.result)}</${textTag}></w:r>`
          : "";
  const lrpb = cf.lastRenderedPageBreak ? "<w:lastRenderedPageBreak/>" : "";
  const endFldChar = endFieldCharXml(cf.endNumberingChange);
  return (
    `<w:r${beginAttrs}>${ctrl}${lrpb}${createBegin(
      undefined,
      undefined,
      cf.fieldLock,
      cf.fieldData,
      cf.beginNumberingChange,
    )}</w:r>` +
    instrXml +
    separatorXml +
    resultXml +
    `<w:r${runAttrs([cf.endAdditionRsid, cf.endRunPropertiesRsid, cf.endDeletionRsid])}>${cf.endRPrXml ?? ctrl}${endFldChar}</w:r>`
  );
}

function endFieldCharXml(change: NumberingChangeOptions | undefined): string {
  if (!change) return '<w:fldChar w:fldCharType="end"/>';
  const child = element("w:numberingChange", {
    "w:original": change.original,
    "w:id": change.id,
    "w:author": change.author,
    "w:date": change.date,
  });
  return `<w:fldChar w:fldCharType="end">${child}</w:fldChar>`;
}

function runAttrs(ids: readonly (string | undefined)[]): string {
  let attrs = "";
  if (ids[0]) attrs += ` w:rsidR="${ids[0]}"`;
  if (ids[1]) attrs += ` w:rsidRPr="${ids[1]}"`;
  if (ids[2]) attrs += ` w:rsidDel="${ids[2]}"`;
  return attrs;
}

function runIdentityAttrs(child: ParagraphChild): string {
  const attrs = child as {
    additionRsid?: string;
    runPropertiesRsid?: string;
    deletionRsid?: string;
  };
  return runAttrs([attrs.additionRsid, attrs.runPropertiesRsid, attrs.deletionRsid]);
}

function textElementXml(tag: string, text: string, preserve?: boolean): string {
  const preserveSpace = preserve === true || /^[\t\n\r ]|[\t\n\r ]$/.test(text);
  const attr = preserveSpace ? ' xml:space="preserve"' : "";
  return `<${tag}${attr}>${escapeXml(text)}</${tag}>`;
}

function stringifyRubyContent(opts: RubyOptions["text"], ctx: BodyContext): string {
  return (
    opts.children
      ?.map((child) => {
        if (typeof child === "string") return stringifyRunInline({ text: child }, ctx);
        if (
          "symbolRun" in child ||
          "object" in child ||
          "pict" in child ||
          !("style" in child && "text" in child)
        ) {
          return stringifyChildDispatch(child, ctx) ?? "";
        }
        return stringifyRunInline(child, ctx);
      })
      .join("") ?? ""
  );
}

function stringifyRuby(opts: RubyOptions, ctx: BodyContext): string {
  const properties = opts.properties;
  const dirty = properties.dirty !== undefined ? onOff("w:dirty", properties.dirty) : "";
  return (
    `<w:ruby><w:rubyPr>` +
    `<w:rubyAlign w:val="${properties.alignment}"/>` +
    `<w:hps w:val="${hpsMeasureValue(properties.fontSize * 2)}"/>` +
    `<w:hpsRaise w:val="${hpsMeasureValue(properties.raise * 2)}"/>` +
    `<w:hpsBaseText w:val="${hpsMeasureValue(properties.baseFontSize * 2)}"/>` +
    `<w:lid w:val="${escapeXml(properties.languageId)}"/>` +
    `${dirty}</w:rubyPr>` +
    `<w:rt>${stringifyRubyContent(opts.text, ctx)}</w:rt>` +
    `<w:rubyBase>${stringifyRubyContent(opts.base, ctx)}</w:rubyBase>` +
    `</w:ruby>`
  );
}

export function stringifyRunInline(opts: RunOptions, ctx: BodyContext, isDeleted = false): string {
  let body = "";
  const textTag = isDeleted ? "w:delText" : "w:t";

  const rPr = stringifyRunProperties(opts);

  if (opts.break) body += breakXml(opts.break);
  if (opts.pageBreak) body += '<w:br w:type="page"/>';
  if (opts.columnBreak) body += '<w:br w:type="column"/>';
  const hasSpaceSegment =
    opts.children?.some(
      (child) => typeof child === "object" && child !== null && "preserveSpace" in child,
    ) === true;
  // Top-level references — a pure reference run flattened by the parse path
  // (e.g. inside w:hyperlink) carries the reference alongside its rPr with no
  // children; without this branch the run would emit empty and drop it.
  const fnRef = opts.footnoteReference;
  const enRef = opts.endnoteReference;
  if (fnRef !== undefined) {
    const id = typeof fnRef === "number" ? fnRef : fnRef.id;
    const cmf = customMarkFollowsXml(fnRef);
    body += `<w:footnoteReference w:id="${id}"${cmf}/>`;
  } else if (enRef !== undefined) {
    const id = typeof enRef === "number" ? enRef : enRef.id;
    const cmf = customMarkFollowsXml(enRef);
    body += `<w:endnoteReference w:id="${id}"${cmf}/>`;
  }

  if (opts.children) {
    for (const child of opts.children) {
      if (typeof child === "string") {
        body += textElementXml(
          textTag,
          child,
          hasSpaceSegment ? /^[\t\n\r ]|[\t\n\r ]$/.test(child) : opts.preserveSpace,
        );
      } else if (typeof child === "object" && child !== null) {
        // Bare run-inner elements — emit directly inside this <w:r>. Must run
        // before stringifyChildDispatch, which wraps paragraph-level children
        // in their own <w:r> (correct for paragraphs, nested/invalid in a run).
        if ("tab" in child) {
          body += "<w:tab/>";
          continue;
        }
        const textSegment = child as { text?: string; preserveSpace?: boolean; nil?: true };
        if (
          typeof textSegment.text === "string" &&
          ("preserveSpace" in textSegment || "nil" in textSegment)
        ) {
          body += textSegment.nil
            ? `<${textTag} xsi:nil="true">${escapeXml(textSegment.text)}</${textTag}>`
            : textElementXml(textTag, textSegment.text, textSegment.preserveSpace);
          continue;
        }
        if ("instructionText" in textSegment) {
          const segment = textSegment as { instructionText: string; preserveSpace?: boolean };
          body += textElementXml("w:instrText", segment.instructionText, segment.preserveSpace);
          continue;
        }
        if ("deletedInstructionText" in textSegment) {
          const segment = textSegment as {
            deletedInstructionText: string;
            preserveSpace?: boolean;
          };
          body += textElementXml(
            "w:delInstrText",
            segment.deletedInstructionText,
            segment.preserveSpace,
          );
          continue;
        }
        if ("deletedText" in textSegment) {
          const segment = textSegment as { deletedText: string; preserveSpace?: boolean };
          body += textElementXml("w:delText", segment.deletedText, segment.preserveSpace);
          continue;
        }
        if ("pageBreak" in child) {
          body += '<w:br w:type="page"/>';
          continue;
        }
        if ("columnBreak" in child) {
          body += '<w:br w:type="column"/>';
          continue;
        }
        if ("break" in child) {
          body += breakXml((child as { break: number | BreakOptions }).break);
          continue;
        }
        if ("commentReference" in child) {
          body += `<w:commentReference w:id="${Number(child.commentReference)}"/>`;
          continue;
        }
        // Symbol run — emit bare; the dispatch branch wraps it in its own
        // <w:r>, which would nest illegally inside this run.
        if ("symbolRun" in child) {
          const sym = child.symbolRun;
          body += `${stringifyRunProperties(sym) ?? ""}${stringifySymbolRunInner(sym)}`;
          continue;
        }
        // Reference elements inside a run's children[] (mixed-content runs)
        // emit bare — the paragraph-level dispatch would wrap them in a
        // nested <w:r>, which is invalid inside a run.
        const bareFnRef = (child as RunOptions).footnoteReference;
        if (bareFnRef !== undefined) {
          const id = typeof bareFnRef === "number" ? bareFnRef : bareFnRef.id;
          const cmf = customMarkFollowsXml(bareFnRef);
          body += `<w:footnoteReference w:id="${id}"${cmf}/>`;
          continue;
        }
        const bareEnRef = (child as RunOptions).endnoteReference;
        if (bareEnRef !== undefined) {
          const id = typeof bareEnRef === "number" ? bareEnRef : bareEnRef.id;
          const cmf = customMarkFollowsXml(bareEnRef);
          body += `<w:endnoteReference w:id="${id}"${cmf}/>`;
          continue;
        }
        // Empty run elements — separator, noBreakHyphen, pgNum, etc.
        let firstKey: string | undefined;
        for (const key in child) {
          firstKey = key;
          break;
        }
        const emptyXml = firstKey !== undefined ? EMPTY_RUN_ELEMENTS[firstKey] : undefined;
        if (emptyXml) {
          body += emptyXml;
          continue;
        }
        if ("ruby" in child) {
          body += stringifyRuby((child as { ruby: RubyOptions }).ruby, ctx);
          continue;
        }
        // OLE object — w:object (VML shape + objectEmbed/link/control/movie)
        if ("object" in child) {
          // RunOptions.children carries Record<string, unknown>, so `in`
          // narrows to unknown — cast back to the object variant payload.
          body +=
            objectDesc.stringify((child as { object: ObjectElementOptions }).object, ctx) ?? "";
          continue;
        }
        // VML picture — w:pict emits bare inside this <w:r>
        if ("pict" in child) {
          body += stringifyPict((child as { pict: PictOptions }).pict, ctx);
          continue;
        }
        // Drawing-family members share the surrounding run when parsed from
        // mixed/multi-drawing runs; paragraph-level dispatch would nest a run.
        if (
          "picture" in child ||
          "chart" in child ||
          "chartEx" in child ||
          "smartArt" in child ||
          "wpsShape" in child ||
          "wpgGroup" in child ||
          "wordprocessingCanvas" in child ||
          "unsupportedDrawing" in child ||
          "contentPart" in child
        ) {
          body += stringifyDrawingChild(child as ParagraphChild, ctx, { bareInRun: true });
          continue;
        }
        // JSON child dispatch (images, charts, hyperlinks, etc.)
        const jsonResult = stringifyChildDispatch(child as ParagraphChild, ctx);
        if (jsonResult !== undefined) {
          body += Array.isArray(jsonResult) ? jsonResult.join("") : jsonResult;
        } else if ("text" in child || "children" in child || "break" in child) {
          body += stringifyRunInline(child as RunOptions, ctx, isDeleted);
        }
        // Anything else is an unknown field — silent no-op (runtime convention
        // for unrecognized JSON: TS excess-property checks guard authoring,
        // unknown runtime fields pass through without emitting).
      }
    }
  } else if (opts.text !== undefined) {
    body += textElementXml(textTag, String(opts.text), opts.preserveSpace);
  }

  let attr = "";
  if (opts.additionRsid) attr += ` w:rsidR="${opts.additionRsid}"`;
  if (opts.runPropertiesRsid) attr += ` w:rsidRPr="${opts.runPropertiesRsid}"`;
  if (opts.deletionRsid) attr += ` w:rsidDel="${opts.deletionRsid}"`;

  const orderedBody = opts.childOrder?.some((child) => child === "runProperties")
    ? opts.childOrder[0] === "runProperties"
      ? [rPr, body].filter(Boolean).join("")
      : [body, rPr].filter(Boolean).join("")
    : [rPr, body].filter(Boolean).join("");

  return orderedBody.length === 0
    ? attr
      ? `<w:r${attr}/>`
      : "<w:r/>"
    : `<w:r${attr}>${orderedBody}</w:r>`;
}

// ── JSON child dispatch ──

/**
 * Stringify a ParagraphChild into one or more XML strings.
 *
 * Handles side effects (media, chart, smartArt, relationship registration)
 * directly without creating temporary class instances.
 *
 * Returns `undefined` if the child is not a recognized JSON wrapper.
 */
/**
 * Build the rPr XML for a break/tab run from its structured run properties.
 */
/** Shared attribute string for CT_MarkupRange end markers (commentRange, move range end). */
export function buildMarkupRangeAttrs(m: MarkupRangeOptions): string {
  const a: string[] = [`w:id="${m.id}"`];
  if (m.displacedByCustomXml) a.push(`w:displacedByCustomXml="${m.displacedByCustomXml}"`);
  return a.join(" ");
}

/** Shared attribute string for w:bookmarkStart (CT_Bookmark). */
export function buildBookmarkStartAttrs(bs: BookmarkStartOptions): string {
  const a: string[] = [`w:id="${bs.id}"`, `w:name="${escapeXml(bs.name)}"`];
  if (bs.displacedByCustomXml) a.push(`w:displacedByCustomXml="${bs.displacedByCustomXml}"`);
  if (bs.colFirst !== undefined) a.push(`w:colFirst="${bs.colFirst}"`);
  if (bs.colLast !== undefined) a.push(`w:colLast="${bs.colLast}"`);
  return a.join(" ");
}

/** Shared attribute string for w:moveFromRangeStart / w:moveToRangeStart (CT_MoveBookmark). */
export function buildMoveRangeStartAttrs(m: MoveRangeStartOptions): string {
  const a: string[] = [`w:id="${m.id}"`];
  if (m.name) a.push(`w:name="${escapeXml(m.name)}"`);
  if (m.author) a.push(`w:author="${escapeXml(m.author)}"`);
  if (m.date) a.push(`w:date="${m.date}"`);
  if (m.displacedByCustomXml) a.push(`w:displacedByCustomXml="${m.displacedByCustomXml}"`);
  if (m.colFirst !== undefined) a.push(`w:colFirst="${m.colFirst}"`);
  if (m.colLast !== undefined) a.push(`w:colLast="${m.colLast}"`);
  return a.join(" ");
}

/** Revision wrapper attributes; an absent source date remains absent. */
function revisionElementAttrs(
  revision: ChangedProperties,
  scope: BodyContext["reproducible"],
): string {
  const id = revision.id ?? autoRevisionId(scope);
  const author = escapeXml(String(revision.author));
  const date = revision.date ? ` w:date="${escapeXml(revision.date)}"` : "";
  return ` w:id="${id}" w:author="${author}"${date}`;
}

/** Stringify inline run/text content — the `wrap` shared by every sugar child. */
function stringifyInlineWrap(wrap: (string | RunOptions)[] | undefined, ctx: BodyContext): string {
  const parts: string[] = [];
  for (const item of wrap ?? []) {
    parts.push(
      typeof item === "string"
        ? stringifyRunInline({ text: item }, ctx)
        : stringifyRunInline(item, ctx),
    );
  }
  return parts.join("");
}

/**
 * Serialize a break run inside a track-change wrapper. A parsed break run
 * keeps its run properties and rsid attributes — the same flat-props shape
 * the paragraph-level break variants carry.
 */
function stringifyTrackChangeBreak(opts: RunOptions, type: "page" | "column"): string {
  const rPr = stringifyRunProperties(opts) ?? "";
  let attr = "";
  if (opts.additionRsid) attr += ` w:rsidR="${opts.additionRsid}"`;
  if (opts.runPropertiesRsid) attr += ` w:rsidRPr="${opts.runPropertiesRsid}"`;
  if (opts.deletionRsid) attr += ` w:rsidDel="${opts.deletionRsid}"`;
  return `<w:r${attr}>${rPr}<w:br w:type="${type}"/></w:r>`;
}

/**
 * Serialize the children of a track-change wrapper: runs (delText inside
 * w:del, plain text inside w:ins), comment range markers, and nested
 * same-family wrappers — each child's text form follows its nearest wrapper.
 */
function stringifySdtTrackChangeContent(
  children: readonly (ParagraphChild | string)[] | undefined,
  ctx: BodyContext,
  isDelete: boolean,
): string {
  const parts: string[] = [];
  for (const child of children ?? []) {
    if (typeof child === "string") {
      parts.push(stringifyRunInline({ text: child }, ctx, isDelete));
      continue;
    }
    if ("complexField" in child) {
      parts.push(stringifyComplexFieldRuns(child.complexField, isDelete, ctx));
      continue;
    }
    if (isPlainTextRun(child)) {
      parts.push(stringifyRunInline(child, ctx, isDelete));
      continue;
    }
    const xml = stringifyChildDispatch(child, ctx);
    if (xml !== undefined) {
      parts.push(Array.isArray(xml) ? xml.join("") : xml);
    } else {
      parts.push(stringifyRunInline(child as RunOptions, ctx, isDelete));
    }
  }
  return parts.join("");
}

function stringifyTrackChangeChildren(
  children: readonly TrackChangeChild[],
  ctx: BodyContext,
  isDelete: boolean,
): string {
  const parts: string[] = [];
  for (const c of children) {
    if (typeof c !== "string" && "commentRangeStart" in c) {
      parts.push(`<w:commentRangeStart ${buildMarkupRangeAttrs(c.commentRangeStart)}/>`);
    } else if (typeof c !== "string" && "commentRangeEnd" in c) {
      parts.push(`<w:commentRangeEnd ${buildMarkupRangeAttrs(c.commentRangeEnd)}/>`);
    } else if (typeof c !== "string" && "insertion" in c) {
      const { id, author, date, children: nested } = c.insertion;
      parts.push(
        `<w:ins${revisionElementAttrs({ id, author, date }, ctx.reproducible)}>` +
          stringifyTrackChangeChildren(nested, ctx, false) +
          "</w:ins>",
      );
    } else if (typeof c !== "string" && "deletion" in c) {
      const { id, author, date, children: nested } = c.deletion;
      parts.push(
        `<w:del${revisionElementAttrs({ id, author, date }, ctx.reproducible)}>` +
          stringifyTrackChangeChildren(nested, ctx, true) +
          "</w:del>",
      );
    } else if (typeof c !== "string" && "pageBreak" in c) {
      parts.push(stringifyTrackChangeBreak(c as RunOptions, "page"));
    } else if (typeof c !== "string" && "columnBreak" in c) {
      parts.push(stringifyTrackChangeBreak(c as RunOptions, "column"));
    } else if (typeof c !== "string" && "proofErr" in c) {
      parts.push(`<w:proofErr w:type="${c.proofErr}"/>`);
    } else if (typeof c !== "string" && "complexField" in c) {
      parts.push(stringifyComplexFieldRuns(c.complexField, isDelete, ctx));
    } else if (typeof c !== "string" && "sdt" in c) {
      const sdt = c.sdt;
      parts.push(
        stringifySdtShell(
          sdt.properties,
          sdt.endProperties,
          stringifySdtTrackChangeContent(sdt.children, ctx, isDelete),
          ctx.reproducible,
        ),
      );
    } else if (typeof c !== "string" && "formField" in c) {
      const xml = stringifyChildDispatch(c as ParagraphChild, ctx);
      if (typeof xml === "string") parts.push(xml);
    } else if (typeof c !== "string") {
      // A deleted plain-text run must become w:delText here, before dispatch —
      // dispatch is blind to the track-change context and its text fast path
      // would emit a plain w:t run. Everything else keeps the old route:
      // dispatch builds the complete <w:r> for drawings and other variants,
      // and plain-run shapes it does not recognize take the wrapper-aware
      // fallback below.
      if (isDelete && isPlainTextRun(c)) {
        parts.push(stringifyRunInline(c, ctx, true));
      } else {
        const xml = stringifyChildDispatch(c as ParagraphChild, ctx);
        if (typeof xml === "string") {
          parts.push(xml);
        } else {
          parts.push(
            isDelete
              ? stringifyRunInline(c as RunOptions, ctx, true)
              : stringifyRunInline(c as RunOptions, ctx),
          );
        }
      }
    } else {
      // Plain string child — inside w:del it serializes as delText.
      parts.push(
        isDelete
          ? stringifyRunInline({ text: c }, ctx, true)
          : stringifyRunInline({ text: c }, ctx),
      );
    }
  }
  return parts.join("");
}

/**
 * Expand a `{ comment }` sugar child: allocate the comment id, register the
 * comment entry (side effect, consumed when word/comments.xml is stringified),
 * and emit the range markers + anchored content + reference with one shared id.
 *
 * The caller never supplies an id — the library owns id allocation and pairing.
 */
function stringifyCommentChild(c: CommentChildOptions, ctx: BodyContext): string {
  const id = ctx.file.comments.nextId++;
  ctx.file.comments.entries.push({
    id,
    author: c.author,
    initials: c.initials,
    date: c.date,
    children: c.children,
  });

  return (
    `<w:commentRangeStart w:id="${id}"/>` +
    stringifyInlineWrap(c.wrap, ctx) +
    `<w:commentRangeEnd w:id="${id}"/>` +
    `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`
  );
}

/**
 * Expand a `{ bookmark }` sugar child: allocate the bookmark id and emit the
 * paired bookmarkStart/bookmarkEnd with the anchored content between them.
 * Bookmarks are pure markup — the only effect is the two markers.
 */
function stringifyBookmarkChild(b: BookmarkOptions, ctx: BodyContext): string {
  const id = ctx.file.markupIds.rangeNext++;
  const startAttrs = buildBookmarkStartAttrs({
    id,
    name: b.name,
    displacedByCustomXml: b.displacedByCustomXml,
    colFirst: b.colFirst,
    colLast: b.colLast,
  });
  const endAttrs = buildMarkupRangeAttrs({ id, displacedByCustomXml: b.displacedByCustomXml });
  return `<w:bookmarkStart ${startAttrs}/>${stringifyInlineWrap(b.wrap, ctx)}<w:bookmarkEnd ${endAttrs}/>`;
}

/**
 * Expand a `{ moveFrom }` / `{ moveTo }` sugar child: allocate the range id and
 * the move-run id, then emit the paired range markers with the moved run between
 * them. The move run (CT_TrackChange) carries the moved content.
 */
function stringifyMoveRangeChild(
  kind: "moveFrom" | "moveTo",
  opts: MoveRangeOptions,
  ctx: BodyContext,
): string {
  const rangeId = ctx.file.markupIds.rangeNext++;
  const runId = ctx.file.markupIds.moveRunNext++;
  const isMoveFrom = kind === "moveFrom";
  const startTag = isMoveFrom ? "w:moveFromRangeStart" : "w:moveToRangeStart";
  const endTag = isMoveFrom ? "w:moveFromRangeEnd" : "w:moveToRangeEnd";
  const runTag = isMoveFrom ? "w:moveFrom" : "w:moveTo";
  const rangeStartAttrs = buildMoveRangeStartAttrs({
    id: rangeId,
    name: opts.name,
    author: opts.author,
    date: opts.date,
    displacedByCustomXml: opts.displacedByCustomXml,
    colFirst: opts.colFirst,
    colLast: opts.colLast,
  });
  const endAttrs = buildMarkupRangeAttrs({
    id: rangeId,
    displacedByCustomXml: opts.displacedByCustomXml,
  });
  return (
    `<${startTag} ${rangeStartAttrs}/>` +
    `<${runTag} w:id="${runId}" w:author="${escapeXml(opts.author)}" w:date="${opts.date}">${stringifyInlineWrap(opts.wrap, ctx)}</${runTag}>` +
    `<${endTag} ${endAttrs}/>`
  );
}

function runPropertiesXml(child: ParagraphChild): string {
  return stringifyRunProperties(child as RunOptions) ?? "";
}

/**
 * The plain-run shape the dispatch text fast path claims: text set, with none
 * of the variants that legally coexist with text and have their own dispatch
 * branch (references, the { text, hyperlink } shorthand). Callers that need
 * context-specific text spelling (w:delText inside a deletion wrapper) use
 * this to claim the shape before dispatch does.
 */
function isPlainTextRun(child: ParagraphChild): child is RunOptions {
  const run = child as RunOptions;
  return (
    "text" in child &&
    run.text !== undefined &&
    run.footnoteReference === undefined &&
    run.endnoteReference === undefined &&
    !("commentReference" in child) &&
    !("hyperlink" in child)
  );
}

export function stringifyChildDispatch(
  child: ParagraphChild,
  ctx: BodyContext,
): string | string[] | undefined {
  // Verbatim passthrough (unrecognized drawings and other shapes the parse
  // path keeps byte-faithful) — comes first so no run-like branch claims it.
  if ("rawXml" in child) {
    return child.rawXml;
  }
  // Fast path: a plain run with text set is by far the most common child.
  // The chain below otherwise runs ~40 discriminant checks before the caller
  // falls back to stringifyRunInline — the exact function this returns.
  if (isPlainTextRun(child)) {
    return stringifyRunInline(child as RunOptions, ctx);
  }
  // Simple break types — pure XML, no side effects. A break run may carry run
  // properties (round-tripped from <w:r><w:rPr>…</w:rPr><w:br…/></w:r>).
  if ("pageBreak" in child) {
    return `<w:r${runIdentityAttrs(child)}>${runPropertiesXml(child)}<w:br w:type="page"/></w:r>`;
  }
  if ("columnBreak" in child) {
    return `<w:r${runIdentityAttrs(child)}>${runPropertiesXml(child)}<w:br w:type="column"/></w:r>`;
  }
  if ("tab" in child) {
    return `<w:r${runIdentityAttrs(child)}>${runPropertiesXml(child)}<w:tab/></w:r>`;
  }
  // Sub-document insertion (w:subDoc, EG_PContent member like hyperlink — no
  // run wrapper, the element sits directly in the paragraph)
  if ("subDoc" in child) {
    return subDocDesc.stringify(child.subDoc, ctx) ?? "";
  }

  // Reference types — pure XML, no side effects
  const fnRefChild = (child as RunOptions).footnoteReference;
  if (fnRefChild !== undefined) {
    const id = typeof fnRefChild === "number" ? fnRefChild : fnRefChild.id;
    const cmf = customMarkFollowsXml(fnRefChild);
    // Round-tripped run properties win; a fresh reference gets the
    // conventional FootnoteReference character style. A RunOptions-flavored
    // child carries them flattened, so read both shapes.
    const props = "properties" in child ? child.properties : undefined;
    const rPr = props
      ? (stringifyRunProperties(props) ?? "")
      : '<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>';
    return `<w:r${runIdentityAttrs(child)}>${rPr}<w:footnoteReference w:id="${id}"${cmf}/></w:r>`;
  }
  const enRefChild = (child as RunOptions).endnoteReference;
  if (enRefChild !== undefined) {
    const id = typeof enRefChild === "number" ? enRefChild : enRefChild.id;
    const cmf = customMarkFollowsXml(enRefChild);
    const props = "properties" in child ? child.properties : undefined;
    const rPr = props
      ? (stringifyRunProperties(props) ?? "")
      : '<w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>';
    return `<w:r${runIdentityAttrs(child)}>${rPr}<w:endnoteReference w:id="${id}"${cmf}/></w:r>`;
  }

  // Comment sugar — library allocates the id, emits the range markers +
  // reference, and registers the comment entry (see stringifyCommentChild).
  if ("comment" in child) return stringifyCommentChild(child.comment, ctx);

  // Comment markers — pure XML
  if ("commentRangeStart" in child)
    return `<w:commentRangeStart ${buildMarkupRangeAttrs(child.commentRangeStart)}/>`;
  if ("commentRangeEnd" in child)
    return `<w:commentRangeEnd ${buildMarkupRangeAttrs(child.commentRangeEnd)}/>`;
  if ("commentReference" in child) {
    // Run properties parsed with the reference survive verbatim — Word does
    // not always style comment references, so nothing is injected.
    const rPr = child.properties ? (stringifyRunProperties(child.properties) ?? "") : "";
    return `<w:r${runIdentityAttrs(child)}>${rPr}<w:commentReference w:id="${child.commentReference}"/></w:r>`;
  }
  // Bookmark markers — pure XML
  if ("bookmarkStart" in child) {
    return `<w:bookmarkStart ${buildBookmarkStartAttrs(child.bookmarkStart)}/>`;
  }
  if ("bookmarkEnd" in child) {
    return `<w:bookmarkEnd ${buildMarkupRangeAttrs(child.bookmarkEnd)}/>`;
  }
  // Bookmark sugar — library allocates the id and pairs start/end.
  if ("bookmark" in child) return stringifyBookmarkChild(child.bookmark, ctx);

  // Symbol runs are self-closing elements, not text: emit them directly so
  // they are not escaped into a <w:t> by the run children path.
  if ("symbolRun" in child) {
    const opts = child.symbolRun;
    const rPr = stringifyRunProperties(opts) ?? "";
    return `<w:r${runIdentityAttrs(opts)}>${rPr}${stringifySymbolRunInner(opts)}</w:r>`;
  }

  // OLE object run — the parse path flattens a pure w:object run to a bare
  // { object } paragraph child (with run properties merged in), so emit its
  // own <w:r> here rather than falling through to the run-children path.
  if ("object" in child) {
    const rPr = stringifyRunProperties(child as RunOptions) ?? "";
    return `<w:r${runIdentityAttrs(child as RunOptions)}>${rPr}${objectDesc.stringify(
      child.object,
      ctx,
    )}</w:r>`;
  }

  // VML picture run — same flattening: a bare { pict } child carries its run
  // properties merged in, and w:pict is emitted inside its own <w:r>.
  if ("pict" in child) {
    const rPr = stringifyRunProperties(child as RunOptions) ?? "";
    return `<w:r${pictRunAttrs(child.pict)}>${rPr}${stringifyPict(child.pict, ctx)}</w:r>`;
  }

  // Form field (checkbox / dropdown list / text input) — fldChar sequence.
  // Word needs the field code (instrText) between begin and separate to
  // recognize the field type and render its result.
  if ("formField" in child) {
    const ff = child.formField;
    const ctrl = child.rPrXml ?? "";
    const resultRPr = child.resultRPrXml ?? "";
    const endRPr = child.endRPrXml ?? ctrl;
    let result = "";
    let instrCode = "";
    let symbolFont = false;
    if (ff.checkBox) {
      result = ff.checkBox.checked ? "☒" : "☐";
      instrCode = "FORMCHECKBOX";
      // U+2610/U+2612 are absent from common body fonts (Calibri/Times);
      // MS Gothic holds them and matches the SDT w14:checkbox default.
      symbolFont = true;
    } else if (ff.dropDownList) {
      const idx = ff.dropDownList.result ?? ff.dropDownList.default;
      result = idx !== undefined ? (ff.dropDownList.entries[idx] ?? "") : "";
      instrCode = "FORMDROPDOWN";
    } else if (ff.textInput) {
      // Prefer the user-entered value (result run) over the placeholder default.
      result = ff.textInput.value ?? ff.textInput.default ?? "";
      instrCode = "FORMTEXT";
    }
    const rPr = symbolFont
      ? '<w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/></w:rPr>'
      : "";
    const instructionRPr = child.instructionRPrXml ?? ctrl;
    const instructionXml =
      child.instructionMembers
        ?.map((member) => stringifyComplexFieldMember(member, ctx))
        .join("") ??
      child.instrRunsXml ??
      `<w:r${runAttrs([child.instructionAdditionRsid, child.instructionRunPropertiesRsid])}>${instructionRPr}<w:instrText xml:space="preserve"> ${instrCode} </w:instrText></w:r>`;
    const resultXml =
      child.hasResult === false
        ? ""
        : `<w:r${runAttrs([child.separatorAdditionRsid, child.separatorRunPropertiesRsid])}>${child.separatorRPrXml ?? ctrl}${createSeparate()}</w:r>` +
          (child.resultMembers
            ? child.resultMembers.map((member) => stringifyComplexFieldMember(member, ctx)).join("")
            : (child.resultRunsXml ??
              `<w:r${runAttrs([child.resultAdditionRsid, child.resultRunPropertiesRsid])}>${child.resultRPrXml ?? (resultRPr || rPr)}<w:t xml:space="preserve">${escapeXml(result)}</w:t></w:r>`));
    return (
      `<w:r${runAttrs([child.additionRsid, child.runPropertiesRsid])}>${ctrl}${child.lastRenderedPageBreak ? "<w:lastRenderedPageBreak/>" : ""}${createBegin(
        undefined,
        ff,
        child.fieldLock,
      )}</w:r>` +
      instructionXml +
      resultXml +
      `<w:r${runAttrs([child.endAdditionRsid, child.endRunPropertiesRsid])}>${endRPr}${createEnd()}</w:r>`
    );
  }

  // Drawing-family children — picture, chart, SmartArt, WPS shape,
  // contentPart, and wpgGroup (media/part registration + descriptor
  // stringification live in ./inline-drawings).
  const drawingXml = stringifyDrawingChild(child, ctx);
  if (drawingXml !== undefined) return drawingXml;

  // Math — pure string concatenation. A justification or the display flag
  // wraps the equation in a display m:oMathPara; without it stays inline m:oMath.
  if ("math" in child && typeof child.math === "object" && child.math !== null) {
    const mathOpts = child.math;
    const equations = mathOpts.equations;
    if (mathOpts.display || mathOpts.justification !== undefined) {
      return stringifyMathParagraph(equations, mathOpts.justification, ctx);
    }
    return equations.map((equation) => stringifyMath(equation, ctx)).join("");
  }

  // Inserted text run(s) — w:ins wraps one or more runs (CT_RunTrackChange)
  if ("insertion" in child) {
    const { id, author, date, children } = child.insertion;
    const body = stringifyTrackChangeChildren(children, ctx, false);
    return `<w:ins${revisionElementAttrs({ id, author, date }, ctx.reproducible)}>${body}</w:ins>`;
  }

  // Deleted text run(s) — w:del wraps one or more runs (delText content)
  if ("deletion" in child) {
    const { id, author, date, children } = child.deletion;
    const body = stringifyTrackChangeChildren(children, ctx, true);
    return `<w:del${revisionElementAttrs({ id, author, date }, ctx.reproducible)}>${body}</w:del>`;
  }

  // Hyperlink — side effect: relationship registration
  if ("hyperlink" in child) {
    const hl = child.hyperlink;

    // Serialize children using the same dispatch as paragraph children so
    // reference runs, drawings and objects inside hyperlinks keep their
    // full-run emission ({ footnoteReference, properties } carries the rPr in
    // a nested field only the dispatch branch reads). Plain runs fall back to
    // stringifyRunInline. A top-level `text` is a shorthand for a single text
    // run; without it `{ text, hyperlink }` would emit an empty <w:hyperlink>.
    const childParts: string[] = [];
    if (child.text !== undefined) {
      childParts.push(stringifyRunInline({ text: child.text }, ctx));
    }
    if (hl.children) {
      for (const rc of hl.children) {
        if (typeof rc === "string") {
          childParts.push(stringifyRunInline({ text: rc }, ctx));
        } else {
          const jr = stringifyChildDispatch(rc as ParagraphChild, ctx);
          childParts.push(
            jr !== undefined
              ? Array.isArray(jr)
                ? jr.join("")
                : jr
              : stringifyRunInline(rc as RunOptions, ctx),
          );
        }
      }
    }
    const body = childParts.join("");

    const pushHlAttrs = (attrs: string[]): void => {
      // Presence-based: an unset history keeps the source's attribute-less
      // form (injecting "1" here would oscillate across re-generations —
      // absent parses back as undefined, which would then inject "1").
      if (hl.history !== undefined) attrs.push(`w:history="${hl.history ? "1" : "0"}"`);
      if (hl.anchor !== undefined) attrs.push(`w:anchor="${escapeXml(hl.anchor)}"`);
      if (hl.tooltip !== undefined) attrs.push(`w:tooltip="${escapeXml(hl.tooltip)}"`);
      if (hl.targetFrame !== undefined) attrs.push(`w:tgtFrame="${escapeXml(hl.targetFrame)}"`);
      if (hl.docLocation !== undefined) attrs.push(`w:docLocation="${escapeXml(hl.docLocation)}"`);
    };
    if (hl.url) {
      // Auto-allocated sequential id — deterministic across runs, so a
      // re-generated document keeps byte-stable relationships (random ids
      // would drift on every compile). Later media/embedding offsets sample
      // relationshipCount after this registration, so no id collision.
      // One relationship per hyperlink element — Word emits a distinct rel
      // per reference even when the URL repeats (sharing would collapse the
      // source's per-reference entries on round-trip).
      const relType = RELATIONSHIP_TYPES.hyperlink;
      let linkId: string;
      if (hl.sourceRelationshipId === undefined) {
        // Fresh authoring: one relationship per hyperlink element — Word emits
        // a distinct rel per reference even when the URL repeats.
        linkId = `rId${ctx.viewWrapper.relationships.add(
          relType,
          hl.url,
          TargetModeType.EXTERNAL,
        )}`;
      } else if (!ctx.viewWrapper.relationships.hasId(hl.sourceRelationshipId)) {
        // Round-trip: the source id slot is free — re-emit at that exact id so
        // per-reference entries keep the source topology.
        ctx.viewWrapper.relationships.addRelationship(
          hl.sourceRelationshipId,
          relType,
          hl.url,
          TargetModeType.EXTERNAL,
        );
        linkId = hl.sourceRelationshipId;
      } else {
        // The slot is taken. When the occupant is the same kind+target the
        // source shared one rel across references — reuse it. Otherwise the id
        // was reassigned and this reference allocates a fresh rel.
        const shared = ctx.viewWrapper.relationships.idOf(relType, hl.url, TargetModeType.EXTERNAL);
        linkId =
          shared ??
          `rId${ctx.viewWrapper.relationships.add(relType, hl.url, TargetModeType.EXTERNAL)}`;
      }
      const attrs = [`r:id="${linkId}"`];
      pushHlAttrs(attrs);
      return `<w:hyperlink ${attrs.join(" ")}>${body}</w:hyperlink>`;
    }
    if (hl.anchor !== undefined) {
      const attrs: string[] = [];
      pushHlAttrs(attrs);
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
    const a: string[] = [`w:id="${ps.id ?? autoRevisionId(ctx.reproducible)}"`];
    if (ps.editor !== undefined) a.push(`w:ed="${escapeXml(String(ps.editor))}"`);
    if (ps.editGroup !== undefined) a.push(`w:edGrp="${ps.editGroup}"`);
    if (ps.colFirst !== undefined) a.push(`w:colFirst="${ps.colFirst}"`);
    if (ps.colLast !== undefined) a.push(`w:colLast="${ps.colLast}"`);
    return `<w:permStart ${a.join(" ")}/>`;
  }
  if ("permEnd" in child) return `<w:permEnd w:id="${child.permEnd}"/>`;

  // ── Move revision range markers ──
  if ("moveFromRangeStart" in child) {
    return `<w:moveFromRangeStart ${buildMoveRangeStartAttrs(child.moveFromRangeStart)}/>`;
  }
  if ("moveFromRangeEnd" in child)
    return `<w:moveFromRangeEnd ${buildMarkupRangeAttrs(child.moveFromRangeEnd)}/>`;
  if ("moveToRangeStart" in child) {
    return `<w:moveToRangeStart ${buildMoveRangeStartAttrs(child.moveToRangeStart)}/>`;
  }
  if ("moveToRangeEnd" in child)
    return `<w:moveToRangeEnd ${buildMarkupRangeAttrs(child.moveToRangeEnd)}/>`;
  // Move revision sugar — library allocates range + run ids and pairs markers.
  if ("moveFrom" in child) return stringifyMoveRangeChild("moveFrom", child.moveFrom, ctx);
  if ("moveTo" in child) return stringifyMoveRangeChild("moveTo", child.moveTo, ctx);

  // ── Move revision text runs ──
  if ("movedFrom" in child) {
    const { id, author, date, children } = child.movedFrom;
    const body = stringifyTrackChangeChildren(children, ctx, false);
    return `<w:moveFrom${revisionElementAttrs({ id, author, date }, ctx.reproducible)}>${body}</w:moveFrom>`;
  }
  if ("movedTo" in child) {
    const { id, author, date, children } = child.movedTo;
    const body = stringifyTrackChangeChildren(children, ctx, false);
    return `<w:moveTo${revisionElementAttrs({ id, author, date }, ctx.reproducible)}>${body}</w:moveTo>`;
  }

  // ── Merge-conflict text runs (w14, Word 2010+) ──
  if ("conflictIns" in child) {
    const { id, author, date, children } = child.conflictIns;
    const body = stringifyTrackChangeChildren(children, ctx, false);
    return `<w14:conflictIns${revisionElementAttrs({ id, author, date }, ctx.reproducible)}>${body}</w14:conflictIns>`;
  }
  if ("conflictDel" in child) {
    const { id, author, date, children } = child.conflictDel;
    const body = stringifyTrackChangeChildren(children, ctx, true);
    return `<w14:conflictDel${revisionElementAttrs({ id, author, date }, ctx.reproducible)}>${body}</w14:conflictDel>`;
  }

  // ── Custom XML range markers (track changes) ──
  if ("customXmlInsRangeStart" in child) {
    const o = child.customXmlInsRangeStart;
    return `<w:customXmlInsRangeStart w:id="${o.id}"${o.author ? ` w:author="${escapeXml(o.author)}"` : ""}${o.date ? ` w:date="${o.date}"` : ""}/>`;
  }
  if ("customXmlInsRangeEnd" in child)
    return `<w:customXmlInsRangeEnd w:id="${child.customXmlInsRangeEnd}"/>`;
  if ("customXmlDelRangeStart" in child) {
    const o = child.customXmlDelRangeStart;
    return `<w:customXmlDelRangeStart w:id="${o.id}"${o.author ? ` w:author="${escapeXml(o.author)}"` : ""}${o.date ? ` w:date="${o.date}"` : ""}/>`;
  }
  if ("customXmlDelRangeEnd" in child)
    return `<w:customXmlDelRangeEnd w:id="${child.customXmlDelRangeEnd}"/>`;
  if ("customXmlMoveFromRangeStart" in child) {
    const o = child.customXmlMoveFromRangeStart;
    return `<w:customXmlMoveFromRangeStart w:id="${o.id}"${o.author ? ` w:author="${escapeXml(o.author)}"` : ""}${o.date ? ` w:date="${o.date}"` : ""}/>`;
  }
  if ("customXmlMoveFromRangeEnd" in child)
    return `<w:customXmlMoveFromRangeEnd w:id="${child.customXmlMoveFromRangeEnd}"/>`;
  if ("customXmlMoveToRangeStart" in child) {
    const o = child.customXmlMoveToRangeStart;
    return `<w:customXmlMoveToRangeStart w:id="${o.id}"${o.author ? ` w:author="${escapeXml(o.author)}"` : ""}${o.date ? ` w:date="${o.date}"` : ""}/>`;
  }
  if ("customXmlMoveToRangeEnd" in child)
    return `<w:customXmlMoveToRangeEnd w:id="${child.customXmlMoveToRangeEnd}"/>`;

  // ── Merge-conflict range markers (w14, Word 2010+) ──
  if ("customXmlConflictInsRangeStart" in child) {
    const o = child.customXmlConflictInsRangeStart;
    return `<w14:customXmlConflictInsRangeStart w:id="${o.id}"${o.author ? ` w:author="${escapeXml(o.author)}"` : ""}${o.date ? ` w:date="${o.date}"` : ""}/>`;
  }
  if ("customXmlConflictInsRangeEnd" in child)
    return `<w14:customXmlConflictInsRangeEnd w:id="${child.customXmlConflictInsRangeEnd}"/>`;
  if ("customXmlConflictDelRangeStart" in child) {
    const o = child.customXmlConflictDelRangeStart;
    return `<w14:customXmlConflictDelRangeStart w:id="${o.id}"${o.author ? ` w:author="${escapeXml(o.author)}"` : ""}${o.date ? ` w:date="${o.date}"` : ""}/>`;
  }
  if ("customXmlConflictDelRangeEnd" in child)
    return `<w14:customXmlConflictDelRangeEnd w:id="${child.customXmlConflictDelRangeEnd}"/>`;

  // ── Simple field ──
  if ("simpleField" in child) {
    return stringifySimpleField(child.simpleField, ctx);
  }

  // ── Complex field (PAGE/DATE/TOC/... — fldChar field without w:ffData) ──
  if ("complexField" in child) {
    return stringifyComplexFieldRuns(child.complexField, false, ctx);
  }

  if ("fieldMarker" in child) {
    return stringifyFieldMarker(child.fieldMarker);
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
    const childXml = serializeDispatchChildren(d.children, ctx);
    return `<w:dir w:val="${d.val}">${childXml}</w:dir>`;
  }
  if ("bdo" in child) {
    const b = child.bdo;
    const childXml = serializeDispatchChildren(b.children, ctx);
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
    parts.push(serializeDispatchChildren(st.children, ctx));
    return `<w:smartTag ${attrs.join(" ")}>${parts.join("")}</w:smartTag>`;
  }

  // ── Custom XML run (CT_CustomXmlRun) ──
  if ("customXml" in child) {
    const cx = child.customXml;
    return stringifyCustomXmlShell(cx, serializeDispatchChildren(cx.children, ctx));
  }

  // ── Inline structured document tag (CT_SdtRun) ──
  if ("sdt" in child) {
    const s = child.sdt;
    let contentXml = "";
    if (s.properties.checkbox && !s.children?.length) {
      // Inline checkbox: render the state symbol as a run (no <w:p> wrapper).
      contentXml = checkboxSymbolRunInner(s.properties.checkbox);
    } else if (s.children && s.children.length > 0) {
      const cparts: string[] = [];
      for (const c of s.children) {
        if (typeof c === "string") {
          cparts.push(stringifyRunInline({ text: c }, ctx));
        } else {
          const jr = stringifyChildDispatch(c as ParagraphChild, ctx);
          if (jr !== undefined) {
            cparts.push(Array.isArray(jr) ? jr.join("") : jr);
          } else {
            cparts.push(stringifyRunInline(c as RunOptions, ctx));
          }
        }
      }
      contentXml = cparts.join("");
    }
    return stringifySdtShell(s.properties, s.endProperties, contentXml, ctx.reproducible);
  }

  return undefined;
}

/**
 * Serialize `(ParagraphChild | string)[]` content: each child goes through the
 * JSON dispatch, with an unconditional run fallback for unrecognized wrappers
 * (shared by Dir/Bdo, smartTag and customXml). Sdt and paragraph children use
 * their own loops — their fallback-drop semantics differ.
 */
function serializeDispatchChildren(
  children: (ParagraphChild | string)[] | undefined,
  ctx: BodyContext,
): string {
  if (!children) return "";
  let body = "";
  for (const c of children) {
    if (typeof c === "string") {
      body += stringifyRunInline({ text: c }, ctx);
      continue;
    }
    const jr = stringifyChildDispatch(c, ctx);
    body +=
      jr !== undefined
        ? Array.isArray(jr)
          ? jr.join("")
          : jr
        : stringifyRunInline(c as RunOptions, ctx);
  }
  return body;
}

// ── Paragraph ──

export function stringifyParagraphInline(
  opts: string | ParagraphOptions,
  ctx: BodyContext,
): string {
  const resolved: ParagraphOptions = typeof opts === "string" ? { text: opts } : opts;
  let body = "";

  const props = stringifyParagraphProperties(resolved, ctx.reproducible);
  if (props.xml) body += props.xml;

  // Register numbering references from inline paragraphs (footnotes, endnotes, etc.)
  // so that concrete numbering instances are created and placeholders get resolved.
  if (props.numberingReferences.length > 0) {
    for (const ref of props.numberingReferences) {
      ctx.file.numbering.createConcreteNumberingInstance(ref.reference, ref.instance);
    }
  }

  if (resolved.text !== undefined) {
    body += stringifyRunInline({ text: resolved.text }, ctx);
  }

  if (resolved.children) {
    for (const child of resolved.children) {
      if (typeof child === "string") {
        body += stringifyRunInline({ text: child }, ctx);
      } else if (typeof child === "object" && child !== null) {
        if ("table" in child) {
          body += ctx.stringifyChild({ table: child.table });
          continue;
        }
        if ("blockBreak" in child) {
          body += breakXml(child.blockBreak);
          continue;
        }
        // Try JSON child dispatch first (image, chart, hyperlink, etc.)
        const jsonResult = stringifyChildDispatch(child as ParagraphChild, ctx);
        if (jsonResult !== undefined) {
          body += Array.isArray(jsonResult) ? jsonResult.join("") : jsonResult;
        } else {
          // RunOptions-like plain object — may be an empty run carrying only
          // run properties (round-tripped from <w:r><w:rPr>…</w:rPr></w:r>);
          // same fallback contract as the body-level paragraph stringifier.
          body += stringifyRunInline(child as RunOptions, ctx);
        }
      }
    }
  }

  // Identity attributes — inline containers (comments, footnotes, headers)
  // round-trip the body-level identity set too; w15:commentEx links comment
  // replies through w14:paraId.
  const attr = paragraphIdentityAttrs(resolved);
  return body ? `<w:p${attr}>${body}</w:p>` : `<w:p${attr}/>`;
}
