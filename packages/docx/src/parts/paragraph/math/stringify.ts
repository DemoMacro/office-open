/**
 * Direct XML string builders for Office MathML (OMML).
 *
 * Replaces `coerceMathInput()` + `new Math().toXml()` recursive class chain
 * with direct string concatenation — zero XmlComponent instances.
 *
 * Processes `MathInput` discriminated union directly to XML strings.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import { xsdMathStyle } from "@office-open/core";
import { attr, children, escapeXml, findChild, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { BodyContext, DocxReadContext } from "../../../context";
import { autoRevisionId } from "../../../shared/track-revision/track-revision";
import { stringifyDrawingChild } from "../../inline-drawings";
import type {
  MathArgumentPropertiesOptions,
  MathBorderBoxProperties,
  MathBoxProperties,
  MathControlPropertiesOptions,
  MathDelimiterProperties,
  MathEquationArrayProperties,
  MathGroupCharacterProperties,
  MathInput,
  MathMatrixColumnOptions,
  MathMatrixProperties,
  MathNaryProperties,
  MathPhantomProperties,
  MathRunPropertiesOptions,
} from "../math";
import type { RunPropertiesOptions } from "../run/properties";
import { EMPTY_RUN_ELEMENTS, breakXml } from "../run/run";
import type { RunOptions } from "../run/run";
import { parseRun, parseRunProperties, parsedRunToOptions } from "../run/run-parse";
import { stringifyRunProperties } from "../stringify";

// ── MathRun properties ──

function mathRunPropsStr(opts: MathRunPropertiesOptions): string {
  const parts: string[] = [];
  if (opts.literal !== undefined) parts.push(`<m:lit m:val="${onOff(opts.literal)}"/>`);
  if (opts.normal !== undefined) {
    parts.push(`<m:nor m:val="${onOff(opts.normal)}"/>`);
  } else {
    if (opts.script) parts.push(`<m:scr m:val="${opts.script}"/>`);
    if (opts.style) parts.push(`<m:sty m:val="${xsdMathStyle.to(opts.style)}"/>`);
  }
  if (opts.breakAlignment !== undefined) parts.push(`<m:brk m:alnAt="${opts.breakAlignment}"/>`);
  if (opts.align !== undefined) parts.push(`<m:aln m:val="${onOff(opts.align)}"/>`);
  return parts.length ? `<m:rPr>${parts.join("")}</m:rPr>` : "";
}

/** w:rPr serialization shared by the m:r run form and m:ctrlPr. */
function wRPrXml(opts: RunPropertiesOptions | undefined): string {
  return opts ? (stringifyRunProperties(opts) ?? "") : "";
}

/** m:ctrlPr — always the last child of a structure's *Pr element. */
function revisionAttrs(revision: { id?: number; author: string; date?: string }): string {
  const date = revision.date ? ` w:date="${escapeXml(revision.date)}"` : "";
  return ` w:id="${revision.id ?? autoRevisionId()}" w:author="${escapeXml(revision.author)}"${date}`;
}

function ctrlPrXml(ctrlPr: MathControlPropertiesOptions | undefined): string {
  if (!ctrlPr) return "";
  if ("insertion" in ctrlPr) {
    const { deletion, runProperties, ...revision } = ctrlPr.insertion;
    let inner = runProperties ? wRPrXml(runProperties) : "";
    if (deletion) {
      const { author: _author, date: _date, id: _id, ...original } = deletion;
      inner += `<w:del${revisionAttrs(deletion)}><w:rPr>${wRPrXml(original)}</w:rPr></w:del>`;
    }
    return `<m:ctrlPr><w:ins${revisionAttrs(revision)}>${inner}</w:ins></m:ctrlPr>`;
  }
  if ("deletion" in ctrlPr) {
    const { runProperties, ...revision } = ctrlPr.deletion;
    return `<m:ctrlPr><w:del${revisionAttrs(revision)}>${runProperties ? wRPrXml(runProperties) : ""}</w:del></m:ctrlPr>`;
  }
  const rPr = wRPrXml(ctrlPr.runProperties);
  return rPr ? `<m:ctrlPr>${rPr}</m:ctrlPr>` : "<m:ctrlPr/>";
}

/** Read the EG_RPrMath choice inside an m:ctrlPr. */
function readCtrlPr(prEl: Element | undefined): MathControlPropertiesOptions | undefined {
  const ctrlPr = prEl ? findChild(prEl, "m:ctrlPr") : undefined;
  return ctrlPr ? readCtrlPrElement(ctrlPr) : undefined;
}

function readCtrlPrElement(ctrlPr: Element): MathControlPropertiesOptions | undefined {
  const rPr = findChild(ctrlPr, "w:rPr");
  if (rPr) return { runProperties: parseRunProperties(rPr) };
  const ins = findChild(ctrlPr, "w:ins");
  if (ins) {
    const insRPr = findChild(ins, "w:rPr");
    const delEl = findChild(ins, "w:del");
    const delRPr = delEl ? findChild(delEl, "w:rPr") : undefined;
    return {
      insertion: {
        ...readRevision(ins),
        ...(insRPr ? { runProperties: parseRunProperties(insRPr) } : {}),
        ...(delEl
          ? { deletion: { ...readRevision(delEl), ...(delRPr ? parseRunProperties(delRPr) : {}) } }
          : {}),
      },
    };
  }
  const del = findChild(ctrlPr, "w:del");
  if (del) {
    const delRPr = findChild(del, "w:rPr");
    return {
      deletion: {
        ...readRevision(del),
        ...(delRPr ? { runProperties: parseRunProperties(delRPr) } : {}),
      },
    };
  }
  return {};
}

function readRevision(el: Element): { id?: number; author: string; date?: string } {
  const id = Number(attr(el, "w:id"));
  return {
    ...(Number.isFinite(id) ? { id } : {}),
    author: attr(el, "w:author") ?? "",
    ...(attr(el, "w:date") ? { date: attr(el, "w:date")! } : {}),
  };
}

// ── Children array ──

function mathTextXml(text: string, preserveSpace?: boolean): string {
  return `<m:t${preserveSpace ? ' xml:space="preserve"' : ""}>${escapeXml(text)}</m:t>`;
}

function stringifyWordRunChild(
  child: NonNullable<RunOptions["children"]>[number],
  ctx?: BodyContext,
): string {
  if (typeof child === "string") return mathTextXml(child);
  if ("text" in child && typeof child.text === "string") {
    const segment = child as { text: string; preserveSpace?: boolean };
    return mathTextXml(segment.text, segment.preserveSpace);
  }
  if ("tab" in child) return "<w:tab/>";
  if ("break" in child || "pageBreak" in child || "columnBreak" in child) {
    if ("break" in child) return breakXml(child.break);
    if ("pageBreak" in child) return '<w:br w:type="page"/>';
    return '<w:br w:type="column"/>';
  }
  let firstKey: string | undefined;
  for (const key in child) {
    firstKey = key;
    break;
  }
  const emptyXml = firstKey !== undefined ? EMPTY_RUN_ELEMENTS[firstKey] : undefined;
  if (emptyXml !== undefined) return emptyXml;
  if (
    ctx &&
    ("picture" in child ||
      "chart" in child ||
      "chartEx" in child ||
      "smartArt" in child ||
      "wpsShape" in child ||
      "wpgGroup" in child ||
      "unsupportedDrawing" in child ||
      "contentPart" in child)
  ) {
    return stringifyDrawingChild(child as never, ctx, { bareInRun: true }) ?? "";
  }
  return "";
}

let currentMathReadContext: DocxReadContext | undefined;
let currentMathWriteContext: BodyContext | undefined;

function stringifyChildren(items: MathInput[], ctx?: BodyContext): string {
  return items.map((child) => stringifyMathInput(child, ctx)).join("");
}

// ── Main recursive stringifier ──

export function stringifyMathInput(value: MathInput, ctx?: BodyContext): string {
  const previousContext = currentMathWriteContext;
  if (ctx !== undefined) currentMathWriteContext = ctx;
  try {
    return stringifyMathInputInner(value, ctx ?? currentMathWriteContext);
  } finally {
    currentMathWriteContext = previousContext;
  }
}

function stringifyMathInputInner(value: MathInput, ctx?: BodyContext): string {
  // String → MathRun shorthand
  if (typeof value === "string") {
    return `<m:r><m:t>${escapeXml(value)}</m:t></m:r>`;
  }

  // Class instances — still need toXml (shouldn't happen in compile/ JSON path)
  if (typeof value !== "object" || value === null) return "";

  // Trailing ctrlPr of an argument element (m:e/m:num/m:den...) — emitted in
  // place so its position among the children round-trips.
  if ("argumentControlProperties" in value) {
    return ctrlPrXml(value.argumentControlProperties);
  }

  // Discriminated union: check unique keys (order matters — subSuperScript first)
  if ("subSuperScript" in value) {
    const opts = value.subSuperScript;
    const ctrl = ctrlPrXml(opts.controlProperties);
    const pr =
      opts.alignScript !== undefined || ctrl
        ? `<m:sSubSupPr>${opts.alignScript !== undefined ? `<m:alnScr m:val="${onOff(opts.alignScript)}"/>` : ""}${ctrl}</m:sSubSupPr>`
        : "";
    return `<m:sSubSup>${pr}<m:e>${argPrXml(opts.baseArgumentProperties)}${stringifyChildren(opts.children)}</m:e><m:sub>${argPrXml(opts.subScriptArgumentProperties)}${stringifyChildren(opts.subScript)}</m:sub><m:sup>${argPrXml(opts.superScriptArgumentProperties)}${stringifyChildren(opts.superScript)}</m:sup></m:sSubSup>`;
  }

  if ("preSubSuperScript" in value) {
    const opts = value.preSubSuperScript;
    const ctrl = ctrlPrXml(opts.controlProperties);
    const pr = ctrl ? `<m:sPrePr>${ctrl}</m:sPrePr>` : "";
    return `<m:sPre>${pr}<m:sub>${argPrXml(opts.subScriptArgumentProperties)}${stringifyChildren(opts.subScript)}</m:sub><m:sup>${argPrXml(opts.superScriptArgumentProperties)}${stringifyChildren(opts.superScript)}</m:sup><m:e>${argPrXml(opts.baseArgumentProperties)}${stringifyChildren(opts.children)}</m:e></m:sPre>`;
  }

  if ("superScript" in value) {
    const opts = value.superScript;
    const ctrl = ctrlPrXml(opts.controlProperties);
    const pr = ctrl ? `<m:sSupPr>${ctrl}</m:sSupPr>` : "";
    return `<m:sSup>${pr}<m:e>${argPrXml(opts.baseArgumentProperties)}${stringifyChildren(opts.children)}</m:e><m:sup>${argPrXml(opts.superScriptArgumentProperties)}${stringifyChildren(opts.superScript)}</m:sup></m:sSup>`;
  }

  if ("subScript" in value) {
    const opts = value.subScript;
    const ctrl = ctrlPrXml(opts.controlProperties);
    const pr = ctrl ? `<m:sSubPr>${ctrl}</m:sSubPr>` : "";
    return `<m:sSub>${pr}<m:e>${argPrXml(opts.baseArgumentProperties)}${stringifyChildren(opts.children)}</m:e><m:sub>${argPrXml(opts.subScriptArgumentProperties)}${stringifyChildren(opts.subScript)}</m:sub></m:sSub>`;
  }

  if ("fraction" in value) {
    const opts = value.fraction;
    const inner =
      (opts.fractionType ? `<m:type m:val="${opts.fractionType}"/>` : "") +
      ctrlPrXml(opts.controlProperties);
    const pr = inner ? `<m:fPr>${inner}</m:fPr>` : "";
    const numArgPr = argPrXml(opts.numeratorArgumentProperties);
    const denArgPr = argPrXml(opts.denominatorArgumentProperties);
    return `<m:f>${pr}<m:num>${numArgPr}${stringifyChildren(opts.numerator)}</m:num><m:den>${denArgPr}${stringifyChildren(opts.denominator)}</m:den></m:f>`;
  }

  if ("radical" in value) {
    const opts = value.radical;
    const hasDegree = opts.degree && opts.degree.length > 0;
    const explicitHideDegree = opts.properties?.hideDegree;
    const hideDegree = explicitHideDegree ?? !hasDegree;
    const inner =
      (explicitHideDegree !== undefined || !hasDegree
        ? `<m:degHide m:val="${onOff(hideDegree)}"/>`
        : "") + ctrlPrXml(opts.controlProperties);
    const pr = inner ? `<m:radPr>${inner}</m:radPr>` : "";
    const deg = hasDegree ? `<m:deg>${stringifyChildren(opts.degree!)}</m:deg>` : "<m:deg/>";
    return `<m:rad>${pr}${deg}<m:e>${stringifyChildren(opts.children)}</m:e></m:rad>`;
  }

  if ("sum" in value) {
    return stringifyNAry(value.sum, "∑");
  }

  if ("nary" in value) {
    return stringifyNAry(value.nary, value.nary.character);
  }

  if ("integral" in value) {
    return stringifyNAry(value.integral, "∫");
  }

  if ("limitLower" in value) {
    const opts = value.limitLower;
    const pr = opts.controlProperties
      ? `<m:limLowPr>${ctrlPrXml(opts.controlProperties)}</m:limLowPr>`
      : "";
    return `<m:limLow>${pr}<m:e>${stringifyChildren(opts.children)}</m:e><m:lim>${stringifyChildren(opts.limit)}</m:lim></m:limLow>`;
  }

  if ("limitUpper" in value) {
    const opts = value.limitUpper;
    const pr = opts.controlProperties
      ? `<m:limUppPr>${ctrlPrXml(opts.controlProperties)}</m:limUppPr>`
      : "";
    return `<m:limUpp>${pr}<m:e>${stringifyChildren(opts.children)}</m:e><m:lim>${stringifyChildren(opts.limit)}</m:lim></m:limUpp>`;
  }

  if ("function" in value) {
    const opts = value.function;
    const pr = opts.controlProperties
      ? `<m:funcPr>${ctrlPrXml(opts.controlProperties)}</m:funcPr>`
      : "";
    return `<m:func>${pr}<m:fName>${stringifyChildren(opts.name)}</m:fName><m:e>${stringifyChildren(opts.children)}</m:e></m:func>`;
  }

  if ("matrix" in value) {
    const opts = value.matrix;
    const rows = opts.rows
      .map(
        (row) =>
          `<m:mr>${row
            .map(
              (cell) =>
                `<m:e>${
                  Array.isArray(cell)
                    ? cell.map((item) => stringifyMathInput(item)).join("")
                    : stringifyMathInput(cell)
                }</m:e>`,
            )
            .join("")}</m:mr>`,
      )
      .join("");
    const p = opts.properties;
    const prParts: string[] = [];
    if (p?.baseline) prParts.push(`<m:baseJc m:val="${p.baseline}"/>`);
    if (p?.hidePlaceholder !== undefined)
      prParts.push(`<m:plcHide m:val="${onOff(p.hidePlaceholder)}"/>`);
    if (p?.rowSpacingRule) prParts.push(`<m:rSpRule m:val="${p.rowSpacingRule}"/>`);
    if (p?.columnGapRule) prParts.push(`<m:cGpRule m:val="${p.columnGapRule}"/>`);
    if (p?.rowSpacing !== undefined) prParts.push(`<m:rSp m:val="${p.rowSpacing}"/>`);
    if (p?.columnSpacing !== undefined) prParts.push(`<m:cSp m:val="${p.columnSpacing}"/>`);
    if (p?.columnGap !== undefined) prParts.push(`<m:cGp m:val="${p.columnGap}"/>`);
    if (p?.columns?.length) {
      const mcItems = p.columns
        .map(
          (mc) =>
            `<m:mc><m:mcPr><m:count m:val="${mc.count}"/><m:mcJc m:val="${mc.justification}"/></m:mcPr></m:mc>`,
        )
        .join("");
      prParts.push(`<m:mcs>${mcItems}</m:mcs>`);
    }
    prParts.push(ctrlPrXml(opts.controlProperties));
    const pr = prParts.some(Boolean) ? `<m:mPr>${prParts.join("")}</m:mPr>` : "";
    return `<m:m>${pr}${rows}</m:m>`;
  }

  // Bracket types
  if ("roundBrackets" in value) {
    const spec = bracketSpec(value.roundBrackets);
    return stringifyDelimiters(spec);
  }
  if ("curlyBrackets" in value) {
    const spec = bracketSpec(value.curlyBrackets);
    return stringifyDelimiters(spec);
  }
  if ("angledBrackets" in value) {
    const spec = bracketSpec(value.angledBrackets);
    return stringifyDelimiters(spec);
  }
  if ("squareBrackets" in value) {
    const spec = bracketSpec(value.squareBrackets);
    return stringifyDelimiters(spec);
  }

  if ("borderBox" in value) {
    const opts = value.borderBox;
    const p = opts.properties;
    const parts: string[] = [];
    if (p?.hideTop !== undefined) parts.push(`<m:hideTop m:val="${onOff(p.hideTop)}"/>`);
    if (p?.hideBottom !== undefined) parts.push(`<m:hideBot m:val="${onOff(p.hideBottom)}"/>`);
    if (p?.hideLeft !== undefined) parts.push(`<m:hideLeft m:val="${onOff(p.hideLeft)}"/>`);
    if (p?.hideRight !== undefined) parts.push(`<m:hideRight m:val="${onOff(p.hideRight)}"/>`);
    if (p?.strikeHorizontal !== undefined)
      parts.push(`<m:strikeH m:val="${onOff(p.strikeHorizontal)}"/>`);
    if (p?.strikeVertical !== undefined)
      parts.push(`<m:strikeV m:val="${onOff(p.strikeVertical)}"/>`);
    if (p?.strikeDiagonalUp !== undefined)
      parts.push(`<m:strikeBLTR m:val="${onOff(p.strikeDiagonalUp)}"/>`);
    if (p?.strikeDiagonalDown !== undefined)
      parts.push(`<m:strikeTLBR m:val="${onOff(p.strikeDiagonalDown)}"/>`);
    parts.push(ctrlPrXml(opts.controlProperties));
    const pr = parts.some(Boolean) ? `<m:borderBoxPr>${parts.join("")}</m:borderBoxPr>` : "";
    return `<m:borderBox>${pr}<m:e>${stringifyChildren(opts.children)}</m:e></m:borderBox>`;
  }

  if ("box" in value) {
    const opts = value.box;
    const p = opts.properties;
    const parts: string[] = [];
    if (p?.operatorEmulation !== undefined)
      parts.push(`<m:opEmu m:val="${onOff(p.operatorEmulation)}"/>`);
    if (p?.noBreak !== undefined) parts.push(`<m:noBreak m:val="${onOff(p.noBreak)}"/>`);
    if (p?.differential !== undefined) parts.push(`<m:diff m:val="${onOff(p.differential)}"/>`);
    if (p?.align !== undefined) parts.push(`<m:aln m:val="${onOff(p.align)}"/>`);
    parts.push(ctrlPrXml(opts.controlProperties));
    const pr = parts.some(Boolean) ? `<m:boxPr>${parts.join("")}</m:boxPr>` : "";
    return `<m:box>${pr}<m:e>${stringifyChildren(opts.children)}</m:e></m:box>`;
  }

  if ("groupChr" in value) {
    const opts = value.groupChr;
    const p = opts.properties;
    const parts: string[] = [];
    if (p?.character) parts.push(`<m:chr m:val="${escapeXml(p.character)}"/>`);
    if (p?.position) parts.push(`<m:pos m:val="${p.position}"/>`);
    if (p?.vertical) parts.push(`<m:vertJc m:val="${p.vertical}"/>`);
    parts.push(ctrlPrXml(opts.controlProperties));
    const pr = parts.some(Boolean) ? `<m:groupChrPr>${parts.join("")}</m:groupChrPr>` : "";
    return `<m:groupChr>${pr}<m:e>${stringifyChildren(opts.children)}</m:e></m:groupChr>`;
  }

  if ("phant" in value) {
    const opts = value.phant;
    const p = opts.properties;
    const parts: string[] = [];
    if (p?.show !== undefined) parts.push(`<m:show m:val="${onOff(p.show)}"/>`);
    if (p?.zeroWidth !== undefined) parts.push(`<m:zeroWid m:val="${onOff(p.zeroWidth)}"/>`);
    if (p?.zeroAscent !== undefined) parts.push(`<m:zeroAsc m:val="${onOff(p.zeroAscent)}"/>`);
    if (p?.zeroDescent !== undefined) parts.push(`<m:zeroDesc m:val="${onOff(p.zeroDescent)}"/>`);
    if (p?.transparent !== undefined) parts.push(`<m:transp m:val="${onOff(p.transparent)}"/>`);
    parts.push(ctrlPrXml(opts.controlProperties));
    const pr = parts.some(Boolean) ? `<m:phantPr>${parts.join("")}</m:phantPr>` : "";
    return `<m:phant>${pr}<m:e>${stringifyChildren(opts.children)}</m:e></m:phant>`;
  }

  if ("eqArr" in value) {
    const opts = value.eqArr;
    const p = opts.properties;
    const parts: string[] = [];
    if (p?.baseline) parts.push(`<m:baseJc m:val="${p.baseline}"/>`);
    if (p?.distributeRows !== undefined)
      parts.push(`<m:maxDist m:val="${onOff(p.distributeRows)}"/>`);
    if (p?.objectDistance !== undefined)
      parts.push(`<m:objDist m:val="${onOff(p.objectDistance)}"/>`);
    if (p?.rowSpacingRule) parts.push(`<m:rSpRule m:val="${p.rowSpacingRule}"/>`);
    if (p?.rowSpacing !== undefined) parts.push(`<m:rSp m:val="${p.rowSpacing}"/>`);
    parts.push(ctrlPrXml(opts.controlProperties));
    const pr = parts.some(Boolean) ? `<m:eqArrPr>${parts.join("")}</m:eqArrPr>` : "";
    const rows = opts.rows.map((row) => `<m:e>${stringifyChildren(row)}</m:e>`).join("");
    return `<m:eqArr>${pr}${rows}</m:eqArr>`;
  }

  if ("accent" in value) {
    const opts = value.accent;
    const inner =
      (opts.accentCharacter ? `<m:chr m:val="${escapeXml(opts.accentCharacter)}"/>` : "") +
      ctrlPrXml(opts.controlProperties);
    const pr = inner ? `<m:accPr>${inner}</m:accPr>` : "";
    return `<m:acc>${pr}<m:e>${stringifyChildren(opts.children)}</m:e></m:acc>`;
  }

  if ("bar" in value) {
    const opts = value.bar;
    const inner = `<m:pos m:val="${opts.type}"/>` + ctrlPrXml(opts.controlProperties);
    return `<m:bar><m:barPr>${inner}</m:barPr><m:e>${stringifyChildren(opts.children)}</m:e></m:bar>`;
  }

  // Fallback: { text, properties?, runProperties? } → MathRun
  if ("text" in value) {
    const props = value.properties ? mathRunPropsStr(value.properties) : "";
    const rPr = wRPrXml(value.runProperties);
    const space = value.preserveSpace ? ` xml:space="preserve"` : "";
    return `<m:r>${props}${rPr}<m:t${space}>${escapeXml(value.text)}</m:t></m:r>`;
  }

  if ("children" in value) {
    const props = value.properties ? mathRunPropsStr(value.properties) : "";
    const rPr = wRPrXml(value.runProperties);
    return `<m:r>${props}${rPr}${value.children.map((child) => stringifyWordRunChild(child, ctx)).join("")}</m:r>`;
  }

  return "";
}

// ── N-ary operator (sum/integral/custom) ──

function stringifyNAry(
  opts: {
    children: MathInput[];
    subScript?: MathInput[];
    superScript?: MathInput[];
    properties?: MathNaryProperties;
    controlProperties?: MathControlPropertiesOptions;
  },
  character: string,
): string {
  const hasSub = opts.subScript && opts.subScript.length > 0;
  const hasSup = opts.superScript && opts.superScript.length > 0;
  const prParts: string[] = [`<m:chr m:val="${character}"/>`];
  if (opts.properties?.limitLocation)
    prParts.push(`<m:limLoc m:val="${opts.properties.limitLocation}"/>`);
  if (opts.properties?.grow !== undefined)
    prParts.push(`<m:grow m:val="${onOff(opts.properties.grow)}"/>`);
  if (opts.properties?.hideSubscript !== undefined || !hasSub)
    prParts.push(`<m:subHide m:val="${onOff(opts.properties?.hideSubscript ?? true)}"/>`);
  if (opts.properties?.hideSuperscript !== undefined || !hasSup)
    prParts.push(`<m:supHide m:val="${onOff(opts.properties?.hideSuperscript ?? true)}"/>`);
  prParts.push(ctrlPrXml(opts.controlProperties));
  const pr = `<m:naryPr>${prParts.join("")}</m:naryPr>`;
  const sub = hasSub ? `<m:sub>${stringifyChildren(opts.subScript!)}</m:sub>` : "<m:sub/>";
  const sup = hasSup ? `<m:sup>${stringifyChildren(opts.superScript!)}</m:sup>` : "<m:sup/>";
  return `<m:nary>${pr}${sub}${sup}<m:e>${stringifyChildren(opts.children)}</m:e></m:nary>`;
}

// ── Delimiters (brackets) ──

function stringifyDelimiters(spec: {
  children?: MathInput[];
  elements?: MathInput[][];
  properties?: MathDelimiterProperties;
}): string {
  // CT_D holds e+ — one m:e per separator-split group; a bare children array
  // is a single group.
  const groups = spec.elements ?? (spec.children ? [spec.children] : []);
  const eXml = groups.map((g) => `<m:e>${stringifyChildren(g)}</m:e>`).join("");
  // begChr/endChr are optional with XSD defaults "(" / ")" — omit when default
  // (Office writes bare <m:dPr> for round brackets).
  const beg = spec.properties?.beginCharacter;
  const end = spec.properties?.endCharacter;
  const prParts: string[] = [];
  if (beg !== undefined) prParts.push(`<m:begChr m:val="${escapeXml(beg)}"/>`);
  if (spec.properties?.separatorCharacter !== undefined)
    prParts.push(`<m:sepChr m:val="${escapeXml(spec.properties.separatorCharacter)}"/>`);
  if (end !== undefined) prParts.push(`<m:endChr m:val="${escapeXml(end)}"/>`);
  if (spec.properties?.grow !== undefined)
    prParts.push(`<m:grow m:val="${onOff(spec.properties.grow)}"/>`);
  if (spec.properties?.shape) prParts.push(`<m:shp m:val="${spec.properties.shape}"/>`);
  prParts.push(ctrlPrXml(spec.properties?.controlProperties));
  return `<m:d><m:dPr>${prParts.join("")}</m:dPr>${eXml}</m:d>`;
}

/** Build an m:argPr/m:argSz block for an argument size scaling value. */
function argPrXml(properties: MathArgumentPropertiesOptions | undefined): string {
  if (properties === undefined) return "";
  if (properties.size === undefined) return "<m:argPr/>";
  return `<m:argPr><m:argSz m:val="${properties.size}"/></m:argPr>`;
}

/** Split a bracket shorthand into children/elements + delimiter properties. */
function bracketSpec(
  v:
    | MathInput[]
    | {
        children?: MathInput[];
        elements?: MathInput[][];
        properties?: MathDelimiterProperties;
      },
): {
  children?: MathInput[];
  elements?: MathInput[][];
  properties?: MathDelimiterProperties;
} {
  if (Array.isArray(v)) return { children: v };
  return v;
}

// ── Top-level Math wrapper ──

export function stringifyMath(children: MathInput[], ctx?: BodyContext): string {
  const inner = children.map((c) => stringifyMathInput(c, ctx)).join("");
  return `<m:oMath>${inner}</m:oMath>`;
}

export function stringifyMathParagraph(
  equations: MathInput[][],
  justification?: "left" | "right" | "center" | "centerGroup",
  ctx?: BodyContext,
): string {
  const inner = equations
    .map((children) => `<m:oMath>${stringifyChildren(children, ctx)}</m:oMath>`)
    .join("");
  const pr = justification ? `<m:oMathParaPr><m:jc m:val="${justification}"/></m:oMathParaPr>` : "";
  return `<m:oMathPara>${pr}${inner}</m:oMathPara>`;
}

// ────────────────────────────────────────────────────────────────────────────────
// Parse (OMML XML → MathInput)
// ────────────────────────────────────────────────────────────────────────────────

/**
 * Parse all math children from an m:oMath (or similar container) element.
 */
export function parseMathChildren(el: Element, ctx?: DocxReadContext): MathInput[] {
  const previousContext = currentMathReadContext;
  if (ctx !== undefined) currentMathReadContext = ctx;
  try {
    const result: MathInput[] = [];
    for (const child of el.elements ?? []) {
      const parsed = parseMathElement(child);
      if (parsed !== undefined) result.push(parsed);
    }
    return result;
  } finally {
    currentMathReadContext = previousContext;
  }
}

/** Parse an m:oMathPara into its equation list and justification. */
export function parseMathParagraph(
  el: Element,
  ctx?: DocxReadContext,
): {
  equations: MathInput[][];
  justification?: "left" | "right" | "center" | "centerGroup";
} {
  const jcEl = findChild(findChild(el, "m:oMathParaPr") ?? el, "m:jc");
  const jc = jcEl ? attr(jcEl, "m:val") : undefined;
  return {
    equations: children(el, "m:oMath").map((equation) => parseMathChildren(equation, ctx)),
    ...(jc === "left" || jc === "right" || jc === "center" || jc === "centerGroup"
      ? { justification: jc }
      : {}),
  };
}

function parseMathElement(el: Element): MathInput | undefined {
  switch (el.name) {
    case "m:r":
      return parseMathRun(el);
    case "m:f":
      return parseMathFraction(el);
    case "m:rad":
      return parseMathRadical(el);
    case "m:sSup":
      return parseMathSuperScript(el);
    case "m:sSub":
      return parseMathSubScript(el);
    case "m:sSubSup":
      return parseMathSubSuperScript(el);
    case "m:sPre":
      return parseMathPreSubSuperScript(el);
    case "m:nary":
      return parseMathNAry(el);
    case "m:func":
      return parseMathFunction(el);
    case "m:d":
      return parseMathDelimiter(el);
    case "m:m":
      return parseMathMatrix(el);
    case "m:acc":
      return parseMathAccent(el);
    case "m:bar":
      return parseMathBar(el);
    case "m:borderBox":
      return parseMathBorderBox(el);
    case "m:box":
      return parseMathBox(el);
    case "m:groupChr":
      return parseMathGroupChr(el);
    case "m:phant":
      return parseMathPhant(el);
    case "m:eqArr":
      return parseMathEqArr(el);
    case "m:limLow":
      return parseMathLimitLower(el);
    case "m:limUpp":
      return parseMathLimitUpper(el);
    // Property elements — not standalone content
    case "m:rPr":
    case "m:fPr":
    case "m:radPr":
    case "m:sSupPr":
    case "m:sSubPr":
    case "m:sSubSupPr":
    case "m:sPrePr":
    case "m:naryPr":
    case "m:funcPr":
    case "m:dPr":
    case "m:mPr":
    case "m:accPr":
    case "m:barPr":
    case "m:borderBoxPr":
    case "m:boxPr":
    case "m:groupChrPr":
    case "m:phantPr":
    case "m:eqArrPr":
    case "m:limLowPr":
    case "m:limUppPr":
      return undefined;
    case "m:ctrlPr": {
      // Content-stream ctrlPr: the trailing control formatting of an argument
      // element (m:e/m:num/m:den...) — keep it as a children member in place.
      return { argumentControlProperties: readCtrlPrElement(el) ?? {} };
    }
    default:
      return undefined;
  }
}

function parseMathRun(el: Element): MathInput {
  const textEl = findChild(el, "m:t");
  const text = textOf(textEl);
  const preserveSpace = attr(textEl, "xml:space") === "preserve";
  const rPrEl = findChild(el, "m:rPr");
  const wRPrEl = findChild(el, "w:rPr");
  const contentElements = (el.elements ?? []).filter(
    (child) => child.type === "element" && child.name !== "m:rPr" && child.name !== "w:rPr",
  );
  const hasWordRunContent = contentElements.some((child) => child.name !== "m:t");
  const wordElements = contentElements.map((child) =>
    child.type === "element" && child.name === "m:t" ? { ...child, name: "w:t" } : child,
  );
  const parsedRunOptions = hasWordRunContent
    ? parsedRunToOptions(
        parseRun({ ...el, elements: wordElements }, currentMathReadContext as DocxReadContext),
      )
    : undefined;
  const parsedChildren =
    parsedRunOptions !== undefined && "children" in parsedRunOptions
      ? parsedRunOptions.children
      : undefined;
  const hasMixedWordContent =
    parsedChildren !== undefined ||
    (parsedRunOptions !== undefined && !textEl && "text" in parsedRunOptions);
  if (hasMixedWordContent && parsedChildren !== undefined) {
    const properties = rPrEl ? readMathRunProperties(rPrEl) : undefined;
    const runProperties = wRPrEl ? parseRunProperties(wRPrEl) : undefined;
    return {
      children: parsedChildren,
      ...(properties && Object.keys(properties).length > 0 ? { properties } : {}),
      ...(runProperties && Object.keys(runProperties).length > 0 ? { runProperties } : {}),
    };
  }
  if (!rPrEl && !wRPrEl) {
    return preserveSpace ? { text: text ?? "", preserveSpace: true } : (text ?? "");
  }
  const run: {
    text: string;
    properties?: MathRunPropertiesOptions;
    runProperties?: RunPropertiesOptions;
  } = {
    text: text ?? "",
    ...(preserveSpace ? { preserveSpace: true } : {}),
  };
  if (rPrEl) {
    const props = readMathRunProperties(rPrEl);
    if (Object.keys(props).length > 0) run.properties = props;
  }
  if (wRPrEl) {
    const rPr = parseRunProperties(wRPrEl);
    if (Object.keys(rPr).length > 0) run.runProperties = rPr;
  }
  return run;
}

/** Read m:rPr (CT_RPr) into its structured options. */
function readMathRunProperties(el: Element): MathRunPropertiesOptions {
  const result: MathRunPropertiesOptions = {};
  const literal = readOnOff(findChild(el, "m:lit"));
  if (literal !== undefined) result.literal = literal;
  const normal = readOnOff(findChild(el, "m:nor"));
  if (normal !== undefined) result.normal = normal;
  const scr = attr(findChild(el, "m:scr"), "m:val");
  if (scr) result.script = scr as MathRunPropertiesOptions["script"];
  const sty = attr(findChild(el, "m:sty"), "m:val");
  if (sty) result.style = xsdMathStyle.from(sty) as MathRunPropertiesOptions["style"];
  const brk = attr(findChild(el, "m:brk"), "m:alnAt");
  if (brk !== undefined && brk !== "") result.breakAlignment = Number(brk);
  const align = readOnOff(findChild(el, "m:aln"));
  if (align !== undefined) result.align = align;
  return result;
}

// ── Parse helpers ──

/** Read an m:val on/off attribute (1/0/true/false; empty element = on). */
function readOnOff(el: Element | undefined): boolean | undefined {
  if (!el) return undefined;
  const v = attr(el, "m:val");
  return parseOnOff(v) ?? true;
}

/** ST_OnOff literal — Word writes on/off in the math domain. */
function onOff(v: boolean): string {
  return v ? "on" : "off";
}

/** s:ST_YAlign values (m:baseJc). */
const Y_ALIGN = new Set(["inline", "top", "center", "bottom", "inside", "outside"]);

/** Read an m:val numeric attribute. */
function readNum(el: Element | undefined): number | undefined {
  if (!el) return undefined;
  const v = attr(el, "m:val");
  if (v === undefined || v === "") return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

/** Read m:argPr from an argument element (CT_OMathArg). */
function readArgumentProperties(
  argEl: Element | undefined,
): MathArgumentPropertiesOptions | undefined {
  if (!argEl) return undefined;
  const argPr = findChild(argEl, "m:argPr");
  if (!argPr) return undefined;
  const size = readNum(findChild(argPr, "m:argSz"));
  return size !== undefined ? { size } : {};
}

function parseMathFraction(el: Element): MathInput {
  const fPr = findChild(el, "m:fPr");
  const fractionType = attr(findChild(fPr, "m:type"), "m:val");
  const numeratorArgumentProperties = readArgumentProperties(findChild(el, "m:num"));
  const denominatorArgumentProperties = readArgumentProperties(findChild(el, "m:den"));
  return {
    fraction: {
      numerator: parseMathArg(el, "m:num"),
      denominator: parseMathArg(el, "m:den"),
      ...(fractionType ? { fractionType } : {}),
      ...(numeratorArgumentProperties !== undefined ? { numeratorArgumentProperties } : {}),
      ...(denominatorArgumentProperties !== undefined ? { denominatorArgumentProperties } : {}),
      ...spreadCtrlPr(fPr),
    },
  };
}

function parseMathRadical(el: Element): MathInput {
  const degree = parseMathArg(el, "m:deg");
  const mathChildren = parseMathArg(el, "m:e");
  const radPr = findChild(el, "m:radPr");
  const hideDegree = readOnOff(findChild(radPr, "m:degHide"));
  return {
    radical: {
      children: mathChildren,
      ...(degree.length > 0 ? { degree } : {}),
      ...(hideDegree !== undefined ? { properties: { hideDegree } } : {}),
      ...spreadCtrlPr(radPr),
    },
  };
}

/** Spread a ctrlPr read into its "ctrlPr" option field. */
function spreadCtrlPr(prEl: Element | undefined): {
  controlProperties?: MathControlPropertiesOptions;
} {
  const ctrlPr = readCtrlPr(prEl);
  return ctrlPr !== undefined ? { controlProperties: ctrlPr } : {};
}

function parseMathSuperScript(el: Element): MathInput {
  const baseArgumentProperties = readArgumentProperties(findChild(el, "m:e"));
  const superScriptArgumentProperties = readArgumentProperties(findChild(el, "m:sup"));
  return {
    superScript: {
      children: parseMathArg(el, "m:e"),
      superScript: parseMathArg(el, "m:sup"),
      ...(baseArgumentProperties !== undefined ? { baseArgumentProperties } : {}),
      ...(superScriptArgumentProperties !== undefined ? { superScriptArgumentProperties } : {}),
      ...spreadCtrlPr(findChild(el, "m:sSupPr")),
    },
  };
}

function parseMathSubScript(el: Element): MathInput {
  const baseArgumentProperties = readArgumentProperties(findChild(el, "m:e"));
  const subScriptArgumentProperties = readArgumentProperties(findChild(el, "m:sub"));
  return {
    subScript: {
      children: parseMathArg(el, "m:e"),
      subScript: parseMathArg(el, "m:sub"),
      ...(baseArgumentProperties !== undefined ? { baseArgumentProperties } : {}),
      ...(subScriptArgumentProperties !== undefined ? { subScriptArgumentProperties } : {}),
      ...spreadCtrlPr(findChild(el, "m:sSubPr")),
    },
  };
}

function parseMathSubSuperScript(el: Element): MathInput {
  const pr = findChild(el, "m:sSubSupPr");
  const alignScript = pr ? readOnOff(findChild(pr, "m:alnScr")) : undefined;
  const baseArgumentProperties = readArgumentProperties(findChild(el, "m:e"));
  const subScriptArgumentProperties = readArgumentProperties(findChild(el, "m:sub"));
  const superScriptArgumentProperties = readArgumentProperties(findChild(el, "m:sup"));
  return {
    subSuperScript: {
      children: parseMathArg(el, "m:e"),
      subScript: parseMathArg(el, "m:sub"),
      superScript: parseMathArg(el, "m:sup"),
      ...(baseArgumentProperties !== undefined ? { baseArgumentProperties } : {}),
      ...(subScriptArgumentProperties !== undefined ? { subScriptArgumentProperties } : {}),
      ...(superScriptArgumentProperties !== undefined ? { superScriptArgumentProperties } : {}),
      ...(alignScript !== undefined ? { alignScript } : {}),
      ...spreadCtrlPr(pr),
    },
  };
}

function parseMathNAry(el: Element): MathInput {
  const naryPr = findChild(el, "m:naryPr");
  const chrEl = naryPr ? findChild(naryPr, "m:chr") : undefined;
  const chrVal = chrEl ? attr(chrEl, "m:val") : undefined;

  const baseChildren = parseMathArg(el, "m:e");
  const sub = parseMathArg(el, "m:sub");
  const sup = parseMathArg(el, "m:sup");

  const properties: MathNaryProperties = {};
  if (naryPr) {
    const limLocEl = findChild(naryPr, "m:limLoc");
    if (limLocEl) {
      const limLoc = attr(limLocEl, "m:val");
      if (limLoc === "subSup" || limLoc === "undOvr") properties.limitLocation = limLoc;
    }
    const grow = readOnOff(findChild(naryPr, "m:grow"));
    if (grow !== undefined) properties.grow = grow;
    const hideSubscript = readOnOff(findChild(naryPr, "m:subHide"));
    if (hideSubscript !== undefined) properties.hideSubscript = hideSubscript;
    const hideSuperscript = readOnOff(findChild(naryPr, "m:supHide"));
    if (hideSuperscript !== undefined) properties.hideSuperscript = hideSuperscript;
  }

  const common = {
    children: baseChildren,
    ...(sub.length > 0 ? { subScript: sub } : {}),
    ...(sup.length > 0 ? { superScript: sup } : {}),
    ...(Object.keys(properties).length > 0 ? { properties } : {}),
    ...spreadCtrlPr(naryPr),
  };

  if (chrVal === "∑") return { sum: common };
  if (chrVal === undefined || chrVal === "∫") return { integral: common };
  return { nary: { ...common, character: chrVal } };
}

function parseMathFunction(el: Element): MathInput {
  return {
    function: {
      name: parseMathArg(el, "m:fName"),
      children: parseMathArg(el, "m:e"),
      ...spreadCtrlPr(findChild(el, "m:funcPr")),
    },
  };
}

function parseMathDelimiter(el: Element): MathInput {
  const dPr = findChild(el, "m:dPr");
  const begChrEl = dPr ? findChild(dPr, "m:begChr") : undefined;
  const begChr = begChrEl ? attr(begChrEl, "m:val") : "(";
  // CT_D holds e+ — keep each m:e as its own separator-split group.
  const groups = children(el, "m:e").map((e) => parseMathChildren(e));

  // Collect delimiter properties when present (sepChr/grow/shp/non-default chars).
  const properties: MathDelimiterProperties = {};
  if (dPr) {
    if (begChrEl) properties.beginCharacter = begChr;
    const endChrEl = findChild(dPr, "m:endChr");
    if (endChrEl) properties.endCharacter = attr(endChrEl, "m:val");
    const sepChrEl = findChild(dPr, "m:sepChr");
    if (sepChrEl) properties.separatorCharacter = attr(sepChrEl, "m:val");
    const grow = readOnOff(findChild(dPr, "m:grow"));
    if (grow !== undefined) properties.grow = grow;
    const shpEl = findChild(dPr, "m:shp");
    if (shpEl) {
      const shp = attr(shpEl, "m:val");
      if (shp === "centered" || shp === "match") properties.shape = shp;
    }
  }
  const ctrlPr = readCtrlPr(dPr);
  if (ctrlPr !== undefined) properties.controlProperties = ctrlPr;
  const hasProperties = Object.keys(properties).length > 0;
  const value =
    groups.length > 1
      ? { elements: groups, ...(hasProperties ? { properties } : {}) }
      : hasProperties
        ? { children: groups[0] ?? [], properties }
        : (groups[0] ?? []);

  switch (begChr) {
    case "[":
      return { squareBrackets: value };
    case "{":
      return { curlyBrackets: value };
    case "<":
    case "⟨":
      return { angledBrackets: value };
    default:
      return { roundBrackets: value };
  }
}

function parseMathMatrix(el: Element): MathInput {
  const rows: (MathInput | MathInput[])[][] = [];
  for (const mr of children(el, "m:mr")) {
    // CT_MR holds e+ — one cell per m:e.
    rows.push(children(mr, "m:e").map((e) => parseMathChildren(e)));
  }
  const mPr = findChild(el, "m:mPr");
  const props = readMatrixProperties(mPr);
  return {
    matrix: {
      rows,
      ...(props ? { properties: props } : {}),
      ...spreadCtrlPr(mPr),
    },
  };
}

/** Read m:mPr (CT_MPr) into structured matrix properties. */
function readMatrixProperties(mPr: Element | undefined): MathMatrixProperties | undefined {
  if (!mPr) return undefined;
  const result: MathMatrixProperties = {};
  const baseline = attr(findChild(mPr, "m:baseJc"), "m:val");
  if (baseline !== undefined && Y_ALIGN.has(baseline))
    result.baseline = baseline as MathMatrixProperties["baseline"];
  const hidePlaceholder = readOnOff(findChild(mPr, "m:plcHide"));
  if (hidePlaceholder !== undefined) result.hidePlaceholder = hidePlaceholder;
  const rowSpacingRule = readNum(findChild(mPr, "m:rSpRule"));
  if (rowSpacingRule !== undefined)
    result.rowSpacingRule = rowSpacingRule as MathMatrixProperties["rowSpacingRule"];
  const columnGapRule = readNum(findChild(mPr, "m:cGpRule"));
  if (columnGapRule !== undefined)
    result.columnGapRule = columnGapRule as MathMatrixProperties["columnGapRule"];
  const rowSpacing = readNum(findChild(mPr, "m:rSp"));
  if (rowSpacing !== undefined) result.rowSpacing = rowSpacing;
  const columnSpacing = readNum(findChild(mPr, "m:cSp"));
  if (columnSpacing !== undefined) result.columnSpacing = columnSpacing;
  const columnGap = readNum(findChild(mPr, "m:cGp"));
  if (columnGap !== undefined) result.columnGap = columnGap;
  const columns = findChild(mPr, "m:mcs");
  if (columns) {
    const cols: MathMatrixColumnOptions[] = [];
    for (const mc of children(columns, "m:mc")) {
      const mcPr = findChild(mc, "m:mcPr");
      const count = readNum(findChild(mcPr, "m:count"));
      const mcJc = attr(findChild(mcPr, "m:mcJc"), "m:val");
      if (count !== undefined && mcJc) cols.push({ count, justification: mcJc });
    }
    if (cols.length > 0) result.columns = cols;
  }
  return Object.keys(result).length > 0 ? result : undefined;
}

function parseMathAccent(el: Element): MathInput {
  const accPr = findChild(el, "m:accPr");
  const chrEl = accPr ? findChild(accPr, "m:chr") : undefined;
  const accentChar = chrEl ? attr(chrEl, "m:val") : undefined;

  return {
    accent: {
      children: parseMathArg(el, "m:e"),
      ...(accentChar ? { accentCharacter: accentChar } : {}),
      ...spreadCtrlPr(accPr),
    },
  };
}

function parseMathBar(el: Element): MathInput {
  const barPr = findChild(el, "m:barPr");
  const posEl = barPr ? findChild(barPr, "m:pos") : undefined;
  const position = posEl ? attr(posEl, "m:val") : "top";

  return {
    bar: {
      children: parseMathArg(el, "m:e"),
      type: (position as "top" | "bot") ?? "top",
      ...spreadCtrlPr(barPr),
    },
  };
}

function parseMathEqArr(el: Element): MathInput {
  const rows: MathInput[][] = [];
  for (const e of children(el, "m:e")) {
    rows.push(parseMathChildren(e));
  }
  const eqArrPr = findChild(el, "m:eqArrPr");
  const result: MathEquationArrayProperties = {};
  const baseline = attr(findChild(eqArrPr, "m:baseJc"), "m:val");
  if (baseline !== undefined && Y_ALIGN.has(baseline))
    result.baseline = baseline as MathEquationArrayProperties["baseline"];
  const distributeRows = readOnOff(findChild(eqArrPr, "m:maxDist"));
  if (distributeRows !== undefined) result.distributeRows = distributeRows;
  const objectDistance = readOnOff(findChild(eqArrPr, "m:objDist"));
  if (objectDistance !== undefined) result.objectDistance = objectDistance;
  const rowSpacingRule = readNum(findChild(eqArrPr, "m:rSpRule"));
  if (rowSpacingRule !== undefined)
    result.rowSpacingRule = rowSpacingRule as MathEquationArrayProperties["rowSpacingRule"];
  const rowSpacing = readNum(findChild(eqArrPr, "m:rSp"));
  if (rowSpacing !== undefined) result.rowSpacing = rowSpacing;
  return {
    eqArr: {
      rows,
      ...(Object.keys(result).length > 0 ? { properties: result } : {}),
      ...spreadCtrlPr(eqArrPr),
    },
  };
}

function parseMathLimitLower(el: Element): MathInput {
  return {
    limitLower: {
      children: parseMathArg(el, "m:e"),
      limit: parseMathArg(el, "m:lim"),
      ...spreadCtrlPr(findChild(el, "m:limLowPr")),
    },
  };
}

function parseMathLimitUpper(el: Element): MathInput {
  return {
    limitUpper: {
      children: parseMathArg(el, "m:e"),
      limit: parseMathArg(el, "m:lim"),
      ...spreadCtrlPr(findChild(el, "m:limUppPr")),
    },
  };
}

function parseMathPreSubSuperScript(el: Element): MathInput {
  const baseArgumentProperties = readArgumentProperties(findChild(el, "m:e"));
  const subScriptArgumentProperties = readArgumentProperties(findChild(el, "m:sub"));
  const superScriptArgumentProperties = readArgumentProperties(findChild(el, "m:sup"));
  return {
    preSubSuperScript: {
      children: parseMathArg(el, "m:e"),
      subScript: parseMathArg(el, "m:sub"),
      superScript: parseMathArg(el, "m:sup"),
      ...(baseArgumentProperties !== undefined ? { baseArgumentProperties } : {}),
      ...(subScriptArgumentProperties !== undefined ? { subScriptArgumentProperties } : {}),
      ...(superScriptArgumentProperties !== undefined ? { superScriptArgumentProperties } : {}),
      ...spreadCtrlPr(findChild(el, "m:sPrePr")),
    },
  };
}

function parseMathBorderBox(el: Element): MathInput {
  const pr = findChild(el, "m:borderBoxPr");
  const result: MathBorderBoxProperties = {};
  const flags: Array<[string, keyof MathBorderBoxProperties]> = [
    ["m:hideTop", "hideTop"],
    ["m:hideBot", "hideBottom"],
    ["m:hideLeft", "hideLeft"],
    ["m:hideRight", "hideRight"],
    ["m:strikeH", "strikeHorizontal"],
    ["m:strikeV", "strikeVertical"],
    ["m:strikeBLTR", "strikeDiagonalUp"],
    ["m:strikeTLBR", "strikeDiagonalDown"],
  ];
  for (const [name, key] of flags) {
    const v = readOnOff(findChild(pr, name));
    if (v !== undefined) result[key] = v;
  }
  return {
    borderBox: {
      children: parseMathArg(el, "m:e"),
      ...(Object.keys(result).length > 0 ? { properties: result } : {}),
      ...spreadCtrlPr(pr),
    },
  };
}

function parseMathBox(el: Element): MathInput {
  const pr = findChild(el, "m:boxPr");
  const result: MathBoxProperties = {};
  const operatorEmulation = readOnOff(findChild(pr, "m:opEmu"));
  if (operatorEmulation !== undefined) result.operatorEmulation = operatorEmulation;
  const noBreak = readOnOff(findChild(pr, "m:noBreak"));
  if (noBreak !== undefined) result.noBreak = noBreak;
  const differential = readOnOff(findChild(pr, "m:diff"));
  if (differential !== undefined) result.differential = differential;
  const align = readOnOff(findChild(pr, "m:aln"));
  if (align !== undefined) result.align = align;
  return {
    box: {
      children: parseMathArg(el, "m:e"),
      ...(Object.keys(result).length > 0 ? { properties: result } : {}),
      ...spreadCtrlPr(pr),
    },
  };
}

function parseMathGroupChr(el: Element): MathInput {
  const pr = findChild(el, "m:groupChrPr");
  const result: MathGroupCharacterProperties = {};
  const character = attr(findChild(pr, "m:chr"), "m:val");
  if (character) result.character = character;
  const position = attr(findChild(pr, "m:pos"), "m:val");
  if (position === "top" || position === "bot") result.position = position;
  const vertical = attr(findChild(pr, "m:vertJc"), "m:val");
  if (vertical === "top" || vertical === "bot") result.vertical = vertical;
  return {
    groupChr: {
      children: parseMathArg(el, "m:e"),
      ...(Object.keys(result).length > 0 ? { properties: result } : {}),
      ...spreadCtrlPr(pr),
    },
  };
}

function parseMathPhant(el: Element): MathInput {
  const pr = findChild(el, "m:phantPr");
  const result: MathPhantomProperties = {};
  const show = readOnOff(findChild(pr, "m:show"));
  if (show !== undefined) result.show = show;
  const zeroWidth = readOnOff(findChild(pr, "m:zeroWid"));
  if (zeroWidth !== undefined) result.zeroWidth = zeroWidth;
  const zeroAscent = readOnOff(findChild(pr, "m:zeroAsc"));
  if (zeroAscent !== undefined) result.zeroAscent = zeroAscent;
  const zeroDescent = readOnOff(findChild(pr, "m:zeroDesc"));
  if (zeroDescent !== undefined) result.zeroDescent = zeroDescent;
  const transparent = readOnOff(findChild(pr, "m:transp"));
  if (transparent !== undefined) result.transparent = transparent;
  return {
    phant: {
      children: parseMathArg(el, "m:e"),
      ...(Object.keys(result).length > 0 ? { properties: result } : {}),
      ...spreadCtrlPr(pr),
    },
  };
}

function parseMathArg(parent: Element, childName: string): MathInput[] {
  const container = findChild(parent, childName);
  if (!container) return [];
  return parseMathChildren(container);
}
