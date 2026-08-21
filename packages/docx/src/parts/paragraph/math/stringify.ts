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
import type {
  MathBorderBoxProperties,
  MathBoxProperties,
  MathDelimiterProperties,
  MathEquationArrayProperties,
  MathGroupCharacterProperties,
  MathInput,
  MathMatrixColumnOptions,
  MathMatrixProperties,
  MathNaryProperties,
  MathPhantomProperties,
  MathRunPropertiesOptions,
} from "@parts/paragraph/math";
import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";

import { parseRunProperties } from "../run/run-parse";
import { stringifyRunProperties } from "../stringify";

// ── MathRun properties ──

function mathRunPropsStr(opts: MathRunPropertiesOptions): string {
  const parts: string[] = [];
  if (opts.literal) parts.push(`<m:lit m:val="${onOff(true)}"/>`);
  if (opts.normal) parts.push(`<m:nor m:val="${onOff(true)}"/>`);
  if (opts.script) parts.push(`<m:scr m:val="${opts.script}"/>`);
  if (opts.style) parts.push(`<m:sty m:val="${xsdMathStyle.to(opts.style)}"/>`);
  if (opts.breakAlignment) parts.push(`<m:brk m:alnAt="${opts.breakAlignment}"/>`);
  if (opts.align) parts.push(`<m:aln m:val="${onOff(true)}"/>`);
  return parts.length ? `<m:rPr>${parts.join("")}</m:rPr>` : "";
}

/** w:rPr serialization shared by the m:r run form and m:ctrlPr. */
function wRPrXml(opts: RunPropertiesOptions | undefined): string {
  return opts ? (stringifyRunProperties(opts) ?? "") : "";
}

/** m:ctrlPr — always the last child of a structure's *Pr element. */
function ctrlPrXml(ctrlPr: RunPropertiesOptions | undefined): string {
  if (!ctrlPr) return "";
  const rPr = wRPrXml(ctrlPr);
  return rPr ? `<m:ctrlPr>${rPr}</m:ctrlPr>` : "<m:ctrlPr/>";
}

/** Read the w:rPr inside an m:ctrlPr (CT_CtrlPr's only child). */
function readCtrlPr(prEl: Element | undefined): RunPropertiesOptions | undefined {
  const ctrlPr = prEl ? findChild(prEl, "m:ctrlPr") : undefined;
  if (!ctrlPr) return undefined;
  const rPr = findChild(ctrlPr, "w:rPr");
  return rPr ? parseRunProperties(rPr) : {};
}

// ── Children array ──

function stringifyChildren(items: MathInput[]): string {
  return items.map(stringifyMathInput).join("");
}

// ── Main recursive stringifier ──

export function stringifyMathInput(value: MathInput): string {
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
      opts.alignScript || ctrl
        ? `<m:sSubSupPr>${opts.alignScript ? `<m:alnScr m:val="${onOff(true)}"/>` : ""}${ctrl}</m:sSubSupPr>`
        : "<m:sSubSupPr/>";
    return `<m:sSubSup>${pr}<m:e>${argPrXml(opts.baseArgumentSize)}${stringifyChildren(opts.children)}</m:e><m:sub>${argPrXml(opts.subScriptArgumentSize)}${stringifyChildren(opts.subScript)}</m:sub><m:sup>${argPrXml(opts.superScriptArgumentSize)}${stringifyChildren(opts.superScript)}</m:sup></m:sSubSup>`;
  }

  if ("preSubSuperScript" in value) {
    const opts = value.preSubSuperScript;
    const ctrl = ctrlPrXml(opts.controlProperties);
    const pr = ctrl ? `<m:sPrePr>${ctrl}</m:sPrePr>` : "<m:sPrePr/>";
    return `<m:sPre>${pr}<m:sub>${argPrXml(opts.subScriptArgumentSize)}${stringifyChildren(opts.subScript)}</m:sub><m:sup>${argPrXml(opts.superScriptArgumentSize)}${stringifyChildren(opts.superScript)}</m:sup><m:e>${argPrXml(opts.baseArgumentSize)}${stringifyChildren(opts.children)}</m:e></m:sPre>`;
  }

  if ("superScript" in value) {
    const opts = value.superScript;
    const ctrl = ctrlPrXml(opts.controlProperties);
    const pr = ctrl ? `<m:sSupPr>${ctrl}</m:sSupPr>` : "<m:sSupPr/>";
    return `<m:sSup>${pr}<m:e>${argPrXml(opts.baseArgumentSize)}${stringifyChildren(opts.children)}</m:e><m:sup>${argPrXml(opts.superScriptArgumentSize)}${stringifyChildren(opts.superScript)}</m:sup></m:sSup>`;
  }

  if ("subScript" in value) {
    const opts = value.subScript;
    const ctrl = ctrlPrXml(opts.controlProperties);
    const pr = ctrl ? `<m:sSubPr>${ctrl}</m:sSubPr>` : "<m:sSubPr/>";
    return `<m:sSub>${pr}<m:e>${argPrXml(opts.baseArgumentSize)}${stringifyChildren(opts.children)}</m:e><m:sub>${argPrXml(opts.subScriptArgumentSize)}${stringifyChildren(opts.subScript)}</m:sub></m:sSub>`;
  }

  if ("fraction" in value) {
    const opts = value.fraction;
    const inner =
      (opts.fractionType ? `<m:type m:val="${opts.fractionType}"/>` : "") +
      ctrlPrXml(opts.controlProperties);
    const pr = inner ? `<m:fPr>${inner}</m:fPr>` : "";
    const numArgPr = argPrXml(opts.numeratorArgumentSize);
    const denArgPr = argPrXml(opts.denominatorArgumentSize);
    return `<m:f>${pr}<m:num>${numArgPr}${stringifyChildren(opts.numerator)}</m:num><m:den>${denArgPr}${stringifyChildren(opts.denominator)}</m:den></m:f>`;
  }

  if ("radical" in value) {
    const opts = value.radical;
    const hasDegree = opts.degree && opts.degree.length > 0;
    const hideDegree = opts.properties?.hideDegree ?? !hasDegree;
    const inner =
      (hideDegree ? `<m:degHide m:val="${onOff(true)}"/>` : "") + ctrlPrXml(opts.controlProperties);
    const pr = inner ? `<m:radPr>${inner}</m:radPr>` : "";
    const deg = hasDegree ? `<m:deg>${stringifyChildren(opts.degree!)}</m:deg>` : "<m:deg/>";
    return `<m:rad>${pr}${deg}<m:e>${stringifyChildren(opts.children)}</m:e></m:rad>`;
  }

  if ("sum" in value) {
    return stringifyNAry(value.sum, "∑");
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
                    ? cell.map(stringifyMathInput).join("")
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
    return stringifyDelimiters(spec, "(", ")");
  }
  if ("curlyBrackets" in value) {
    const spec = bracketSpec(value.curlyBrackets);
    return stringifyDelimiters(spec, "{", "}");
  }
  if ("angledBrackets" in value) {
    const spec = bracketSpec(value.angledBrackets);
    return stringifyDelimiters(spec, "〈", "〉");
  }
  if ("squareBrackets" in value) {
    const spec = bracketSpec(value.squareBrackets);
    return stringifyDelimiters(spec, "[", "]");
  }

  if ("borderBox" in value) {
    const opts = value.borderBox;
    const p = opts.properties;
    const parts: string[] = [];
    if (p?.hideTop) parts.push(`<m:hideTop m:val="${onOff(true)}"/>`);
    if (p?.hideBottom) parts.push(`<m:hideBot m:val="${onOff(true)}"/>`);
    if (p?.hideLeft) parts.push(`<m:hideLeft m:val="${onOff(true)}"/>`);
    if (p?.hideRight) parts.push(`<m:hideRight m:val="${onOff(true)}"/>`);
    if (p?.strikeHorizontal) parts.push(`<m:strikeH m:val="${onOff(true)}"/>`);
    if (p?.strikeVertical) parts.push(`<m:strikeV m:val="${onOff(true)}"/>`);
    if (p?.strikeDiagonalUp) parts.push(`<m:strikeBLTR m:val="${onOff(true)}"/>`);
    if (p?.strikeDiagonalDown) parts.push(`<m:strikeTLBR m:val="${onOff(true)}"/>`);
    parts.push(ctrlPrXml(opts.controlProperties));
    const pr = parts.some(Boolean) ? `<m:borderBoxPr>${parts.join("")}</m:borderBoxPr>` : "";
    return `<m:borderBox>${pr}<m:e>${stringifyChildren(opts.children)}</m:e></m:borderBox>`;
  }

  if ("box" in value) {
    const opts = value.box;
    const p = opts.properties;
    const parts: string[] = [];
    if (p?.operatorEmulation) parts.push(`<m:opEmu m:val="${onOff(true)}"/>`);
    if (p?.noBreak) parts.push(`<m:noBreak m:val="${onOff(true)}"/>`);
    if (p?.differential) parts.push(`<m:diff m:val="${onOff(true)}"/>`);
    if (p?.align) parts.push(`<m:aln m:val="${onOff(true)}"/>`);
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
    if (p?.zeroWidth) parts.push(`<m:zeroWid m:val="${onOff(true)}"/>`);
    if (p?.zeroAscent) parts.push(`<m:zeroAsc m:val="${onOff(true)}"/>`);
    if (p?.zeroDescent) parts.push(`<m:zeroDesc m:val="${onOff(true)}"/>`);
    if (p?.transparent) parts.push(`<m:transp m:val="${onOff(true)}"/>`);
    parts.push(ctrlPrXml(opts.controlProperties));
    const pr = parts.some(Boolean) ? `<m:phantPr>${parts.join("")}</m:phantPr>` : "";
    return `<m:phant>${pr}<m:e>${stringifyChildren(opts.children)}</m:e></m:phant>`;
  }

  if ("eqArr" in value) {
    const opts = value.eqArr;
    const p = opts.properties;
    const parts: string[] = [];
    if (p?.baseline) parts.push(`<m:baseJc m:val="${p.baseline}"/>`);
    if (p?.distributeRows) parts.push(`<m:maxDist m:val="${onOff(true)}"/>`);
    if (p?.objectDistance) parts.push(`<m:objDist m:val="${onOff(true)}"/>`);
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
    return `<m:r>${props}${rPr}<m:t>${escapeXml(value.text)}</m:t></m:r>`;
  }

  return "";
}

// ── N-ary operator (sum/integral) ──

function stringifyNAry(
  opts: {
    children: MathInput[];
    subScript?: MathInput[];
    superScript?: MathInput[];
    properties?: MathNaryProperties;
    controlProperties?: RunPropertiesOptions;
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
  if (!hasSub) prParts.push(`<m:subHide m:val="${onOff(true)}"/>`);
  if (!hasSup) prParts.push(`<m:supHide m:val="${onOff(true)}"/>`);
  prParts.push(ctrlPrXml(opts.controlProperties));
  const pr = `<m:naryPr>${prParts.join("")}</m:naryPr>`;
  const sub = hasSub ? `<m:sub>${stringifyChildren(opts.subScript!)}</m:sub>` : "<m:sub/>";
  const sup = hasSup ? `<m:sup>${stringifyChildren(opts.superScript!)}</m:sup>` : "<m:sup/>";
  return `<m:nary>${pr}${sub}${sup}<m:e>${stringifyChildren(opts.children)}</m:e></m:nary>`;
}

// ── Delimiters (brackets) ──

function stringifyDelimiters(
  spec: {
    children?: MathInput[];
    elements?: MathInput[][];
    properties?: MathDelimiterProperties;
  },
  begChr: string,
  endChr: string,
): string {
  // CT_D holds e+ — one m:e per separator-split group; a bare children array
  // is a single group.
  const groups = spec.elements ?? (spec.children ? [spec.children] : []);
  const eXml = groups.map((g) => `<m:e>${stringifyChildren(g)}</m:e>`).join("");
  // begChr/endChr are optional with XSD defaults "(" / ")" — omit when default
  // (Office writes bare <m:dPr> for round brackets).
  const beg = spec.properties?.beginCharacter ?? begChr;
  const end = spec.properties?.endCharacter ?? endChr;
  const prParts: string[] = [];
  if (beg !== "(") prParts.push(`<m:begChr m:val="${escapeXml(beg)}"/>`);
  if (spec.properties?.separatorCharacter)
    prParts.push(`<m:sepChr m:val="${escapeXml(spec.properties.separatorCharacter)}"/>`);
  if (end !== ")") prParts.push(`<m:endChr m:val="${escapeXml(end)}"/>`);
  if (spec.properties?.grow !== undefined)
    prParts.push(`<m:grow m:val="${onOff(spec.properties.grow)}"/>`);
  if (spec.properties?.shape) prParts.push(`<m:shp m:val="${spec.properties.shape}"/>`);
  prParts.push(ctrlPrXml(spec.properties?.controlProperties));
  return `<m:d><m:dPr>${prParts.join("")}</m:dPr>${eXml}</m:d>`;
}

/** Build an m:argPr/m:argSz block for an argument size scaling value. */
function argPrXml(size: number | undefined): string {
  return size !== undefined ? `<m:argPr><m:argSz m:val="${size}"/></m:argPr>` : "";
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

export function stringifyMath(children: MathInput[]): string {
  const inner = children.map((c) => stringifyMathInput(c)).join("");
  return `<m:oMath>${inner}</m:oMath>`;
}

export function stringifyMathParagraph(
  children: MathInput[],
  justification?: "left" | "right" | "center" | "centerGroup",
): string {
  const inner = children.map((c) => stringifyMathInput(c)).join("");
  const pr = justification ? `<m:oMathParaPr><m:jc m:val="${justification}"/></m:oMathParaPr>` : "";
  return `<m:oMathPara>${pr}<m:oMath>${inner}</m:oMath></m:oMathPara>`;
}

// ────────────────────────────────────────────────────────────────────────────────
// Parse (OMML XML → MathInput)
// ────────────────────────────────────────────────────────────────────────────────

/**
 * Parse all math children from an m:oMath (or similar container) element.
 */
export function parseMathChildren(el: Element): MathInput[] {
  const result: MathInput[] = [];
  for (const child of el.elements ?? []) {
    const parsed = parseMathElement(child);
    if (parsed !== undefined) result.push(parsed);
  }
  return result;
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
      const rPr = findChild(el, "w:rPr");
      return { argumentControlProperties: rPr ? parseRunProperties(rPr) : {} };
    }
    default:
      return undefined;
  }
}

function parseMathRun(el: Element): MathInput {
  const text = textOf(findChild(el, "m:t"));
  const rPrEl = findChild(el, "m:rPr");
  const wRPrEl = findChild(el, "w:rPr");
  if (!rPrEl && !wRPrEl) return text ?? "";
  const run: {
    text: string;
    properties?: MathRunPropertiesOptions;
    runProperties?: RunPropertiesOptions;
  } = {
    text: text ?? "",
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
  if (findChild(el, "m:lit")) result.literal = true;
  if (findChild(el, "m:nor")) result.normal = true;
  const scr = attr(findChild(el, "m:scr"), "m:val");
  if (scr) result.script = scr as MathRunPropertiesOptions["script"];
  const sty = attr(findChild(el, "m:sty"), "m:val");
  if (sty) result.style = xsdMathStyle.from(sty) as MathRunPropertiesOptions["style"];
  const brk = attr(findChild(el, "m:brk"), "m:alnAt");
  if (brk !== undefined && brk !== "") result.breakAlignment = Number(brk);
  if (findChild(el, "m:aln")) result.align = true;
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

/** Read an m:argSz scaling value from an argument element (m:arg > m:argPr > m:argSz). */
function readArgSize(argEl: Element | undefined): number | undefined {
  if (!argEl) return undefined;
  return readNum(findChild(findChild(argEl, "m:argPr"), "m:argSz"));
}

function parseMathFraction(el: Element): MathInput {
  const fPr = findChild(el, "m:fPr");
  const fractionType = attr(findChild(fPr, "m:type"), "m:val");
  const numeratorArgumentSize = readArgSize(findChild(el, "m:num"));
  const denominatorArgumentSize = readArgSize(findChild(el, "m:den"));
  return {
    fraction: {
      numerator: parseMathArg(el, "m:num"),
      denominator: parseMathArg(el, "m:den"),
      ...(fractionType ? { fractionType } : {}),
      ...(numeratorArgumentSize !== undefined ? { numeratorArgumentSize } : {}),
      ...(denominatorArgumentSize !== undefined ? { denominatorArgumentSize } : {}),
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
function spreadCtrlPr(prEl: Element | undefined): { controlProperties?: RunPropertiesOptions } {
  const ctrlPr = readCtrlPr(prEl);
  return ctrlPr !== undefined ? { controlProperties: ctrlPr } : {};
}

function parseMathSuperScript(el: Element): MathInput {
  const baseArgumentSize = readArgSize(findChild(el, "m:e"));
  const supArgumentSize = readArgSize(findChild(el, "m:sup"));
  return {
    superScript: {
      children: parseMathArg(el, "m:e"),
      superScript: parseMathArg(el, "m:sup"),
      ...(baseArgumentSize !== undefined ? { baseArgumentSize } : {}),
      ...(supArgumentSize !== undefined ? { superScriptArgumentSize: supArgumentSize } : {}),
      ...spreadCtrlPr(findChild(el, "m:sSupPr")),
    },
  };
}

function parseMathSubScript(el: Element): MathInput {
  const baseArgumentSize = readArgSize(findChild(el, "m:e"));
  const subArgumentSize = readArgSize(findChild(el, "m:sub"));
  return {
    subScript: {
      children: parseMathArg(el, "m:e"),
      subScript: parseMathArg(el, "m:sub"),
      ...(baseArgumentSize !== undefined ? { baseArgumentSize } : {}),
      ...(subArgumentSize !== undefined ? { subScriptArgumentSize: subArgumentSize } : {}),
      ...spreadCtrlPr(findChild(el, "m:sSubPr")),
    },
  };
}

function parseMathSubSuperScript(el: Element): MathInput {
  const pr = findChild(el, "m:sSubSupPr");
  const alignScript = pr ? readOnOff(findChild(pr, "m:alnScr")) : undefined;
  const baseArgumentSize = readArgSize(findChild(el, "m:e"));
  const subArgumentSize = readArgSize(findChild(el, "m:sub"));
  const supArgumentSize = readArgSize(findChild(el, "m:sup"));
  return {
    subSuperScript: {
      children: parseMathArg(el, "m:e"),
      subScript: parseMathArg(el, "m:sub"),
      superScript: parseMathArg(el, "m:sup"),
      ...(baseArgumentSize !== undefined ? { baseArgumentSize } : {}),
      ...(subArgumentSize !== undefined ? { subScriptArgumentSize: subArgumentSize } : {}),
      ...(supArgumentSize !== undefined ? { superScriptArgumentSize: supArgumentSize } : {}),
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
  }

  const common = {
    children: baseChildren,
    ...(sub.length > 0 ? { subScript: sub } : {}),
    ...(sup.length > 0 ? { superScript: sup } : {}),
    ...(Object.keys(properties).length > 0 ? { properties } : {}),
    ...spreadCtrlPr(naryPr),
  };

  if (chrVal === "∑") return { sum: common };
  return { integral: common };
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
  const baseArgumentSize = readArgSize(findChild(el, "m:e"));
  const subArgumentSize = readArgSize(findChild(el, "m:sub"));
  const supArgumentSize = readArgSize(findChild(el, "m:sup"));
  return {
    preSubSuperScript: {
      children: parseMathArg(el, "m:e"),
      subScript: parseMathArg(el, "m:sub"),
      superScript: parseMathArg(el, "m:sup"),
      ...(baseArgumentSize !== undefined ? { baseArgumentSize } : {}),
      ...(subArgumentSize !== undefined ? { subScriptArgumentSize: subArgumentSize } : {}),
      ...(supArgumentSize !== undefined ? { superScriptArgumentSize: supArgumentSize } : {}),
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
