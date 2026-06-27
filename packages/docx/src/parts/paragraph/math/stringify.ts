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

import { attr, children, escapeXml, findChild, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type {
  MathDelimiterProperties,
  MathInput,
  MathNaryProperties,
  MathRunPropertiesOptions,
} from "@parts/paragraph/math";

// ── MathRun properties ──

function mathRunPropsStr(opts: MathRunPropertiesOptions): string {
  const parts: string[] = [];
  if (opts.lit) parts.push('<m:lit m:val="1"/>');
  if (opts.normal) parts.push('<m:nor m:val="1"/>');
  if (opts.script) parts.push(`<m:scr m:val="${opts.script}"/>`);
  if (opts.style) parts.push(`<m:sty m:val="${opts.style}"/>`);
  if (opts.breakAlignment) parts.push(`<m:brk m:alnAt="${opts.breakAlignment}"/>`);
  if (opts.align) parts.push('<m:aln m:val="1"/>');
  return parts.length ? `<m:rPr>${parts.join("")}</m:rPr>` : "";
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

  // Discriminated union: check unique keys (order matters — subSuperScript first)
  if ("subSuperScript" in value) {
    const opts = value.subSuperScript;
    const pr = opts.alignScript
      ? '<m:sSubSupPr><m:alnScr m:val="1"/></m:sSubSupPr>'
      : "<m:sSubSupPr/>";
    return `<m:sSubSup>${pr}<m:e>${stringifyChildren(opts.children)}</m:e><m:sub>${stringifyChildren(opts.subScript)}</m:sub><m:sup>${stringifyChildren(opts.superScript)}</m:sup></m:sSubSup>`;
  }

  if ("preSubSuperScript" in value) {
    const opts = value.preSubSuperScript;
    return `<m:sPre><m:sPrePr/><m:sub>${stringifyChildren(opts.subScript)}</m:sub><m:sup>${stringifyChildren(opts.superScript)}</m:sup><m:e>${stringifyChildren(opts.children)}</m:e></m:sPre>`;
  }

  if ("superScript" in value) {
    const opts = value.superScript;
    return `<m:sSup><m:sSupPr/><m:e>${stringifyChildren(opts.children)}</m:e><m:sup>${stringifyChildren(opts.superScript)}</m:sup></m:sSup>`;
  }

  if ("subScript" in value) {
    const opts = value.subScript;
    return `<m:sSub><m:sSubPr/><m:e>${stringifyChildren(opts.children)}</m:e><m:sub>${stringifyChildren(opts.subScript)}</m:sub></m:sSub>`;
  }

  if ("fraction" in value) {
    const opts = value.fraction;
    const pr = opts.fractionType ? `<m:fPr><m:type m:val="${opts.fractionType}"/></m:fPr>` : "";
    const numArgPr = argPrXml(opts.numeratorArgumentSize);
    const denArgPr = argPrXml(opts.denominatorArgumentSize);
    return `<m:f>${pr}<m:num>${numArgPr}${stringifyChildren(opts.numerator)}</m:num><m:den>${denArgPr}${stringifyChildren(opts.denominator)}</m:den></m:f>`;
  }

  if ("radical" in value) {
    const opts = value.radical;
    const hasDegree = opts.degree && opts.degree.length > 0;
    const pr = !hasDegree ? '<m:radPr><m:degHide m:val="1"/></m:radPr>' : "<m:radPr/>";
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
    return `<m:limLow><m:e>${stringifyChildren(opts.children)}</m:e><m:lim>${stringifyChildren(opts.limit)}</m:lim></m:limLow>`;
  }

  if ("limitUpper" in value) {
    const opts = value.limitUpper;
    return `<m:limUpp><m:e>${stringifyChildren(opts.children)}</m:e><m:lim>${stringifyChildren(opts.limit)}</m:lim></m:limUpp>`;
  }

  if ("function" in value) {
    const opts = value.function;
    return `<m:func><m:fName>${stringifyChildren(opts.name)}</m:fName><m:e>${stringifyChildren(opts.children)}</m:e></m:func>`;
  }

  if ("matrix" in value) {
    const opts = value.matrix;
    const rows = opts.rows
      .map(
        (row) =>
          `<m:mr>${row.map((cell) => `<m:e>${stringifyMathInput(cell)}</m:e>`).join("")}</m:mr>`,
      )
      .join("");
    // Matrix properties are complex — emit basic structure only when needed
    let pr = "";
    if (opts.properties) {
      const p = opts.properties;
      const prParts: string[] = [];
      if (p.baseJc) prParts.push(`<m:baseJc m:val="${p.baseJc as string}"/>`);
      if (p.plcHide) prParts.push('<m:plcHide m:val="1"/>');
      if (p.rSpRule) prParts.push(`<m:rSpRule m:val="${p.rSpRule as string}"/>`);
      if (p.cGpRule) prParts.push(`<m:cGpRule m:val="${p.cGpRule as string}"/>`);
      if (p.rSp) prParts.push(`<m:rSp m:val="${p.rSp as string}"/>`);
      if (p.cSp) prParts.push(`<m:cSp m:val="${p.cSp as string}"/>`);
      if (p.cGp) prParts.push(`<m:cGp m:val="${p.cGp as string}"/>`);
      if (p.mcs) {
        const mcItems = (p.mcs as Array<{ count: number; mcJc: string }>)
          .map(
            (mc) =>
              `<m:mc><m:mcPr><m:count m:val="${mc.count}"/><m:mcJc m:val="${mc.mcJc}"/></m:mcPr></m:mc>`,
          )
          .join("");
        prParts.push(`<m:mcs>${mcItems}</m:mcs>`);
      }
      if (prParts.length) pr = `<m:mPr>${prParts.join("")}</m:mPr>`;
    }
    return `<m:m>${pr}${rows}</m:m>`;
  }

  // Bracket types
  if ("roundBrackets" in value) {
    const spec = bracketSpec(value.roundBrackets);
    return stringifyDelimiters(spec.children, "(", ")", spec.properties);
  }
  if ("curlyBrackets" in value) {
    const spec = bracketSpec(value.curlyBrackets);
    return stringifyDelimiters(spec.children, "{", "}", spec.properties);
  }
  if ("angledBrackets" in value) {
    const spec = bracketSpec(value.angledBrackets);
    return stringifyDelimiters(spec.children, "〈", "〉", spec.properties);
  }
  if ("squareBrackets" in value) {
    const spec = bracketSpec(value.squareBrackets);
    return stringifyDelimiters(spec.children, "[", "]", spec.properties);
  }

  if ("borderBox" in value) {
    const opts = value.borderBox;
    let pr = "";
    if (opts.properties) {
      const p = opts.properties;
      const parts: string[] = [];
      if (p.hideTop) parts.push('<m:hideTop m:val="1"/>');
      if (p.hideBottom) parts.push('<m:hideBot m:val="1"/>');
      if (p.hideLeft) parts.push('<m:hideLeft m:val="1"/>');
      if (p.hideRight) parts.push('<m:hideRight m:val="1"/>');
      if (p.strikeHorizontal) parts.push('<m:strikeH m:val="1"/>');
      if (p.strikeVertical) parts.push('<m:strikeV m:val="1"/>');
      if (p.strikeDiagonalUp) parts.push('<m:strikeBLTR m:val="1"/>');
      if (p.strikeDiagonalDown) parts.push('<m:strikeTLBR m:val="1"/>');
      if (parts.length) pr = `<m:borderBoxPr>${parts.join("")}</m:borderBoxPr>`;
    }
    return `<m:borderBox>${pr}<m:e>${stringifyChildren(opts.children)}</m:e></m:borderBox>`;
  }

  if ("box" in value) {
    const opts = value.box;
    let pr = "";
    if (opts.properties) {
      const p = opts.properties;
      const parts: string[] = [];
      if (p.opEmu) parts.push('<m:opEmu m:val="1"/>');
      if (p.noBreak) parts.push('<m:noBreak m:val="1"/>');
      if (p.diff) parts.push('<m:diff m:val="1"/>');
      if (p.aln) parts.push('<m:aln m:val="1"/>');
      if (parts.length) pr = `<m:boxPr>${parts.join("")}</m:boxPr>`;
    }
    return `<m:box>${pr}<m:e>${stringifyChildren(opts.children)}</m:e></m:box>`;
  }

  if ("groupChr" in value) {
    const opts = value.groupChr;
    let pr = "";
    if (opts.properties) {
      const p = opts.properties;
      const parts: string[] = [];
      if (p.chr) parts.push(`<m:chr m:val="${p.chr as string}"/>`);
      if (p.pos) parts.push(`<m:pos m:val="${p.pos as string}"/>`);
      if (p.vertJc) parts.push(`<m:vertJc m:val="${p.vertJc as string}"/>`);
      if (parts.length) pr = `<m:groupChrPr>${parts.join("")}</m:groupChrPr>`;
    }
    return `<m:groupChr>${pr}<m:e>${stringifyChildren(opts.children)}</m:e></m:groupChr>`;
  }

  if ("phant" in value) {
    const opts = value.phant;
    let pr = "";
    if (opts.properties) {
      const p = opts.properties;
      const parts: string[] = [];
      if (p.show !== undefined) parts.push(`<m:show m:val="${p.show ? 1 : 0}"/>`);
      if (p.zeroWid) parts.push('<m:zeroWid m:val="1"/>');
      if (p.zeroAsc) parts.push('<m:zeroAsc m:val="1"/>');
      if (p.zeroDesc) parts.push('<m:zeroDesc m:val="1"/>');
      if (p.transp) parts.push('<m:transp m:val="1"/>');
      if (parts.length) pr = `<m:phantPr>${parts.join("")}</m:phantPr>`;
    }
    return `<m:phant>${pr}<m:e>${stringifyChildren(opts.children)}</m:e></m:phant>`;
  }

  if ("eqArr" in value) {
    const opts = value.eqArr;
    let pr = "";
    if (opts.properties) {
      const p = opts.properties;
      const parts: string[] = [];
      if (p.baseJc) parts.push(`<m:baseJc m:val="${p.baseJc as string}"/>`);
      if (p.maxDist) parts.push('<m:maxDist m:val="1"/>');
      if (p.objDist) parts.push('<m:objDist m:val="1"/>');
      if (p.rSpRule) parts.push(`<m:rSpRule m:val="${p.rSpRule as string}"/>`);
      if (p.rSp) parts.push(`<m:rSp m:val="${p.rSp as string}"/>`);
      if (parts.length) pr = `<m:eqArrPr>${parts.join("")}</m:eqArrPr>`;
    }
    const rows = opts.rows.map((row) => `<m:e>${stringifyChildren(row)}</m:e>`).join("");
    return `<m:eqArr>${pr}${rows}</m:eqArr>`;
  }

  if ("accent" in value) {
    const opts = value.accent;
    const pr = opts.accentCharacter
      ? `<m:accPr><m:chr m:val="${opts.accentCharacter}"/></m:accPr>`
      : "";
    return `<m:acc>${pr}<m:e>${stringifyChildren(opts.children)}</m:e></m:acc>`;
  }

  if ("bar" in value) {
    const opts = value.bar;
    return `<m:bar><m:barPr><m:pos m:val="${opts.type}"/></m:barPr><m:e>${stringifyChildren(opts.children)}</m:e></m:bar>`;
  }

  // Fallback: { text: string; properties?: ... } → MathRun
  if ("text" in value) {
    const props = value.properties ? mathRunPropsStr(value.properties) : "";
    return `<m:r>${props}<m:t>${escapeXml(value.text)}</m:t></m:r>`;
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
  },
  chr: string,
): string {
  const hasSub = opts.subScript && opts.subScript.length > 0;
  const hasSup = opts.superScript && opts.superScript.length > 0;
  const prParts: string[] = [`<m:chr m:val="${chr}"/>`];
  if (opts.properties?.limitLocation)
    prParts.push(`<m:limLoc m:val="${opts.properties.limitLocation}"/>`);
  if (opts.properties?.grow !== undefined)
    prParts.push(`<m:grow m:val="${opts.properties.grow ? 1 : 0}"/>`);
  if (!hasSub) prParts.push('<m:subHide m:val="1"/>');
  if (!hasSup) prParts.push('<m:supHide m:val="1"/>');
  const pr = `<m:naryPr>${prParts.join("")}</m:naryPr>`;
  const sub = hasSub ? `<m:sub>${stringifyChildren(opts.subScript!)}</m:sub>` : "<m:sub/>";
  const sup = hasSup ? `<m:sup>${stringifyChildren(opts.superScript!)}</m:sup>` : "<m:sup/>";
  return `<m:nary>${pr}${sub}${sup}<m:e>${stringifyChildren(opts.children)}</m:e></m:nary>`;
}

// ── Delimiters (brackets) ──

function stringifyDelimiters(
  children: MathInput[],
  begChr: string,
  endChr: string,
  properties?: MathDelimiterProperties,
): string {
  const prParts: string[] = [`<m:begChr m:val="${properties?.beginCharacter ?? begChr}"/>`];
  if (properties?.separatorCharacter)
    prParts.push(`<m:sepChr m:val="${properties.separatorCharacter}"/>`);
  prParts.push(`<m:endChr m:val="${properties?.endCharacter ?? endChr}"/>`);
  if (properties?.grow !== undefined) prParts.push(`<m:grow m:val="${properties.grow ? 1 : 0}"/>`);
  if (properties?.shape) prParts.push(`<m:shp m:val="${properties.shape}"/>`);
  return `<m:d><m:dPr>${prParts.join("")}</m:dPr><m:e>${stringifyChildren(children)}</m:e></m:d>`;
}

/** Build an m:argPr/m:argSz block for an argument size scaling value. */
function argPrXml(size: number | undefined): string {
  return size !== undefined ? `<m:argPr><m:argSz m:val="${size}"/></m:argPr>` : "";
}

/** Split a bracket shorthand into children + optional delimiter properties. */
function bracketSpec(
  v: MathInput[] | { children: MathInput[]; properties?: MathDelimiterProperties },
): {
  children: MathInput[];
  properties?: MathDelimiterProperties;
} {
  if (Array.isArray(v)) return { children: v };
  return { children: v.children, properties: v.properties };
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
      return { borderBox: { children: parseMathArg(el, "m:e") } };
    case "m:box":
      return { box: { children: parseMathArg(el, "m:e") } };
    case "m:groupChr":
      return { groupChr: { children: parseMathArg(el, "m:e") } };
    case "m:phant":
      return { phant: { children: parseMathArg(el, "m:e") } };
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
    case "m:ctrlPr":
      return undefined;
    default:
      return undefined;
  }
}

function parseMathRun(el: Element): MathInput {
  const text = textOf(findChild(el, "m:t"));
  return text ?? "";
}

// ── Parse helpers ──

/** Read an m:val on/off attribute (1/0/true/false; empty element = on). */
function readOnOff(el: Element | undefined): boolean | undefined {
  if (!el) return undefined;
  const v = attr(el, "m:val");
  return v === undefined ? true : v === "1" || v === "true" || v === "on";
}

/** Read an m:val numeric attribute. */
function readNum(el: Element | undefined): number | undefined {
  if (!el) return undefined;
  const v = attr(el, "m:val");
  if (v === undefined || v === "") return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

/** Read an m:argSz scaling value from an m:argPr-bearing argument element. */
function readArgSize(argEl: Element | undefined): number | undefined {
  if (!argEl) return undefined;
  return readNum(findChild(argEl, "m:argSz"));
}

function parseMathFraction(el: Element): MathInput {
  const numeratorArgumentSize = readArgSize(findChild(el, "m:num"));
  const denominatorArgumentSize = readArgSize(findChild(el, "m:den"));
  return {
    fraction: {
      numerator: parseMathArg(el, "m:num"),
      denominator: parseMathArg(el, "m:den"),
      ...(numeratorArgumentSize !== undefined ? { numeratorArgumentSize } : {}),
      ...(denominatorArgumentSize !== undefined ? { denominatorArgumentSize } : {}),
    },
  };
}

function parseMathRadical(el: Element): MathInput {
  const degree = parseMathArg(el, "m:deg");
  const mathChildren = parseMathArg(el, "m:e");
  return {
    radical: {
      children: mathChildren,
      ...(degree.length > 0 ? { degree } : {}),
    },
  };
}

function parseMathSuperScript(el: Element): MathInput {
  return {
    superScript: {
      children: parseMathArg(el, "m:e"),
      superScript: parseMathArg(el, "m:sup"),
    },
  };
}

function parseMathSubScript(el: Element): MathInput {
  return {
    subScript: {
      children: parseMathArg(el, "m:e"),
      subScript: parseMathArg(el, "m:sub"),
    },
  };
}

function parseMathSubSuperScript(el: Element): MathInput {
  const pr = findChild(el, "m:sSubSupPr");
  const alignScript = pr ? readOnOff(findChild(pr, "m:alnScr")) : undefined;
  return {
    subSuperScript: {
      children: parseMathArg(el, "m:e"),
      subScript: parseMathArg(el, "m:sub"),
      superScript: parseMathArg(el, "m:sup"),
      ...(alignScript !== undefined ? { alignScript } : {}),
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
  };

  if (chrVal === "∑") return { sum: common };
  return { integral: common };
}

function parseMathFunction(el: Element): MathInput {
  return {
    function: {
      name: parseMathArg(el, "m:fName"),
      children: parseMathArg(el, "m:e"),
    },
  };
}

function parseMathDelimiter(el: Element): MathInput {
  const dPr = findChild(el, "m:dPr");
  const begChrEl = dPr ? findChild(dPr, "m:begChr") : undefined;
  const begChr = begChrEl ? attr(begChrEl, "m:val") : "(";
  const mathChildren = parseMathArg(el, "m:e");

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
  const hasProperties = Object.keys(properties).length > 0;
  const value = hasProperties ? { children: mathChildren, properties } : mathChildren;

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
  const rows: MathInput[][] = [];
  for (const mr of children(el, "m:mr")) {
    rows.push(parseMathArg(mr, "m:e"));
  }
  return { matrix: { rows } };
}

function parseMathAccent(el: Element): MathInput {
  const accPr = findChild(el, "m:accPr");
  const chrEl = accPr ? findChild(accPr, "m:chr") : undefined;
  const accentChar = chrEl ? attr(chrEl, "m:val") : undefined;

  return {
    accent: {
      children: parseMathArg(el, "m:e"),
      ...(accentChar ? { accentCharacter: accentChar } : {}),
    },
  };
}

function parseMathBar(el: Element): MathInput {
  const barPr = findChild(el, "m:barPr");
  const posEl = barPr ? findChild(barPr, "m:pos") : undefined;
  const pos = posEl ? attr(posEl, "m:val") : "top";

  return {
    bar: {
      children: parseMathArg(el, "m:e"),
      type: (pos as "top" | "bot") ?? "top",
    },
  };
}

function parseMathEqArr(el: Element): MathInput {
  const rows: MathInput[][] = [];
  for (const e of children(el, "m:e")) {
    rows.push(parseMathChildren(e));
  }
  return { eqArr: { rows } };
}

function parseMathLimitLower(el: Element): MathInput {
  return {
    limitLower: {
      children: parseMathArg(el, "m:e"),
      limit: parseMathArg(el, "m:lim"),
    },
  };
}

function parseMathLimitUpper(el: Element): MathInput {
  return {
    limitUpper: {
      children: parseMathArg(el, "m:e"),
      limit: parseMathArg(el, "m:lim"),
    },
  };
}

function parseMathArg(parent: Element, childName: string): MathInput[] {
  const container = findChild(parent, childName);
  if (!container) return [];
  return parseMathChildren(container);
}
