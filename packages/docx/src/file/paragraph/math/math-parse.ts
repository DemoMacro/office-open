/**
 * OMML XML parser for math components.
 *
 * Parses m:oMath Element trees into MathJson format for round-trip support.
 *
 * @module
 */
import { attr, children, findChild, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { MathJson } from "./math-coerce";

/**
 * Parse all math children from an m:oMath (or similar container) element.
 *
 * Iterates over child elements and dispatches to specific parsers based on
 * element name. Unknown elements are returned as-is for passthrough.
 */
export function parseMathChildren(el: Element): MathJson[] {
  const result: MathJson[] = [];
  for (const child of el.elements ?? []) {
    const parsed = parseMathElement(child);
    if (parsed !== undefined) result.push(parsed);
  }
  return result;
}

/**
 * Parse a single math element into its MathJson representation.
 *
 * Returns `undefined` for property elements (m:*Pr) that are not standalone content.
 */
function parseMathElement(el: Element): MathJson | undefined {
  switch (el.name) {
    // m:r → text run
    case "m:r":
      return parseMathRun(el);

    // m:f → fraction
    case "m:f":
      return parseMathFraction(el);

    // m:rad → radical (root)
    case "m:rad":
      return parseMathRadical(el);

    // m:sSup → superscript
    case "m:sSup":
      return parseMathSuperScript(el);

    // m:sSub → subscript
    case "m:sSub":
      return parseMathSubScript(el);

    // m:sSubSup → combined sub/superscript
    case "m:sSubSup":
      return parseMathSubSuperScript(el);

    // m:nary → n-ary operator (sum, integral, etc.)
    case "m:nary":
      return parseMathNAry(el);

    // m:func → named function
    case "m:func":
      return parseMathFunction(el);

    // m:d → delimiter (brackets)
    case "m:d":
      return parseMathDelimiter(el);

    // m:m → matrix
    case "m:m":
      return parseMathMatrix(el);

    // m:acc → accent
    case "m:acc":
      return parseMathAccent(el);

    // m:bar → overline/underline bar
    case "m:bar":
      return parseMathBar(el);

    // m:borderBox → border box
    case "m:borderBox":
      return { borderBox: { children: parseMathArg(el, "m:e") } };

    // m:box → box
    case "m:box":
      return { box: { children: parseMathArg(el, "m:e") } };

    // m:groupChr → group character
    case "m:groupChr":
      return { groupChr: { children: parseMathArg(el, "m:e") } };

    // m:phant → phantom
    case "m:phant":
      return { phant: { children: parseMathArg(el, "m:e") } };

    // m:eqArr → equation array
    case "m:eqArr":
      return parseMathEqArr(el);

    // m:limLow → lower limit
    case "m:limLow":
      return parseMathLimitLower(el);

    // m:limUpp → upper limit
    case "m:limUpp":
      return parseMathLimitUpper(el);

    // Skip property elements — they are not standalone content
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
      // Skip non-math elements and text nodes
      return undefined;
  }
}

// ─── Specific parsers ──────────────────────────────────────────────────────

/** Parse m:r → string or { text } */
function parseMathRun(el: Element): MathJson {
  const text = textOf(findChild(el, "m:t"));
  if (text) return text;
  return text;
}

/** Parse m:f → { fraction: { numerator, denominator } } */
function parseMathFraction(el: Element): MathJson {
  return {
    fraction: {
      numerator: parseMathArg(el, "m:num"),
      denominator: parseMathArg(el, "m:den"),
    },
  };
}

/** Parse m:rad → { radical: { children, degree? } } */
function parseMathRadical(el: Element): MathJson {
  const degree = parseMathArg(el, "m:deg");
  const mathChildren = parseMathArg(el, "m:e");
  return {
    radical: {
      children: mathChildren,
      ...(degree.length > 0 ? { degree } : {}),
    },
  };
}

/** Parse m:sSup → { superScript: { children, superScript } } */
function parseMathSuperScript(el: Element): MathJson {
  return {
    superScript: {
      children: parseMathArg(el, "m:e"),
      superScript: parseMathArg(el, "m:sup"),
    },
  };
}

/** Parse m:sSub → { subScript: { children, subScript } } */
function parseMathSubScript(el: Element): MathJson {
  return {
    subScript: {
      children: parseMathArg(el, "m:e"),
      subScript: parseMathArg(el, "m:sub"),
    },
  };
}

/** Parse m:sSubSup → { subSuperScript: { children, subScript, superScript } } */
function parseMathSubSuperScript(el: Element): MathJson {
  return {
    subSuperScript: {
      children: parseMathArg(el, "m:e"),
      subScript: parseMathArg(el, "m:sub"),
      superScript: parseMathArg(el, "m:sup"),
    },
  };
}

/**
 * Parse m:nary → { sum: ... } | { integral: ... }
 *
 * Distinguishes between sum and integral based on m:naryPr/m:chr/@m:val:
 * - "∑" → sum
 * - "∫" or empty → integral
 */
function parseMathNAry(el: Element): MathJson {
  const naryPr = findChild(el, "m:naryPr");
  const chrEl = naryPr ? findChild(naryPr, "m:chr") : undefined;
  const chrVal = chrEl ? attr(chrEl, "m:val") : undefined;

  const baseChildren = parseMathArg(el, "m:e");
  const sub = parseMathArg(el, "m:sub");
  const sup = parseMathArg(el, "m:sup");

  const common = {
    children: baseChildren,
    ...(sub.length > 0 ? { subScript: sub } : {}),
    ...(sup.length > 0 ? { superScript: sup } : {}),
  };

  if (chrVal === "∑") {
    return { sum: common };
  }

  // Default to integral for "∫" or any other character
  return { integral: common };
}

/** Parse m:func → { function: { name, children } } */
function parseMathFunction(el: Element): MathJson {
  return {
    function: {
      name: parseMathArg(el, "m:fName"),
      children: parseMathArg(el, "m:e"),
    },
  };
}

/**
 * Parse m:d → { roundBrackets | squareBrackets | curlyBrackets | angledBrackets }
 *
 * Determines bracket type from m:dPr/m:begChr/@m:val:
 * - "(" → roundBrackets
 * - "[" → squareBrackets
 * - "{" → curlyBrackets
 * - default → roundBrackets (parentheses are the default delimiter)
 */
function parseMathDelimiter(el: Element): MathJson {
  const dPr = findChild(el, "m:dPr");
  const begChrEl = dPr ? findChild(dPr, "m:begChr") : undefined;
  const begChr = begChrEl ? attr(begChrEl, "m:val") : "(";

  const mathChildren = parseMathArg(el, "m:e");

  switch (begChr) {
    case "[":
      return { squareBrackets: mathChildren };
    case "{":
      return { curlyBrackets: mathChildren };
    case "<":
    case "⟨":
      return { angledBrackets: mathChildren };
    default:
      return { roundBrackets: mathChildren };
  }
}

/** Parse m:m → { matrix: { rows } } */
function parseMathMatrix(el: Element): MathJson {
  const rows: MathJson[][] = [];
  for (const mr of children(el, "m:mr")) {
    rows.push(parseMathArg(mr, "m:e"));
  }
  return { matrix: { rows } };
}

/**
 * Parse m:acc → { accent: { children, accentCharacter? } }
 *
 * Reads accent character from m:accPr/m:chr/@m:val.
 */
function parseMathAccent(el: Element): MathJson {
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

/** Parse m:bar → { bar: { children, type } } */
function parseMathBar(el: Element): MathJson {
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

/** Parse m:eqArr → { eqArr: { rows } } */
function parseMathEqArr(el: Element): MathJson {
  const rows: MathJson[][] = [];
  for (const e of children(el, "m:e")) {
    rows.push(parseMathChildren(e));
  }
  return { eqArr: { rows } };
}

/** Parse m:limLow → { limitLower: { children, limit } } */
function parseMathLimitLower(el: Element): MathJson {
  return {
    limitLower: {
      children: parseMathArg(el, "m:e"),
      limit: parseMathArg(el, "m:lim"),
    },
  };
}

/** Parse m:limUpp → { limitUpper: { children, limit } } */
function parseMathLimitUpper(el: Element): MathJson {
  return {
    limitUpper: {
      children: parseMathArg(el, "m:e"),
      limit: parseMathArg(el, "m:lim"),
    },
  };
}

// ─── Helpers ───────────────────────────────────────────────────────────────

/**
 * Parse a container element's math children.
 *
 * Used for m:e, m:num, m:den, m:sub, m:sup, m:lim, m:fName, m:deg
 * — all of which are CT_OMathArg containers that hold math child elements.
 */
function parseMathArg(parent: Element, childName: string): MathJson[] {
  const container = findChild(parent, childName);
  if (!container) return [];
  return parseMathChildren(container);
}
