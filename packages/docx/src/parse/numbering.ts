/**
 * Numbering definition parser for DOCX documents.
 *
 * Parses w:numbering elements into NumberingOptions objects for round-trip fidelity.
 *
 * @module
 */
import { attr, attrNum, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

/**
 * Parse w:numbering element into NumberingOptions.
 */
export function parseNumberingDefinitions(el: Element): Record<string, unknown> | undefined {
  // Build a map of abstractNum elements
  const abstractNums = new Map<string, Element>();
  for (const child of el.elements ?? []) {
    if (child.name !== "w:abstractNum") continue;
    const id = attr(child, "w:abstractNumId");
    if (id !== undefined) abstractNums.set(id, child);
  }

  // Build num → abstractNumId mapping
  const numToAbstract = new Map<string, string>();
  for (const child of el.elements ?? []) {
    if (child.name !== "w:num") continue;
    const numId = attr(child, "w:numId");
    const abstractRef = findChild(child, "w:abstractNumId");
    const abstractId = abstractRef ? attr(abstractRef, "w:val") : undefined;
    if (numId !== undefined && abstractId !== undefined) {
      numToAbstract.set(numId, abstractId);
    }
  }

  // Parse each abstractNum → config entry
  const configs: { reference: string; levels: Record<string, unknown>[] }[] = [];

  for (const [numId, abstractId] of numToAbstract) {
    const abstractEl = abstractNums.get(abstractId);
    if (!abstractEl) continue;

    const levels: Record<string, unknown>[] = [];
    for (const child of abstractEl.elements ?? []) {
      if (child.name !== "w:lvl") continue;
      const levelOpts = parseLevel(child);
      if (levelOpts) levels.push(levelOpts);
    }

    // Skip default bullet list (all levels are bullet format) —
    // Numbering constructor already creates a default bullet list.
    if (levels.length > 0 && levels.every((l) => l.format === "bullet")) continue;

    if (levels.length > 0) {
      configs.push({ reference: `list_${numId}`, levels });
    }
  }

  if (configs.length === 0) return undefined;
  return { config: configs };
}

function parseLevel(el: Element): Record<string, unknown> | undefined {
  const opts: Record<string, unknown> = {};

  const level = attrNum(el, "w:ilvl");
  if (level !== undefined) opts.level = level;

  // start
  const start = findChild(el, "w:start");
  if (start) {
    const val = attrNum(start, "w:val");
    if (val !== undefined) opts.start = val;
  }

  // format
  const numFmt = findChild(el, "w:numFmt");
  if (numFmt) {
    const val = attr(numFmt, "w:val");
    if (val) opts.format = val;
  }

  // text
  const lvlText = findChild(el, "w:lvlText");
  if (lvlText) {
    const val = attr(lvlText, "w:val");
    if (val) opts.text = val;
  }

  // alignment
  const lvlJc = findChild(el, "w:lvlJc");
  if (lvlJc) {
    const val = attr(lvlJc, "w:val");
    if (val) opts.alignment = val;
  }

  // suffix
  const suff = findChild(el, "w:suff");
  if (suff) {
    const val = attr(suff, "w:val");
    if (val) opts.suffix = val;
  }

  // isLegalNumberingStyle
  if (findChild(el, "w:isLgl")) opts.isLegalNumberingStyle = true;

  // Run properties
  const rPr = findChild(el, "w:rPr");
  if (rPr) {
    const style: Record<string, unknown> = {};
    // Parse basic run properties for numbering level
    const sz = findChild(rPr, "w:sz");
    if (sz) {
      const val = attrNum(sz, "w:val");
      if (val !== undefined) {
        if (!style.run) style.run = {};
        (style.run as Record<string, unknown>).size = val;
      }
    }
    const rFonts = findChild(rPr, "w:rFonts");
    if (rFonts) {
      const ascii = attr(rFonts, "w:ascii");
      if (ascii) {
        if (!style.run) style.run = {};
        (style.run as Record<string, unknown>).font = ascii;
      }
    }
    if (Object.keys(style).length > 0) opts.style = style;
  }

  // Paragraph properties (indent)
  const pPr = findChild(el, "w:pPr");
  if (pPr) {
    const style = (opts.style as Record<string, unknown>) ?? {};
    const paraStyle: Record<string, unknown> = {};

    const ind = findChild(pPr, "w:ind");
    if (ind) {
      const indent: Record<string, unknown> = {};
      const left = attrNum(ind, "w:left");
      if (left !== undefined) indent.left = left;
      const hanging = attrNum(ind, "w:hanging");
      if (hanging !== undefined) indent.hanging = hanging;
      if (Object.keys(indent).length > 0) paraStyle.indent = indent;
    }

    const tabs = findChild(pPr, "w:tabs");
    if (tabs) {
      const tabStops: Record<string, unknown>[] = [];
      for (const tab of pPr.elements ?? []) {
        if (tab.name !== "w:tab") continue;
        const tabObj: Record<string, unknown> = {};
        const pos = attrNum(tab, "w:pos");
        if (pos !== undefined) tabObj.position = pos;
        const val = attr(tab, "w:val");
        if (val) tabObj.type = val;
        tabStops.push(tabObj);
      }
      if (tabStops.length > 0) paraStyle.tabStops = tabStops;
    }

    if (Object.keys(paraStyle).length > 0) {
      style.paragraph = paraStyle;
      opts.style = style;
    }
  }

  return Object.keys(opts).length > 0 ? opts : undefined;
}
