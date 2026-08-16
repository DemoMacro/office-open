/**
 * Slide Layout (p:sldLayout) descriptor for PPTX.
 *
 * CT_SlideLayout — fully structured. Mirrors {@link slideDesc}: stringify builds
 * cSld/bg/spTree/custDataLst/controls + clrMapOvr/transition/timing/hf with all
 * CT_SlideLayout attributes; parse extracts the same. Placeholder positions are
 * derived from spTree on parse (read-only helper for the fresh/editor path).
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor, ReadContext } from "@office-open/core/descriptor";
import { attr, findChild, parse as parseXml, stringify as stringifyXml } from "@office-open/xml";
import { NS } from "@parts/slide-layout";
import {
  parseControls,
  parseCustDataLst,
  parseSlideHf,
  stringifyControls,
  stringifyCustDataLst,
  stringifySlideHf,
} from "@parts/slide/c-sld";
import type { SlideChild } from "@parts/slide/slide-child";
import { SP_TREE_HEADER } from "@shared/constants";
import type { LayoutDefinition, LayoutPlaceholderOptions } from "@shared/file";
import { extractPlaceholderDefinition } from "@shared/placeholder";

import type { PptxWriteContext } from "../../context";
import { timingDesc } from "./animation";
import { backgroundDesc } from "./background";
import { parseChild, stringifyChild } from "./bridge";
import { colorMappingOverrideDesc } from "./color-map-override";
import { readTransition, stringifyTransition } from "./slide";

// ── Display name → SlideLayoutType mapping (fallback when @type absent) ──

const NAME_TO_TYPE: Record<string, string> = {
  "Title Slide": "title",
  "Title and Content": "obj",
  "Section Header": "secHead",
  "Two Content": "twoObj",
  Comparison: "twoTxTwoObj",
  "Title Only": "titleOnly",
  Blank: "blank",
  "Content with Caption": "objTx",
  "Picture with Caption": "picTx",
  "Vertical Text": "vertTx",
  "Vertical Title and Text": "vertTitleAndTx",
  "Title and Text": "tx",
};

/** Placeholder `@type` → LayoutPlaceholderOptions key. */
const PH_TYPE_TO_KEY: Record<string, keyof LayoutPlaceholderOptions> = {
  title: "title",
  ctrTitle: "title",
  body: "body",
  sub: "subtitle",
  dt: "date",
  ftr: "footer",
  sldNum: "slideNumber",
};

// ── Descriptor ──

export const slideLayoutDesc: CustomDescriptor<LayoutDefinition, PptxWriteContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = [];

    // Root attributes (CT_SlideLayout: type/matchingName/preserve/userDrawn + AG_ChildSlide).
    const attrs: string[] = [`type="${opts.type ?? "cust"}"`];
    if (opts.matchingName !== undefined) attrs.push(`matchingName="${opts.matchingName}"`);
    attrs.push(`preserve="${opts.preserve ? 1 : 0}"`);
    if (opts.userDrawn) attrs.push('userDrawn="1"');
    if (opts.showMasterShapes === false) attrs.push('showMasterSp="0"');
    if (opts.showMasterPlaceholderAnimations === false) attrs.push('showMasterPhAnim="0"');
    parts.push(`<p:sldLayout ${NS} ${attrs.join(" ")}>`);

    // p:cSld (CT_CommonSlideData) — optional name attribute.
    parts.push(`<p:cSld${opts.name !== undefined ? ` name="${opts.name}"` : ""}>`);

    if (opts.background) {
      parts.push(backgroundDesc.stringify(opts.background, ctx) ?? "");
    }

    // p:spTree
    parts.push("<p:spTree>");
    parts.push(SP_TREE_HEADER);
    if (opts.children) {
      for (const child of opts.children) {
        const xml = stringifyChild(child, ctx);
        if (xml) parts.push(xml);
      }
    }
    parts.push("</p:spTree>");

    // custDataLst + controls (inside cSld per CT_CommonSlideData).
    parts.push(stringifyCustDataLst(opts.customerData));
    parts.push(stringifyControls(opts.controls));

    parts.push("</p:cSld>");

    // EG_ChildSlide — p:clrMapOvr (always present; defaults to masterClrMapping).
    parts.push(
      colorMappingOverrideDesc.stringify(opts.colorMappingOverride, ctx) ??
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>",
    );

    // p:transition (optional, after clrMapOvr).
    if (opts.transition) {
      const transitionXml = stringifyTransition(opts.transition, ctx);
      if (transitionXml) parts.push(transitionXml);
    }

    // p:timing (optional).
    if (opts.animations?.length) {
      parts.push(timingDesc.stringify(opts.animations, ctx) ?? "");
    }

    // p:hf (optional, sibling of cSld per CT_SlideLayout; attribute form per
    // CT_HeaderFooter).
    parts.push(stringifySlideHf(opts.headerFooter));

    // p:extLst — verbatim round-trip (last child per CT_SlideLayout sequence).
    if (opts.ext) parts.push(`<p:extLst>${opts.ext}</p:extLst>`);

    parts.push("</p:sldLayout>");
    return parts.join("");
  },

  parse(el, ctx) {
    const result: Partial<LayoutDefinition> = {};

    // Root attributes (CT_SlideLayout).
    const typeAttr = attr(el, "type");
    if (typeAttr) result.type = typeAttr;
    const matchingName = attr(el, "matchingName");
    if (matchingName !== undefined) result.matchingName = matchingName;
    const preserveAttr = attr(el, "preserve");
    if (preserveAttr !== undefined) result.preserve = parseOnOff(preserveAttr) ?? false;
    if (parseOnOff(attr(el, "userDrawn"))) result.userDrawn = true;
    const showMasterShapes = attr(el, "showMasterSp");
    if (showMasterShapes !== undefined)
      result.showMasterShapes = parseOnOff(showMasterShapes) ?? true;
    const showMasterPlaceholderAnimations = attr(el, "showMasterPhAnim");
    if (showMasterPlaceholderAnimations !== undefined)
      result.showMasterPlaceholderAnimations = parseOnOff(showMasterPlaceholderAnimations) ?? true;

    // p:cSld (CT_CommonSlideData).
    const cSld = findChild(el, "p:cSld");
    if (cSld) {
      const name = attr(cSld, "name");
      if (name !== undefined) {
        result.name = name;
        if (!result.type) result.type = NAME_TO_TYPE[name] ?? name;
      }

      const bg = findChild(cSld, "p:bg");
      if (bg) result.background = backgroundDesc.parse(bg, ctx);

      // spTree — structured children + derived placeholder positions.
      const spTree = findChild(cSld, "p:spTree");
      if (spTree) {
        const children: SlideChild[] = [];
        const placeholders: LayoutPlaceholderOptions = {};
        for (const child of spTree.elements ?? []) {
          if (child.name === "p:nvGrpSpPr" || child.name === "p:grpSpPr") continue;
          const parsed = parseChild(child, ctx);
          if (parsed !== undefined) children.push(parsed);
          if (child.name === "p:sp") {
            const ph = extractPlaceholderDefinition(child, ctx, PH_TYPE_TO_KEY);
            if (ph) placeholders[ph.key as keyof LayoutPlaceholderOptions] = ph.def;
          }
        }
        if (children.length > 0) result.children = children;
        if (Object.keys(placeholders).length > 0) result.placeholders = placeholders;
      }

      // custDataLst + controls (inside cSld).
      result.customerData = parseCustDataLst(findChild(cSld, "p:custDataLst"));
      result.controls = parseControls(findChild(cSld, "p:controls"));
    }

    // EG_ChildSlide — p:clrMapOvr.
    const clrMapOvr = findChild(el, "p:clrMapOvr");
    if (clrMapOvr) result.colorMappingOverride = colorMappingOverrideDesc.parse(clrMapOvr, ctx);

    // p:transition.
    const transition = findChild(el, "p:transition");
    if (transition) result.transition = readTransition(transition, ctx);

    // p:timing.
    const timing = findChild(el, "p:timing");
    if (timing) {
      const entries = timingDesc.parse(timing, ctx);
      if (entries.length > 0) result.animations = entries;
    }

    // p:hf (sibling of cSld; attribute form per CT_HeaderFooter).
    result.headerFooter = parseSlideHf(findChild(el, "p:hf"));

    // p:extLst — verbatim inner XML for unmodeled extensions.
    const extLst = findChild(el, "p:extLst");
    if (extLst) {
      const inner = stringifyXml(extLst);
      if (inner) result.ext = inner;
    }

    return result as LayoutDefinition;
  },
};

// ── Template parsing (compiler generate-path) ──

// Stub ReadContext for parsing pre-built template/custom layout XML. Template
// layouts carry no relatable media, so relationship resolution is unnecessary.
const STUB_READ_CONTEXT = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

/**
 * Parse a pre-built slideLayout XML string into a structured LayoutDefinition.
 *
 * Used by the compiler to convert fresh template/custom layout XML (produced by
 * buildLayoutXml/buildCustomLayoutXml) into structured form, so every layout
 * flows through {@link slideLayoutDesc.stringify} uniformly.
 */
export function parseLayoutDef(xml: string): LayoutDefinition {
  const doc = parseXml(xml);
  const root = doc.elements?.[0];
  if (!root) throw new Error("parsed slideLayout document has no root element");
  return slideLayoutDesc.parse(root, STUB_READ_CONTEXT);
}
