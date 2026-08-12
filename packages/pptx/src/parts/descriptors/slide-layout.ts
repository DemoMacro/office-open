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

import type { CustomDescriptor, ReadContext } from "@office-open/core/descriptor";
import { attr, attrNum, findChild, parse as parseXml } from "@office-open/xml";
import { NS } from "@parts/slide-layout";
import type { SlideChild } from "@parts/slide/slide-child";
import { SP_TREE_HEADER } from "@shared/constants";
import type { LayoutDefinition, LayoutPlaceholderOptions } from "@shared/file";
import { extractPlaceholderDefinition } from "@shared/placeholder";

import type { PptxWriteContext } from "../../context";
import { timingDesc } from "./animation";
import { backgroundDesc } from "./background";
import { parseChild, stringifyChild } from "./bridge";
import { colorMapOverrideDesc } from "./color-map-override";
import { readTransition, stringifyTransition } from "./slide";
import type { ControlDescriptorOptions, HeaderFooterDescriptorOptions } from "./slide";

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

/** Placeholder @type → LayoutPlaceholderOptions key. */
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
    if (opts.showMasterSp === false) attrs.push('showMasterSp="0"');
    if (opts.showMasterPhAnim === false) attrs.push('showMasterPhAnim="0"');
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

    // custDataLst (inside cSld per CT_CommonSlideData).
    if (opts.customerData && opts.customerData.length > 0) {
      const cdItems = opts.customerData.map((d) => `<p:custData r:id="${d.rId}"/>`).join("");
      parts.push(`<p:custDataLst>${cdItems}</p:custDataLst>`);
    }

    // controls (inside cSld).
    if (opts.controls && opts.controls.length > 0) {
      const ctrlItems = opts.controls
        .map((c) => {
          const a: string[] = [];
          if (c.shapeId !== undefined) a.push(`spid="${c.shapeId}"`);
          if (c.name) a.push(`name="${c.name}"`);
          if (c.showAsIcon) a.push('showAsIcon="1"');
          if (c.rId) a.push(`r:id="${c.rId}"`);
          if (c.imageWidth !== undefined) a.push(`imgW="${c.imageWidth}"`);
          if (c.imageHeight !== undefined) a.push(`imgH="${c.imageHeight}"`);
          return `<p:control ${a.join(" ")}/>`;
        })
        .join("");
      parts.push(`<p:controls>${ctrlItems}</p:controls>`);
    }

    parts.push("</p:cSld>");

    // EG_ChildSlide — p:clrMapOvr (always present; defaults to masterClrMapping).
    parts.push(
      colorMapOverrideDesc.stringify(opts.colorMapOverride, ctx) ??
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>",
    );

    // p:transition (optional, after clrMapOvr).
    if (opts.transition) {
      const transitionXml = stringifyTransition(opts.transition);
      if (transitionXml) parts.push(transitionXml);
    }

    // p:timing (optional).
    if (opts.timing) {
      parts.push(timingDesc.stringify(opts.timing, ctx) ?? "");
    }

    // p:hf (optional, sibling of cSld per CT_SlideLayout).
    if (opts.headerFooter) {
      const hf = opts.headerFooter;
      const hfChildren: string[] = [];
      if (hf.slideNumber) hfChildren.push("<p:sldNum/>");
      if (hf.dateTime) hfChildren.push("<p:dt/>");
      if (hf.footer) hfChildren.push("<p:ftr/>");
      if (hf.header) hfChildren.push("<p:hdr/>");
      if (hfChildren.length > 0) parts.push(`<p:hf>${hfChildren.join("")}</p:hf>`);
    }

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
    if (preserveAttr !== undefined) result.preserve = preserveAttr === "1";
    if (attr(el, "userDrawn") === "1") result.userDrawn = true;
    const showMasterSp = attr(el, "showMasterSp");
    if (showMasterSp !== undefined) result.showMasterSp = showMasterSp !== "0";
    const showMasterPhAnim = attr(el, "showMasterPhAnim");
    if (showMasterPhAnim !== undefined) result.showMasterPhAnim = showMasterPhAnim !== "0";

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

      // custDataLst (inside cSld).
      const custDataLst = findChild(cSld, "p:custDataLst");
      if (custDataLst) {
        const items: { rId: string }[] = [];
        for (const cd of custDataLst.elements ?? []) {
          if (cd.name === "p:custData") {
            const rId = attr(cd, "r:id");
            if (rId) items.push({ rId });
          }
        }
        if (items.length > 0) result.customerData = items;
      }

      // controls (inside cSld).
      const controls = findChild(cSld, "p:controls");
      if (controls) {
        const items: ControlDescriptorOptions[] = [];
        for (const ctrl of controls.elements ?? []) {
          if (ctrl.name !== "p:control") continue;
          const item: ControlDescriptorOptions = {};
          const spid = attrNum(ctrl, "spid");
          if (spid !== undefined) item.shapeId = spid;
          const ctrlName = attr(ctrl, "name");
          if (ctrlName) item.name = ctrlName;
          if (attr(ctrl, "showAsIcon") === "1") item.showAsIcon = true;
          const rId = attr(ctrl, "r:id");
          if (rId) item.rId = rId;
          const imgW = attrNum(ctrl, "imgW");
          if (imgW !== undefined) item.imageWidth = imgW;
          const imgH = attrNum(ctrl, "imgH");
          if (imgH !== undefined) item.imageHeight = imgH;
          items.push(item);
        }
        if (items.length > 0) result.controls = items;
      }
    }

    // EG_ChildSlide — p:clrMapOvr.
    const clrMapOvr = findChild(el, "p:clrMapOvr");
    if (clrMapOvr) result.colorMapOverride = colorMapOverrideDesc.parse(clrMapOvr, ctx);

    // p:transition.
    const transition = findChild(el, "p:transition");
    if (transition) result.transition = readTransition(transition);

    // p:timing.
    const timing = findChild(el, "p:timing");
    if (timing) result.timing = timingDesc.parse(timing, ctx);

    // p:hf (sibling of cSld).
    const hf = findChild(el, "p:hf");
    if (hf) {
      const hfOpts: HeaderFooterDescriptorOptions = {};
      if (findChild(hf, "p:sldNum")) hfOpts.slideNumber = true;
      if (findChild(hf, "p:dt")) hfOpts.dateTime = true;
      if (findChild(hf, "p:ftr")) hfOpts.footer = true;
      if (findChild(hf, "p:hdr")) hfOpts.header = true;
      if (hfOpts.slideNumber || hfOpts.dateTime || hfOpts.footer || hfOpts.header) {
        result.headerFooter = hfOpts;
      }
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
