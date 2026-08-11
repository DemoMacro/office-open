/**
 * Slide Master (p:sldMaster) descriptor for PPTX.
 *
 * CT_SlideMaster — fully structured, mirroring {@link slideLayoutDesc}:
 * stringify builds cSld/bg/spTree(placeholders+children)/custDataLst/controls +
 * clrMap(EG_TopLevelSlide, required) + sldLayoutIdLst + transition + timing +
 * hf + txStyles with the @preserve attribute; parse extracts the same.
 *
 * Placeholder positions are scaled to the slide width on fresh generation and
 * re-derived from spTree on parse (round-trip). The master's standard text
 * styles (p:txStyles) default to the MS Office block when omitted.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrNum, findChild } from "@office-open/xml";
import {
  buildBackgroundXml,
  buildPlaceholderShapes,
  type MasterPlaceholderOptions,
  type SlideMasterOptions,
} from "@parts/slide-master";
import type { SlideChild } from "@parts/slide/slide-child";
import { SP_TREE_HEADER } from "@shared/constants";
import type { MasterChild } from "@shared/file";
import { extractPlaceholderDefinition } from "@shared/placeholder";

import type { PptxWriteContext } from "../../context";
import {
  buildColorMapAttrs,
  buildHfAttrs,
  parseColorMap,
  parseHeaderFooter,
} from "../handout-master";
import { timingDesc } from "./animation";
import { backgroundDesc } from "./background";
import { parseChild, stringifyChild } from "./bridge";
import { readTransition, stringifyTransition } from "./slide";
import type { ControlDescriptorOptions } from "./slide";
import {
  DEFAULT_TEXT_LIST_STYLE,
  parseTextListStyle,
  stringifyTextListStyle,
} from "./text-list-style";

// ── Types ──

/** Slide master options — structured form (CT_SlideMaster). */
export type SlideMasterDescriptorOptions = SlideMasterOptions & {
  /** cSld/@name. */
  name?: string;
  /** sldLayoutIdLst — layout ids + relationship ids owned by this master. */
  slideLayoutIds?: { id: number; relationshipId: string }[];
};

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

/** Placeholder @type → MasterPlaceholderOptions key. */
const PH_TYPE_TO_KEY: Record<string, keyof MasterPlaceholderOptions> = {
  title: "title",
  ctrTitle: "title",
  body: "body",
  dt: "date",
  ftr: "footer",
  sldNum: "slideNumber",
};

// ── Descriptor ──

export const slideMasterDesc: CustomDescriptor<SlideMasterDescriptorOptions, PptxWriteContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = [];

    // Root (CT_SlideMaster: @preserve).
    const preserveAttr = opts.preserve ? ' preserve="1"' : "";
    parts.push(`<p:sldMaster ${NS}${preserveAttr}>`);

    // p:cSld (CT_CommonSlideData) — optional name attribute.
    parts.push(`<p:cSld${opts.name !== undefined ? ` name="${opts.name}"` : ""}>`);

    // p:bg — undefined background emits the MS Office default bgRef idx="1001".
    parts.push(buildBackgroundXml(opts.background));

    // p:spTree — standard placeholders (scaled to slide width) + custom children.
    parts.push("<p:spTree>");
    parts.push(SP_TREE_HEADER);
    const { xml: placeholderXml, nextId } = buildPlaceholderShapes(opts.placeholders, ctx);
    if (placeholderXml) parts.push(placeholderXml);
    // Children carry explicit cNvPr ids (starting after the placeholders) so they
    // never collide with the module-level shape id counter or the placeholders.
    let childId = nextId;
    for (const child of opts.children ?? []) {
      const shapeOpts = { ...child.shape, id: child.shape.id ?? childId++ };
      const xml = stringifyChild({ shape: shapeOpts } as SlideChild, ctx);
      if (xml) parts.push(xml);
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

    // EG_TopLevelSlide — p:clrMap (required; defaults to the standard mapping).
    parts.push(`<p:clrMap ${buildColorMapAttrs(opts.colorMap)}/>`);

    // p:sldLayoutIdLst (always emitted; empty when the master owns no layouts).
    const layoutIds = (opts.slideLayoutIds ?? []).map(
      (l) => `<p:sldLayoutId id="${l.id}" r:id="${l.relationshipId}"/>`,
    );
    parts.push(`<p:sldLayoutIdLst>${layoutIds.join("")}</p:sldLayoutIdLst>`);

    // p:transition (optional, after sldLayoutIdLst).
    if (opts.transition) {
      const transitionXml = stringifyTransition(opts.transition);
      if (transitionXml) parts.push(transitionXml);
    }

    // p:timing (optional).
    if (opts.timing) {
      parts.push(timingDesc.stringify(opts.timing, ctx) ?? "");
    }

    // p:hf (defaults to all-hidden when omitted).
    parts.push(`<p:hf ${buildHfAttrs(opts.headerFooter)}/>`);

    // p:txStyles (defaults to the MS Office standard title/body/other block).
    parts.push(
      `<p:txStyles>${stringifyTextListStyle(opts.textStyles ?? DEFAULT_TEXT_LIST_STYLE)}</p:txStyles>`,
    );

    parts.push("</p:sldMaster>");
    return parts.join("");
  },

  parse(el, ctx) {
    const result: Partial<SlideMasterDescriptorOptions> = {};

    // @preserve
    if (attr(el, "preserve") === "1") result.preserve = true;

    // p:cSld (CT_CommonSlideData).
    const cSld = findChild(el, "p:cSld");
    if (cSld) {
      const name = attr(cSld, "name");
      if (name !== undefined) result.name = name;

      const bg = findChild(cSld, "p:bg");
      if (bg) {
        const bgOpts = backgroundDesc.parse(bg, ctx);
        if (bgOpts && Object.keys(bgOpts).length > 0) result.background = bgOpts;
      }

      // spTree — structured children + derived placeholder positions.
      const spTree = findChild(cSld, "p:spTree");
      if (spTree) {
        const children: SlideChild[] = [];
        const placeholders: MasterPlaceholderOptions = {};
        for (const child of spTree.elements ?? []) {
          if (child.name === "p:nvGrpSpPr" || child.name === "p:grpSpPr") continue;
          if (child.name === "p:sp") {
            const ph = extractPlaceholderDefinition(child, ctx, PH_TYPE_TO_KEY);
            if (ph) {
              placeholders[ph.key as keyof MasterPlaceholderOptions] = ph.def;
              continue;
            }
          }
          const parsed = parseChild(child, ctx);
          if (parsed !== undefined) children.push(parsed);
        }
        // MasterChild (shared) declares only { shape }; parse yields the full
        // SlideChild union (parts-layer). The shared/parts split prevents unifying
        // them, so the cast stays local to this known impedance.
        if (children.length > 0) result.children = children as unknown as MasterChild[];
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

    // EG_TopLevelSlide — p:clrMap (required).
    const colorMap = parseColorMap(findChild(el, "p:clrMap"));
    if (colorMap) result.colorMap = colorMap;

    // p:sldLayoutIdLst.
    const sldLayoutIdLst = findChild(el, "p:sldLayoutIdLst");
    if (sldLayoutIdLst) {
      const ids = (sldLayoutIdLst.elements ?? [])
        .filter((e) => e.name === "p:sldLayoutId")
        .map((e) => ({
          id: Number(attr(e, "id")),
          relationshipId: attr(e, "r:id") ?? "",
        }));
      if (ids.length > 0) result.slideLayoutIds = ids;
    }

    // p:transition.
    const transition = findChild(el, "p:transition");
    if (transition) result.transition = readTransition(transition);

    // p:timing.
    const timing = findChild(el, "p:timing");
    if (timing) result.timing = timingDesc.parse(timing, ctx);

    // p:hf.
    const headerFooter = parseHeaderFooter(findChild(el, "p:hf"));
    if (headerFooter) result.headerFooter = headerFooter;

    // p:txStyles (CT_SlideMasterTextStyles — title/body/other).
    const txStyles = findChild(el, "p:txStyles");
    if (txStyles) result.textStyles = parseTextListStyle(txStyles);

    return result as SlideMasterDescriptorOptions;
  },
};
