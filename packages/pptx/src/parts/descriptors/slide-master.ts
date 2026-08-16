/**
 * Slide Master (p:sldMaster) descriptor for PPTX.
 *
 * CT_SlideMaster — fully structured, mirroring {@link slideLayoutDesc}:
 * stringify builds cSld/bg/spTree(placeholders+children)/custDataLst/controls +
 * clrMap(EG_TopLevelSlide, required) + sldLayoutIdLst + transition + timing +
 * hf + txStyles with the `@preserve` attribute; parse extracts the same.
 *
 * Placeholder positions are scaled to the slide width on fresh generation and
 * re-derived from spTree on parse (round-trip). The master's standard text
 * styles (p:txStyles) default to the MS Office block when omitted.
 *
 * @module
 */

import { parseColorMapping, parseOnOff, stringifyColorMapping } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, findChild, stringify as stringifyXml } from "@office-open/xml";
import {
  buildBackgroundXml,
  buildPlaceholderShapes,
  type MasterPlaceholderOptions,
  type SlideMasterOptions,
} from "@parts/slide-master";
import {
  parseControls,
  parseCustDataLst,
  stringifyControls,
  stringifyCustDataLst,
} from "@parts/slide/c-sld";
import type { SlideChild } from "@parts/slide/slide-child";
import { SP_TREE_HEADER } from "@shared/constants";
import { extractPlaceholderDefinition } from "@shared/placeholder";

import type { PptxWriteContext } from "../../context";
import { buildHfAttrs, parseHeaderFooter } from "../handout-master";
import { timingDesc } from "./animation";
import { backgroundDesc } from "./background";
import { parseChild, stringifyChild } from "./bridge";
import { readTransition, stringifyTransition } from "./slide";
import {
  DEFAULT_TEXT_LIST_STYLE,
  parseTextListStyle,
  stringifyTextListStyle,
} from "./text-list-style";

// ── Types ──

/** Slide master options — structured form (CT_SlideMaster). */
export type SlideMasterDescriptorOptions = SlideMasterOptions & {
  /** cSld/`@name`. */
  name?: string;
  /** sldLayoutIdLst — layout ids + relationship ids owned by this master. */
  slideLayoutIds?: { id: number; relationshipId: string }[];
};

/**
 * Assign an explicit cNvPr id to a single-key slide-child wrapper (master
 * children start after the placeholder ids so ids stay unique per part).
 * Verbatim/rawXml children have no id slot and pass through unchanged.
 */
function withChildId(child: SlideChild, id: number): SlideChild {
  const [key, value] = Object.entries(child)[0] as [string, { id?: number } | string];
  if (typeof value !== "object") return child;
  return { [key]: { ...value, id: value.id ?? id } } as SlideChild;
}

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

/** Placeholder `@type` → MasterPlaceholderOptions key. */
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
      const xml = stringifyChild(withChildId(child, childId++), ctx);
      if (xml) parts.push(xml);
    }
    parts.push("</p:spTree>");

    // custDataLst + controls (inside cSld per CT_CommonSlideData).
    parts.push(stringifyCustDataLst(opts.customerData));
    parts.push(stringifyControls(opts.controls));

    parts.push("</p:cSld>");

    // EG_TopLevelSlide — p:clrMap (required; defaults to the standard mapping).
    parts.push(stringifyColorMapping(opts.colorMapping, "p:clrMap"));

    // p:sldLayoutIdLst (always emitted; empty when the master owns no layouts).
    const layoutIds = (opts.slideLayoutIds ?? []).map(
      (l) => `<p:sldLayoutId id="${l.id}" r:id="${l.relationshipId}"/>`,
    );
    parts.push(`<p:sldLayoutIdLst>${layoutIds.join("")}</p:sldLayoutIdLst>`);

    // p:transition (optional, after sldLayoutIdLst).
    if (opts.transition) {
      const transitionXml = stringifyTransition(opts.transition, ctx);
      if (transitionXml) parts.push(transitionXml);
    }

    // p:timing (optional).
    if (opts.animations?.length) {
      parts.push(timingDesc.stringify(opts.animations, ctx) ?? "");
    }

    // p:hf (defaults to all-hidden when omitted).
    parts.push(`<p:hf ${buildHfAttrs(opts.headerFooter)}/>`);

    // p:txStyles (defaults to the MS Office standard title/body/other block).
    parts.push(
      `<p:txStyles>${stringifyTextListStyle(opts.textStyles ?? DEFAULT_TEXT_LIST_STYLE)}</p:txStyles>`,
    );

    // p:extLst — verbatim round-trip (last child per CT_SlideMaster sequence).
    if (opts.ext) parts.push(`<p:extLst>${opts.ext}</p:extLst>`);

    parts.push("</p:sldMaster>");
    return parts.join("");
  },

  parse(el, ctx) {
    const result: Partial<SlideMasterDescriptorOptions> = {};

    // @preserve
    if (parseOnOff(attr(el, "preserve"))) result.preserve = true;

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
        if (children.length > 0) result.children = children;
        if (Object.keys(placeholders).length > 0) result.placeholders = placeholders;
      }

      // custDataLst + controls (inside cSld).
      result.customerData = parseCustDataLst(findChild(cSld, "p:custDataLst"));
      result.controls = parseControls(findChild(cSld, "p:controls"));
    }

    // EG_TopLevelSlide — p:clrMap (required).
    const colorMapping = parseColorMapping(findChild(el, "p:clrMap"));
    if (colorMapping) result.colorMapping = colorMapping;

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
    if (transition) result.transition = readTransition(transition, ctx);

    // p:timing.
    const timing = findChild(el, "p:timing");
    if (timing) {
      const entries = timingDesc.parse(timing, ctx);
      if (entries.length > 0) result.animations = entries;
    }

    // p:hf.
    const headerFooter = parseHeaderFooter(findChild(el, "p:hf"));
    if (headerFooter) result.headerFooter = headerFooter;

    // p:txStyles (CT_SlideMasterTextStyles — title/body/other).
    const txStyles = findChild(el, "p:txStyles");
    if (txStyles) result.textStyles = parseTextListStyle(txStyles);

    // p:extLst — verbatim inner XML for unmodeled extensions.
    const extLst = findChild(el, "p:extLst");
    if (extLst) {
      const inner = stringifyXml(extLst);
      if (inner) result.ext = inner;
    }

    return result as SlideMasterDescriptorOptions;
  },
};
