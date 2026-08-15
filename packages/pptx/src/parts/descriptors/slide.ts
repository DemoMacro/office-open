/**
 * Slide (p:sld) descriptor for PPTX.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor, ReadContext, WriteContext } from "@office-open/core/descriptor";
import type { TextBodyOptions } from "@office-open/core/drawing";
import { attr, attrNum, findChild, stringify } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import type { ControlOptions } from "@parts/slide/slide";
import type { SlideChild as LegacySlideChild } from "@parts/slide/slide-child";
import { SP_TREE_HEADER } from "@shared/constants";
import type { SlideAnimation } from "@shared/file";
import type { SlideHeaderFooterOptions } from "@shared/header-footer";
import type { PictureOptions } from "@shared/picture";
import type { ShapeOptions } from "@shared/shape/shape";
import type { TransitionDirection, TransitionOptions } from "@shared/transition";
import { buildTransition } from "@shared/transition";

import type { BackgroundOptions } from "../background";
import { timingDesc } from "./animation";
import { backgroundDesc } from "./background";
import { parseChild } from "./bridge";
import { shapeDesc, pictureDesc } from "./shape";

// ── Types ──

export interface SlideDescriptorOptions {
  children?: SlideChild[];
  background?: BackgroundOptions;
  transition?: TransitionOptions;
  showMasterShapes?: boolean;
  showMasterPlaceholderAnimations?: boolean;
  controls?: ControlOptions[];
  customerData?: { rId: string }[];
  headerFooter?: SlideHeaderFooterOptions;
  animations?: SlideAnimation[];
  /** Hidden slide — excluded from slideshow (emits p:sld/@show="0"). */
  hidden?: boolean;
  /** Raw extLst inner XML — verbatim round-trip for unmodeled extensions. */
  ext?: string;
}

/** Discriminated union for slide children (JSON-friendly). */
export type SlideChild =
  | { shape: ShapeOptions }
  | { picture: PictureOptions }
  | { text: TextBodyOptions }
  | { contentPart: { rId: string } };

// ── Slide (p:sld) descriptor ──

export const slideDesc: CustomDescriptor<SlideDescriptorOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = [];

    // Opening tag with namespace declarations
    const sldAttrs: string[] = [];
    if (opts.showMasterShapes === false) sldAttrs.push(' showMasterSp="0"');
    if (opts.showMasterPlaceholderAnimations === false) sldAttrs.push(' showMasterPhAnim="0"');
    if (opts.hidden) sldAttrs.push(' show="0"');
    parts.push(
      `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"${sldAttrs.join("")}>`,
    );

    // p:cSld — common slide data
    parts.push("<p:cSld>");

    if (opts.background) {
      parts.push(backgroundDesc.stringify(opts.background, ctx) ?? "");
    }

    // p:spTree — shape tree
    parts.push("<p:spTree>");
    parts.push(SP_TREE_HEADER);

    if (opts.children) {
      for (const child of opts.children) {
        const xml = stringifySlideChild(child, ctx);
        if (xml) parts.push(xml);
      }
    }

    parts.push("</p:spTree>");

    // custDataLst
    if (opts.customerData && opts.customerData.length > 0) {
      const cdItems = opts.customerData.map((d) => `<p:custData r:id="${d.rId}"/>`).join("");
      parts.push(`<p:custDataLst>${cdItems}</p:custDataLst>`);
    }

    // controls
    if (opts.controls && opts.controls.length > 0) {
      const ctrlItems = opts.controls
        .map((c) => {
          const attrs: string[] = [];
          if (c.shapeId !== undefined) attrs.push(`spid="${c.shapeId}"`);
          if (c.name) attrs.push(`name="${c.name}"`);
          if (c.showAsIcon) attrs.push('showAsIcon="1"');
          if (c.rId) attrs.push(`r:id="${c.rId}"`);
          if (c.imageWidth !== undefined) attrs.push(`imgW="${c.imageWidth}"`);
          if (c.imageHeight !== undefined) attrs.push(`imgH="${c.imageHeight}"`);
          return `<p:control ${attrs.join(" ")}/>`;
        })
        .join("");
      parts.push(`<p:controls>${ctrlItems}</p:controls>`);
    }

    // Header/Footer (p:hf)
    if (opts.headerFooter) {
      const hf = opts.headerFooter;
      const hfChildren: string[] = [];
      if (hf.slideNumber) hfChildren.push("<p:sldNum/>");
      if (hf.dateTime) hfChildren.push("<p:dt/>");
      if (hf.footer) hfChildren.push("<p:ftr/>");
      if (hf.header) hfChildren.push("<p:hdr/>");
      if (hfChildren.length > 0) parts.push(`<p:hf>${hfChildren.join("")}</p:hf>`);
    }

    parts.push("</p:cSld>");

    // p:clrMapOvr
    parts.push("<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>");

    // p:transition (optional)
    if (opts.transition) {
      parts.push(stringifyTransition(opts.transition, ctx));
    }

    // p:extLst — verbatim round-trip
    if (opts.ext) parts.push(`<p:extLst>${opts.ext}</p:extLst>`);

    parts.push("</p:sld>");
    return parts.join("");
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};

    // Root attributes
    if (el.attributes) {
      if (el.attributes["showMasterSp"] !== undefined)
        result.showMasterShapes = parseOnOff(el.attributes["showMasterSp"]) ?? true;
      if (el.attributes["showMasterPhAnim"] !== undefined)
        result.showMasterPlaceholderAnimations =
          parseOnOff(el.attributes["showMasterPhAnim"]) ?? true;
      if (el.attributes["show"] === "0") result.hidden = true;
    }

    // p:cSld
    const cSld = findChild(el, "p:cSld");
    if (cSld) {
      // Background
      const bg = findChild(cSld, "p:bg");
      if (bg) result.background = backgroundDesc.parse(bg, _ctx);

      // Shape tree
      const spTree = findChild(cSld, "p:spTree");
      if (spTree) {
        const children: LegacySlideChild[] = [];
        if (spTree.elements) {
          for (const child of spTree.elements) {
            // Skip tree container structure
            if (child.name === "p:nvGrpSpPr" || child.name === "p:grpSpPr") continue;
            const parsed = parseChild(child, _ctx);
            if (parsed !== undefined) children.push(parsed);
          }
        }
        if (children.length > 0) result.children = children;
      }

      // Header/Footer (p:hf)
      const hf = findChild(cSld, "p:hf");
      if (hf) {
        const hfOpts: SlideHeaderFooterOptions = {};
        if (findChild(hf, "p:sldNum")) hfOpts.slideNumber = true;
        if (findChild(hf, "p:dt")) hfOpts.dateTime = true;
        if (findChild(hf, "p:ftr")) hfOpts.footer = true;
        if (findChild(hf, "p:hdr")) hfOpts.header = true;
        if (hfOpts.slideNumber || hfOpts.dateTime || hfOpts.footer || hfOpts.header)
          result.headerFooter = hfOpts;
      }
    }

    // p:transition
    const transition = findChild(el, "p:transition");
    if (transition) result.transition = readTransition(transition, _ctx);

    // p:timing → animations
    const timing = findChild(el, "p:timing");
    if (timing) {
      const timingOpts = timingDesc.parse(timing, _ctx);
      if (timingOpts.entries && timingOpts.entries.length > 0) {
        result.animations = timingOpts.entries;
      }
    }

    // custDataLst
    const custDataLst = findChild(el, "p:custDataLst");
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

    // controls
    const controls = findChild(el, "p:controls");
    if (controls) {
      const items: ControlOptions[] = [];
      for (const ctrl of controls.elements ?? []) {
        if (ctrl.name !== "p:control") continue;
        const item: ControlOptions = {};
        const spid = attrNum(ctrl, "spid");
        if (spid !== undefined) item.shapeId = spid;
        const name = attr(ctrl, "name");
        if (name) item.name = name;
        if (parseOnOff(attr(ctrl, "showAsIcon"))) item.showAsIcon = true;
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

    // extLst — verbatim inner XML for unmodeled extensions
    const extLst = findChild(el, "p:extLst");
    if (extLst) {
      const inner = stringify(extLst);
      if (inner) result.ext = inner;
    }

    return result as unknown as SlideDescriptorOptions;
  },
};

// ── Child serializer ──

function stringifySlideChild(child: SlideChild, ctx: WriteContext): string | undefined {
  if ("shape" in child) return shapeDesc.stringify(child.shape, ctx);
  if ("picture" in child) return pictureDesc.stringify(child.picture, ctx);
  if ("contentPart" in child) return `<p:contentPart r:id="${child.contentPart.rId}"/>`;
  return undefined;
}

// ── Transition helpers ──

export function stringifyTransition(opts: TransitionOptions, ctx?: WriteContext): string {
  if (!opts.type) return "";
  return buildTransition(opts, ctx);
}

// Reverse of DIRECTION_MAP in @shared/transition (dir attribute → semantic direction).
const XML_DIR_TO_DIRECTION: Record<string, TransitionDirection> = {
  l: "left",
  r: "right",
  u: "up",
  d: "down",
  lu: "leftUp",
  ru: "rightUp",
  ld: "leftDown",
  rd: "rightDown",
  out: "out",
  in: "in",
};

// All transitional transition element names (EG_SlideTransition). Used to
// detect the type from the matching child and read its attributes generically,
// so every supported transition round-trips without a per-type branch.
const TRANSITION_ELEMENT_TYPES = [
  "fade",
  "push",
  "wipe",
  "split",
  "blinds",
  "checker",
  "comb",
  "randomBar",
  "cover",
  "pull",
  "strips",
  "wheel",
  "zoom",
  "circle",
  "dissolve",
  "diamond",
  "newsflash",
  "plus",
  "wedge",
  "random",
  "cut",
] as const;

export function readTransition(el: XmlElement, ctx?: ReadContext): TransitionOptions {
  const result: TransitionOptions = {};

  if (el.attributes) {
    if (el.attributes["spd"] !== undefined)
      result.speed =
        el.attributes["spd"] === "med" ? "medium" : (el.attributes["spd"] as "slow" | "fast");
    if (el.attributes["advClick"] !== undefined)
      result.advanceOnClick = parseOnOff(el.attributes["advClick"]) ?? false;
    if (el.attributes["advTm"] !== undefined)
      result.advanceAfterTime = Number(el.attributes["advTm"]);
  }

  // Detect transition type from the first matching child, then read its
  // attributes (dir/orient/spokes/thruBlk) in one pass.
  for (const t of TRANSITION_ELEMENT_TYPES) {
    const child = findChild(el, `p:${t}`);
    if (!child) continue;
    result.type = t;
    const attrs = child.attributes ?? {};
    const dir = attrs["dir"];
    if (dir !== undefined) {
      const direction = XML_DIR_TO_DIRECTION[String(dir)];
      if (direction) result.direction = direction;
    }
    if (attrs["orient"] !== undefined) result.orient = attrs["orient"] as "horz" | "vert";
    if (attrs["spokes"] !== undefined) result.spokes = Number(attrs["spokes"]);
    if (attrs["thruBlk"] !== undefined) result.thruBlk = parseOnOff(attrs["thruBlk"]) ?? false;
    break;
  }

  // Sound action (p:sndAc: stSnd with embedded sound, or endSnd to stop previous).
  // The audio part bytes are read back so generate re-registers the media —
  // a bare rId would point at the wrong relationship in a fresh package.
  const sndAc = findChild(el, "p:sndAc");
  if (sndAc) {
    const stSnd = findChild(sndAc, "p:stSnd");
    if (stSnd) {
      const snd = findChild(stSnd, "p:snd");
      const rId = snd ? attr(snd, "r:embed") : undefined;
      if (rId && ctx) {
        const mediaPath = ctx.resolveRelationship(rId);
        const raw = mediaPath ? ctx.getRaw(mediaPath) : undefined;
        const type = mediaPath?.split(".").pop();
        if (raw && (type === "mp3" || type === "wav" || type === "wma" || type === "aac")) {
          const startSound: NonNullable<TransitionOptions["startSound"]> = {
            data: raw,
            type,
          };
          const name = attr(snd, "name");
          if (name) startSound.name = name;
          if (parseOnOff(attr(stSnd, "loop"))) startSound.loop = true;
          result.startSound = startSound;
        }
      }
    } else if (findChild(sndAc, "p:endSnd")) {
      result.stopPreviousSound = true;
    }
  }

  return result;
}
