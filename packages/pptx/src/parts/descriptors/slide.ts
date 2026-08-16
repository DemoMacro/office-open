/**
 * Slide (p:sld) descriptor for PPTX.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor, ReadContext, WriteContext } from "@office-open/core/descriptor";
import type { TextBodyOptions } from "@office-open/core/drawing";
import { attr, findChild, stringify } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { parseControls, parseCustDataLst } from "@parts/slide/c-sld";
import type { ControlOptions } from "@parts/slide/slide";
import type { SlideChild as LegacySlideChild } from "@parts/slide/slide-child";
import type { SlideAnimation, SlideOptions } from "@shared/file";
import type { SlideHeaderFooterOptions } from "@shared/header-footer";
import type { PictureOptions } from "@shared/picture";
import type { ShapeOptions } from "@shared/shape/shape";
import type { TransitionDirection, TransitionOptions } from "@shared/transition";
import { buildTransition } from "@shared/transition";

import { stringifySlide } from "../../compiler";
import type { PptxWriteContext } from "../../context";
import type { BackgroundOptions } from "../background";
import { timingDesc } from "./animation";
import { backgroundDesc } from "./background";
import { parseChild } from "./bridge";
import { colorMappingOverrideDesc, type ColorMappingOverrideOptions } from "./color-map-override";

// ── Types ──

export interface SlideDescriptorOptions {
  children?: SlideChild[];
  background?: BackgroundOptions;
  transition?: TransitionOptions;
  showMasterShapes?: boolean;
  showMasterPlaceholderAnimations?: boolean;
  controls?: ControlOptions[];
  customerData?: { rId: string }[];
  /** Instantiates dt/ftr/sldNum placeholder shapes on the slide (CT_Slide has
   * no p:hf — per-slide visibility lives in the placeholder shapes). */
  headerFooter?: SlideHeaderFooterOptions;
  colorMappingOverride?: ColorMappingOverrideOptions;
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
    // Single implementation lives in compiler.stringifySlide — the descriptor
    // used to keep a near-copy that silently dropped p:timing (animations)
    // and emitted a p:hf that CT_Slide's content model does not allow.
    return stringifySlide(opts as unknown as SlideOptions, ctx as unknown as PptxWriteContext);
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

      // p:hf is master/layout/notes-level (CT_Slide has no hf child) — nothing
      // to read at slide level.

      // custDataLst / controls — children of cSld per CT_CommonSlideData
      // (its own stringify writes them there; read from the same place).
      result.customerData = parseCustDataLst(findChild(cSld, "p:custDataLst"));
      result.controls = parseControls(findChild(cSld, "p:controls"));
    }

    // p:clrMapOvr (between cSld and transition per CT_Slide).
    const clrMapOvr = findChild(el, "p:clrMapOvr");
    if (clrMapOvr) result.colorMappingOverride = colorMappingOverrideDesc.parse(clrMapOvr, _ctx);

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

    // extLst — verbatim inner XML for unmodeled extensions
    const extLst = findChild(el, "p:extLst");
    if (extLst) {
      const inner = stringify(extLst);
      if (inner) result.ext = inner;
    }

    return result as unknown as SlideDescriptorOptions;
  },
};

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
