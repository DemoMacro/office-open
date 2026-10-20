/**
 * Slide (p:sld) descriptor for PPTX.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import { xsdOrient } from "@office-open/core";
import type { CustomDescriptor, ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parseNonVisualDrawingProperties } from "@office-open/core/drawing";
import { attr, attrNum, findChild, stringify } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { parseControls, parseCustDataLst } from "@parts/slide/c-sld";
import type { SlideChild } from "@parts/slide/slide-child";
import type { SlideOptions } from "@shared/file";
import type { TransitionDirection, TransitionOptions } from "@shared/transition";
import { buildTransition } from "@shared/transition";

import { stringifySlide } from "../../compile/slides";
import type { PptxWriteContext } from "../../context";
import { timingDesc } from "./animation";
import { backgroundDesc } from "./background";
import { parseChild } from "./bridge";
import { colorMappingOverrideDesc } from "./color-map-override";

// ── Types ──

// ── Slide (p:sld) descriptor ──

export const slideDesc: CustomDescriptor<SlideOptions, PptxWriteContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    // Single implementation lives in compiler.stringifySlide — the descriptor
    // used to keep a near-copy that silently dropped p:timing (animations)
    // and emitted a p:hf that CT_Slide's content model does not allow.
    return stringifySlide(opts, ctx);
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};

    // Root attributes
    if (el.attributes) {
      if (el.attributes["showMasterSp"] !== undefined)
        result.showMasterShapes = parseOnOff(el.attributes["showMasterSp"]) ?? true;
      if (el.attributes["showMasterSp"] !== undefined) result.showMasterShapesPresent = true;
      if (el.attributes["showMasterPhAnim"] !== undefined)
        result.showMasterPlaceholderAnimations =
          parseOnOff(el.attributes["showMasterPhAnim"]) ?? true;
      if (el.attributes["showMasterPhAnim"] !== undefined)
        result.showMasterPlaceholderAnimationsPresent = true;
      const show = el.attributeRawValues?.["show"] ?? el.attributes?.["show"];
      if (show !== undefined) {
        const isHidden = show === "0" || parseOnOff(show) === false;
        result.showRaw = isHidden ? "0" : "1";
        if (isHidden) result.hidden = true;
      }
    }

    // p:cSld
    const cSld = findChild(el, "p:cSld");
    if (cSld) {
      const name = attr(cSld, "name");
      if (name !== undefined) result.name = name;
      // Background
      const bg = findChild(cSld, "p:bg");
      if (bg) result.background = backgroundDesc.parse(bg, _ctx);

      // Shape tree
      const spTree = findChild(cSld, "p:spTree");
      if (spTree) {
        const spTreeGroupProperties = findChild(spTree, "p:grpSpPr");
        if (spTreeGroupProperties && findChild(spTreeGroupProperties, "a:xfrm") === undefined) {
          result.groupTransformless = true;
        }
        const spTreeBwMode = spTreeGroupProperties?.attributes?.["bwMode"] as string | undefined;
        if (spTreeBwMode !== undefined) result.shapeTreeBlackWhiteMode = spTreeBwMode;
        const nvGrpSpPr = findChild(spTree, "p:nvGrpSpPr");
        const rootCnvPr = nvGrpSpPr ? findChild(nvGrpSpPr, "p:cNvPr") : undefined;
        const rootProperties = parseNonVisualDrawingProperties(rootCnvPr);
        if (rootCnvPr) {
          const shapeTreeId = attrNum(rootCnvPr, "id");
          if (shapeTreeId !== undefined) result.shapeTreeId = shapeTreeId;
          if (rootProperties.name !== undefined) result.shapeTreeName = rootProperties.name;
          if (rootProperties.creationId !== undefined)
            result.shapeTreeCreationId = rootProperties.creationId;
          if (rootProperties.ext !== undefined) result.shapeTreeExt = rootProperties.ext;
        }
        const children: SlideChild[] = [];
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
      result.controls = parseControls(findChild(cSld, "p:controls"), (picture) => {
        const child = parseChild(picture, _ctx);
        return child && "picture" in child ? child.picture : undefined;
      });

      // cSld-tail extLst (p14:creationId's home) — verbatim, distinct from the
      // root-level extLst read below (two separate XSD slots).
      const cSldExtLst = findChild(cSld, "p:extLst");
      if (cSldExtLst) {
        const inner = stringify(cSldExtLst);
        if (inner) result.cSldExt = inner;
      }
    }

    // p:clrMapOvr (between cSld and transition per CT_Slide).
    const clrMapOvr = findChild(el, "p:clrMapOvr");
    if (clrMapOvr) result.colorMappingOverride = colorMappingOverrideDesc.parse(clrMapOvr, _ctx);

    result.transition = parseTransitionElement(el, _ctx);

    // p:timing → animations (structured entries, or verbatim inner XML when
    // the source tree exceeds the model)
    const timing = findChild(el, "p:timing");
    if (timing) {
      const animations = timingDesc.parse(timing, _ctx);
      if (!(Array.isArray(animations) && animations.length === 0)) result.animations = animations;
    }

    // extLst — verbatim inner XML for unmodeled extensions
    const extLst = findChild(el, "p:extLst");
    if (extLst) {
      const inner = stringify(extLst);
      if (inner) result.ext = inner;
    }

    return result as Partial<SlideOptions> as SlideOptions;
  },
};

// ── Transition helpers ──

/** True when a descendant of any depth carries the element name. */
function hasDescendant(el: XmlElement, name: string): boolean {
  for (const child of el.elements ?? []) {
    if (child.name === name || hasDescendant(child, name)) return true;
  }
  return false;
}

export function stringifyTransition(opts: TransitionOptions | string, ctx?: WriteContext): string {
  // Verbatim markup-compatibility block (mc:Choice p14:dur + mc:Fallback).
  if (typeof opts === "string") return opts;
  // A typeless transition stays legal — buildTransition emits the bare
  // <p:transition/> (attributes only) the CT_SlideTransition content model allows.
  return buildTransition(opts, ctx);
}

/** Parse either a direct transition or its markup-compatibility block verbatim. */
export function parseTransitionElement(
  el: XmlElement,
  ctx?: ReadContext,
): TransitionOptions | string | undefined {
  const transition = findChild(el, "p:transition");
  if (transition) return readTransition(transition, ctx);

  const markupCompatibilityBlock = (el.elements ?? []).find(
    (child) => child.name === "mc:AlternateContent" && hasDescendant(child, "p:transition"),
  );
  if (!markupCompatibilityBlock) return undefined;
  return stringify({ elements: [markupCompatibilityBlock] });
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

const ORIENTATION_TRANSITION_TYPES = new Set(["blinds", "checker", "comb", "randomBar"]);

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
    if (el.attributes["p14:dur"] !== undefined) result.duration = Number(el.attributes["p14:dur"]);
  }

  // Detect transition type from the first matching child, then read its
  // attributes (dir/orient/spokes/thruBlk) in one pass.
  for (const t of TRANSITION_ELEMENT_TYPES) {
    const child = findChild(el, `p:${t}`);
    if (!child) continue;
    result.type = t;
    const attrs = child.attributes ?? {};
    const dir = attrs["dir"];
    if (dir !== undefined && ORIENTATION_TRANSITION_TYPES.has(t)) {
      result.orient = xsdOrient.from(String(dir)) as "horizontal" | "vertical";
    } else if (dir !== undefined) {
      const direction = XML_DIR_TO_DIRECTION[String(dir)];
      if (direction) result.direction = direction;
    }
    if (attrs["orient"] !== undefined)
      result.orient = xsdOrient.from(String(attrs["orient"])) as "horizontal" | "vertical";
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
        const type = mediaPath === undefined ? undefined : mediaPath.split(".").pop();
        if (raw && (type === "mp3" || type === "wav" || type === "wma" || type === "aac")) {
          const startSound: NonNullable<TransitionOptions["startSound"]> = {
            data: raw,
            type,
            fileName: mediaPath === undefined ? undefined : mediaPath.split("/").pop(),
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
