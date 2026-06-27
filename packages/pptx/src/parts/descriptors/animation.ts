/**
 * Animation timing (p:timing) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrNum, findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import type { AnimationEntry } from "@shared/animation/timing";
import { SlideTiming } from "@shared/animation/timing";
import type {
  AnimationClass,
  AnimationOptions,
  AnimationType,
  EmphasisType,
  PathAnimationType,
} from "@shared/animation/types";

// ── Types ──

export interface TimingDescriptorOptions {
  entries: AnimationEntry[];
}

// ── Reverse lookup maps: presetID → type name ──

const ENTR_PRESET_TO_TYPE = new Map<number, AnimationType>();
const EXIT_PRESET_TO_TYPE = new Map<number, AnimationType>();
const EMPH_PRESET_TO_ID = new Map<number, EmphasisType>();
const PATH_PRESET_TO_TYPE = new Map<number, PathAnimationType>();

const ENTR_IDS: Record<AnimationType, number> = {
  appear: 1,
  fade: 10,
  fly: 2,
  wipe: 22,
  dissolve: 34,
  split: 21,
  blinds: 25,
  checker: 26,
  randomBars: 24,
  wheel: 27,
  zoom: 53,
  cover: 28,
  push: 19,
  strips: 23,
};
const EXIT_IDS: Record<AnimationType, number> = {
  appear: 53,
  fade: 59,
  fly: 51,
  wipe: 72,
  dissolve: 84,
  split: 71,
  blinds: 75,
  checker: 76,
  randomBars: 74,
  wheel: 77,
  zoom: 60,
  cover: 78,
  push: 69,
  strips: 73,
};
const EMPH_IDS: Record<EmphasisType, number> = {
  growShrink: 53,
  spin: 54,
  growWithTurn: 56,
  colorChange: 29,
  transparency: 57,
  boldFlash: 50,
  wave: 55,
  pulse: 58,
};
const PATH_IDS: Record<PathAnimationType, number> = {
  customPath: 200,
  arc: 201,
  bounce: 202,
  circle: 203,
  curve: 204,
  figureEight: 205,
  line: 206,
  loop: 207,
};

for (const [k, v] of Object.entries(ENTR_IDS)) ENTR_PRESET_TO_TYPE.set(v, k as AnimationType);
for (const [k, v] of Object.entries(EXIT_IDS)) EXIT_PRESET_TO_TYPE.set(v, k as AnimationType);
for (const [k, v] of Object.entries(EMPH_IDS)) EMPH_PRESET_TO_ID.set(v, k as EmphasisType);
for (const [k, v] of Object.entries(PATH_IDS)) PATH_PRESET_TO_TYPE.set(v, k as PathAnimationType);

// Direction subtype reverse map
const SUBTYPE_TO_DIRECTION = new Map<number, string>();
const DIRECTION_SUBTYPES: Record<string, number> = {
  left: 4,
  right: 8,
  up: 2,
  down: 1,
  horizontal: 16,
  vertical: 32,
};
for (const [k, v] of Object.entries(DIRECTION_SUBTYPES)) SUBTYPE_TO_DIRECTION.set(v, k);

// ── Parse helpers ──

/**
 * Parse p:timing element and return a map of shape ID → AnimationOptions.
 */
function parseTiming(el: XmlElement): Map<number, AnimationOptions> {
  const result = new Map<number, AnimationOptions>();

  const tnLst = findChild(el, "p:tnLst");
  if (!tnLst) return result;

  const mainPar = findChild(tnLst, "p:par");
  if (!mainPar) return result;

  const mainCTn = findChild(mainPar, "p:cTn");
  if (!mainCTn) return result;

  const childTnLst = findChild(mainCTn, "p:childTnLst");
  if (!childTnLst) return result;

  const seq = findChild(childTnLst, "p:seq");
  if (!seq) return result;

  const seqCTn = findChild(seq, "p:cTn");
  const seqChildTnLst = seqCTn ? findChild(seqCTn, "p:childTnLst") : undefined;
  if (!seqChildTnLst) return result;

  for (const parEl of seqChildTnLst.elements ?? []) {
    if (parEl.name !== "p:par") continue;

    const parCTn = findChild(parEl, "p:cTn");
    if (!parCTn) continue;

    const parChildTnLst = findChild(parCTn, "p:childTnLst");
    if (!parChildTnLst) continue;

    for (const effectEl of parChildTnLst.elements ?? []) {
      const anim = parseAnimationEffect(effectEl);
      if (!anim) continue;

      const shapeId = extractTargetShapeId(effectEl);
      if (shapeId !== undefined) {
        result.set(shapeId, anim);
      }
    }
  }

  return result;
}

function parseAnimationEffect(el: XmlElement): AnimationOptions | undefined {
  const opts: Partial<AnimationOptions> = {};

  const cTn = findChild(el, "p:cTn") ?? el;

  const nodeType = attr(cTn, "nodeType");
  if (nodeType === "clickEffect") opts.trigger = "onClick";
  else if (nodeType === "withEffect") opts.trigger = "withPrevious";
  else if (nodeType === "afterEffect") opts.trigger = "afterPrevious";

  const presetClass = attr(cTn, "presetClass");
  if (presetClass) opts.class = presetClass as AnimationClass;

  const presetID = attrNum(cTn, "presetID");

  const dur = attr(cTn, "dur");
  if (dur) {
    const ms = parseDuration(dur);
    if (ms !== undefined) opts.duration = ms;
  }

  const stCondLst = findChild(cTn, "p:stCondLst");
  if (stCondLst) {
    const cond = findChild(stCondLst, "p:cond");
    if (cond) {
      const delay = attr(cond, "delay");
      if (delay) {
        const ms = parseDuration(delay);
        if (ms !== undefined) opts.delay = ms;
      }
    }
  }

  const presetSubtype = attrNum(cTn, "presetSubtype");
  if (presetSubtype !== undefined) {
    const dir = SUBTYPE_TO_DIRECTION.get(presetSubtype);
    if (dir) opts.direction = dir as AnimationOptions["direction"];
  }

  if (presetID !== undefined) {
    const cls = presetClass ?? "entr";
    if (cls === "entr") {
      const type = ENTR_PRESET_TO_TYPE.get(presetID);
      if (type) opts.type = type;
    } else if (cls === "exit") {
      const type = EXIT_PRESET_TO_TYPE.get(presetID);
      if (type) opts.type = type;
    } else if (cls === "emph") {
      const emphType = EMPH_PRESET_TO_ID.get(presetID);
      if (emphType) {
        opts.emphasisType = emphType;
        opts.type = "appear";
      }
    } else if (cls === "mediacall") {
      opts.mediaType = "play";
      opts.type = "appear";
    }
  }

  const childTnLst = findChild(cTn, "p:childTnLst");
  if (childTnLst) {
    for (const sub of childTnLst.elements ?? []) {
      switch (sub.name) {
        case "p:animEffect": {
          if (opts.duration === undefined) {
            const cBhvr = findChild(sub, "p:cBhvr");
            const subCTn = cBhvr ? findChild(cBhvr, "p:cTn") : undefined;
            if (subCTn) {
              const subDur = attr(subCTn, "dur");
              if (subDur) {
                const ms = parseDuration(subDur);
                if (ms !== undefined) opts.duration = ms;
              }
            }
          }
          break;
        }
        case "p:anim": {
          const subCTn = findChild(sub, "p:cTn");
          if (subCTn) {
            const attrName = findChild(subCTn, "p:attrNameLst");
            if (attrName) {
              const name = findChild(attrName, "p:attrName");
              if (name) {
                const text = name.elements?.[0]?.text;
                if (text) opts.attributeName = text as string;
              }
            }
          }
          const from = findChild(sub, "p:cb");
          if (from) {
            const val = findChild(from, "p:val");
            if (val) {
              const v = attr(val, "val");
              if (v) opts.from = v;
            }
          }
          const toEl = findChild(sub, "p:tavLst");
          if (toEl) {
            const tav = findChild(toEl, "p:tav");
            if (tav) {
              const toVal = findChild(tav, "p:val");
              if (toVal) {
                const v = attr(toVal, "val");
                if (v) opts.to = v;
              }
            }
          }
          break;
        }
        case "p:animMotion": {
          opts.pathType = "customPath" as PathAnimationType;
          const path = attr(sub, "path");
          if (path) opts.path = path;
          break;
        }
        case "p:animScale": {
          opts.emphasisType = "growShrink" as EmphasisType;
          readSubDuration(sub, opts);
          break;
        }
        case "p:animRot": {
          opts.emphasisType = "spin" as EmphasisType;
          readSubDuration(sub, opts);
          break;
        }
        case "p:animClr": {
          opts.emphasisType = "colorChange" as EmphasisType;
          readSubDuration(sub, opts);
          break;
        }
        case "p:cmd": {
          const cmdType = attr(sub, "type");
          if (cmdType === "call") opts.mediaType = "play";
          break;
        }
      }
    }
  }

  return Object.keys(opts).length > 0 ? (opts as AnimationOptions) : undefined;
}

function extractTargetShapeId(el: XmlElement): number | undefined {
  const cBhvr = findDeep(el, "p:cBhvr")[0];
  if (!cBhvr) return undefined;

  const tgtEl = findChild(cBhvr, "p:tgtEl");
  if (!tgtEl) return undefined;

  const spTgt = findChild(tgtEl, "p:spTgt");
  if (!spTgt) return undefined;

  return attrNum(spTgt, "spid");
}

function findDeep(parent: XmlElement, name: string): XmlElement[] {
  const results: XmlElement[] = [];

  function walk(el: XmlElement): void {
    if (el.name === name) {
      results.push(el);
      return;
    }
    for (const child of el.elements ?? []) {
      walk(child);
    }
  }

  for (const child of parent.elements ?? []) {
    walk(child);
  }
  return results;
}

function readSubDuration(sub: XmlElement, opts: Record<string, unknown>): void {
  if (opts.duration !== undefined) return;
  const cBhvr = findChild(sub, "p:cBhvr");
  const subCTn = cBhvr ? findChild(cBhvr, "p:cTn") : undefined;
  if (subCTn) {
    const subDur = attr(subCTn, "dur");
    if (subDur) {
      const ms = parseDuration(subDur);
      if (ms !== undefined) opts.duration = ms;
    }
  }
}

function parseDuration(val: string): number | undefined {
  if (val.startsWith("PT")) {
    let ms = 0;
    const sMatch = val.match(/(\d+\.?\d*)S/);
    if (sMatch) ms += Math.round(parseFloat(sMatch[1]) * 1000);
    const mMatch = val.match(/(\d+\.?\d*)M/);
    if (mMatch) ms += Math.round(parseFloat(mMatch[1]) * 60000);
    const hMatch = val.match(/(\d+\.?\d*)H/);
    if (hMatch) ms += Math.round(parseFloat(hMatch[1]) * 3600000);
    return ms;
  }
  if (val === "indefinite") return undefined;
  const num = parseInt(val, 10);
  return isNaN(num) ? undefined : num;
}

// ── Descriptor ──

export const timingDesc: CustomDescriptor<TimingDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    if (opts.entries.length === 0) return "";
    const timing = new SlideTiming(opts.entries);
    return timing.toXml();
  },

  parse(el, _ctx) {
    const animMap = parseTiming(el);
    const entries: AnimationEntry[] = [];
    for (const [spid, options] of animMap) {
      entries.push({ spid, options });
    }
    return { entries } as TimingDescriptorOptions;
  },
};
