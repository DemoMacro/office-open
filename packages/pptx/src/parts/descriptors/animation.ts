/**
 * Animation timing (p:timing) descriptor for PPTX.
 *
 * @module
 */

import {
  parseOnOff,
  xsdAnimCalcMode,
  xsdAnimClass,
  xsdAnimValueType,
  xsdIterateType,
} from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrNum, findChild, findFirst, stringify as stringifyXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import type { AnimationEntry, AnimationsOptions } from "@shared/animation/timing";
import { SlideTiming } from "@shared/animation/timing";
import {
  DIRECTION_SUBTYPES,
  EMPH_PRESET_IDS,
  ENTR_PRESET_IDS,
  EXIT_PRESET_IDS,
  PATH_PRESET_IDS,
} from "@shared/animation/timing";
import type {
  AnimationClass,
  AnimationOptions,
  AnimationType,
  EmphasisType,
  PathAnimationType,
} from "@shared/animation/types";

// ── Reverse lookup maps: presetID → type name ──
// Built from the single stringify-side tables in shared/animation/timing.ts.

const ENTR_PRESET_TO_TYPE = new Map<number, AnimationType>();
const EXIT_PRESET_TO_TYPE = new Map<number, AnimationType>();
const EMPH_PRESET_TO_ID = new Map<number, EmphasisType>();
const PATH_PRESET_TO_TYPE = new Map<number, PathAnimationType>();

for (const [k, v] of Object.entries(ENTR_PRESET_IDS))
  ENTR_PRESET_TO_TYPE.set(v, k as AnimationType);
for (const [k, v] of Object.entries(EXIT_PRESET_IDS))
  EXIT_PRESET_TO_TYPE.set(v, k as AnimationType);
for (const [k, v] of Object.entries(EMPH_PRESET_IDS)) EMPH_PRESET_TO_ID.set(v, k as EmphasisType);
for (const [k, v] of Object.entries(PATH_PRESET_IDS))
  PATH_PRESET_TO_TYPE.set(v, k as PathAnimationType);

// Direction subtype reverse map
const SUBTYPE_TO_DIRECTION = new Map<number, string>();
for (const [k, v] of Object.entries(DIRECTION_SUBTYPES)) SUBTYPE_TO_DIRECTION.set(v, k);

// ── Parse helpers ──

/**
 * Parse p:timing element and return animation entries grouped by shape ID.
 * A shape can carry several effects (e.g. entrance + emphasis), so each
 * parsed effect becomes its own entry instead of overwriting the previous one.
 */
function parseTiming(el: XmlElement): Map<number, AnimationOptions[]> {
  const result = new Map<number, AnimationOptions[]>();

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
        const list = result.get(shapeId);
        if (list) list.push(anim);
        else result.set(shapeId, [anim]);
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

  const presetClassAttr = attr(cTn, "presetClass");
  const presetClass = presetClassAttr
    ? (xsdAnimClass.from(presetClassAttr) as AnimationClass)
    : undefined;
  if (presetClass) opts.class = presetClass;

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
    const cls = presetClass ?? "entrance";
    if (cls === "entrance") {
      const type = ENTR_PRESET_TO_TYPE.get(presetID);
      if (type) opts.type = type;
    } else if (cls === "exit") {
      const type = EXIT_PRESET_TO_TYPE.get(presetID);
      if (type) opts.type = type;
    } else if (cls === "emphasis") {
      const emphType = EMPH_PRESET_TO_ID.get(presetID);
      if (emphType) {
        opts.emphasisType = emphType;
        opts.type = "appear";
      }
    } else if (cls === "mediaCall") {
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
          // CT_TLAnimateBehavior attributes — from/to/by/calcmode/valueType
          const calcMode = attr(sub, "calcmode");
          if (calcMode)
            opts.calcMode = xsdAnimCalcMode.from(calcMode) as AnimationOptions["calcMode"];
          const valueType = attr(sub, "valueType");
          if (valueType)
            opts.valueType = xsdAnimValueType.from(valueType) as AnimationOptions["valueType"];
          const fromAttr = attr(sub, "from");
          if (fromAttr) opts.from = fromAttr;
          const toAttr = attr(sub, "to");
          if (toAttr) opts.to = toAttr;
          const byAttr = attr(sub, "by");
          if (byAttr) opts.animBy = byAttr;

          const cBhvr = findChild(sub, "p:cBhvr");
          if (cBhvr) {
            const attrName = findChild(cBhvr, "p:attrNameLst");
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
          // @rAng is ST_Angle (1/60000 degree); the API exposes degrees.
          const rAng = attrNum(sub, "rAng");
          if (rAng !== undefined) opts.rotationAngle = rAng / 60000;
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
          const clrSpc = attr(sub, "clrSpc");
          if (clrSpc) opts.colorSpace = clrSpc as AnimationOptions["colorSpace"];
          readSubDuration(sub, opts);
          break;
        }
        case "p:cmd": {
          // mediaType belongs to mediacall-preset effects (set via presetClass);
          // a bare command must stay a command or the rebuild path would emit
          // playFrom(0.0) instead of the authored cmd string.
          const cmdType = attr(sub, "type");
          if (cmdType) opts.commandType = cmdType as AnimationOptions["commandType"];
          const cmdStr = attr(sub, "cmd");
          if (cmdStr) opts.command = cmdStr;
          break;
        }
      }
    }
  }

  // p:iterate — direct child of cTn (CT_TLCommonTimeNodeData sequence places
  // it before childTnLst, not inside it). tmPct carries percent, tmAbs ms.
  const iterateEl = findChild(cTn, "p:iterate");
  if (iterateEl) {
    const iterate: NonNullable<AnimationOptions["iterate"]> = {};
    const iterType = attr(iterateEl, "type");
    if (iterType)
      iterate.type = xsdIterateType.from(iterType) as NonNullable<
        AnimationOptions["iterate"]
      >["type"];
    if (parseOnOff(attr(iterateEl, "backwards"))) iterate.backwards = true;
    const tmPct = findChild(iterateEl, "p:tmPct");
    if (tmPct) {
      const v = attrNum(tmPct, "val");
      if (v !== undefined) iterate.iteratePercentage = v / 1000;
    } else {
      const tmAbs = findChild(iterateEl, "p:tmAbs");
      if (tmAbs) {
        const v = attrNum(tmAbs, "val");
        if (v !== undefined) iterate.interval = v;
      }
    }
    opts.iterate = iterate;
  }

  return Object.keys(opts).length > 0 ? (opts as AnimationOptions) : undefined;
}

function extractTargetShapeId(el: XmlElement): number | undefined {
  const cBhvr = findFirst(el, "p:cBhvr");
  if (!cBhvr) return undefined;

  const tgtEl = findChild(cBhvr, "p:tgtEl");
  if (!tgtEl) return undefined;

  const spTgt = findChild(tgtEl, "p:spTgt");
  if (!spTgt) return undefined;

  return attrNum(spTgt, "spid");
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
    if (sMatch) ms += Math.round(parseFloat(sMatch[1] ?? "") * 1000);
    const mMatch = val.match(/(\d+\.?\d*)M/);
    if (mMatch) ms += Math.round(parseFloat(mMatch[1] ?? "") * 60000);
    const hMatch = val.match(/(\d+\.?\d*)H/);
    if (hMatch) ms += Math.round(parseFloat(hMatch[1] ?? "") * 3600000);
    return ms;
  }
  if (val === "indefinite") return undefined;
  const num = parseInt(val, 10);
  return isNaN(num) ? undefined : num;
}

// ── Descriptor ──

export const timingDesc: CustomDescriptor<AnimationsOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    // Verbatim fallback — the source tree the structured model cannot rebuild.
    if (typeof opts === "string") return `<p:timing>${opts}</p:timing>`;
    if (opts.length === 0) return "";
    const timing = new SlideTiming(opts);
    return timing.toXml();
  },

  parse(el, _ctx) {
    const animMap = parseTiming(el);
    const entries: AnimationEntry[] = [];
    for (const [shapeId, optionsList] of animMap) {
      for (const options of optionsList) {
        entries.push({ shapeId, ...options });
      }
    }
    // Fidelity gate: rebuilding reorganizes the timing tree, so compare the
    // rebuilt tag multiset against the source — any drift (including trees the
    // model extracts no entries from, like tmRoot-only timing) falls back to
    // the verbatim inner XML rather than silently losing or reshaping nodes.
    const rebuilt = new SlideTiming(entries).toXml();
    const source = `<p:timing>${stringifyXml(el)}</p:timing>`;
    if (rebuilt && sameTagMultiset(rebuilt, source)) return entries;
    return stringifyXml(el);
  },
};

/** Tag multiset of an XML string, keyed `<prefix:name` open tags. */
function tagMultiset(xml: string, into: Map<string, number>): void {
  for (const m of xml.matchAll(/<([\w-]+:[\w-]+)[ >/]/g)) {
    into.set(m[1]!, (into.get(m[1]!) ?? 0) + 1);
  }
}

function sameTagMultiset(rebuiltXml: string, sourceXml: string): boolean {
  const a = new Map<string, number>();
  const b = new Map<string, number>();
  tagMultiset(rebuiltXml, a);
  tagMultiset(sourceXml, b);
  if (a.size !== b.size) return false;
  for (const [tag, n] of a) {
    if (b.get(tag) !== n) return false;
  }
  return true;
}
