import { element as buildXml } from "@office-open/xml";
import { escapeXml } from "@office-open/xml";

import type {
  AnimationBuildOptions,
  AnimationClass,
  AnimationOptions,
  AnimationType,
  AnimationVariantOptions,
  EmphasisType,
  EndConditionOptions,
  PathAnimationType,
} from "./types";

// --- Preset ID mappings ---

const ENTR_PRESET_IDS: Record<AnimationType, number> = {
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
  zoom: 10,
  cover: 28,
  push: 19,
  strips: 23,
};

const EXIT_PRESET_IDS: Record<AnimationType, number> = {
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

const EMPH_PRESET_IDS: Record<EmphasisType, number> = {
  growShrink: 53,
  spin: 54,
  growWithTurn: 56,
  colorChange: 29,
  transparency: 57,
  boldFlash: 50,
  wave: 55,
  pulse: 58,
};

const PATH_PRESET_IDS: Record<PathAnimationType, number> = {
  customPath: 200,
  arc: 201,
  bounce: 202,
  circle: 203,
  curve: 204,
  figureEight: 205,
  line: 206,
  loop: 207,
};

const DIRECTION_SUBTYPES: Record<string, number> = {
  left: 4,
  right: 8,
  up: 2,
  down: 1,
  horizontal: 16,
  vertical: 32,
};

const FILTER_MAP: Record<string, string> = {
  fade: "fade",
  fly: "fly",
  wipe: "wipe",
  dissolve: "dissolve",
  split: "split",
  blinds: "blinds",
  checker: "checkerboard",
  randomBars: "randombars",
  wheel: "wheel",
  zoom: "zoom",
  cover: "cover",
  push: "push",
  strips: "strips",
};

const DIRECTION_FILTER: Record<string, Record<string, string>> = {
  fly: { left: "fromLeft", right: "fromRight", up: "fromTop", down: "fromBottom" },
  wipe: { left: "right", right: "left", up: "down", down: "up" },
  split: { horizontal: "horizontal", vertical: "vertical" },
  blinds: { horizontal: "horizontal", vertical: "vertical" },
  checker: { horizontal: "horizontal", vertical: "vertical" },
  randomBars: { horizontal: "horizontal", vertical: "vertical" },
  cover: { left: "left", right: "right", up: "top", down: "bottom" },
  push: { left: "left", right: "right", up: "top", down: "bottom" },
  strips: {
    left: "fromBottomLeft",
    right: "fromTopRight",
    up: "fromBottomRight",
    down: "fromTopLeft",
  },
  wheel: {},
};

const PATH_STRINGS: Record<PathAnimationType, string> = {
  customPath: "",
  arc: "M 0 0 C 50 -100 100 -100 150 0",
  bounce: "M 0 0 L 50 0 C 50 -30 30 -50 0 -50 C -30 -50 -50 -30 -50 0 L 50 0",
  circle: "M 0 0 C 50 -50 100 0 50 50 C 0 100 -50 50 0 0",
  curve: "M 0 0 C 50 -80 100 -80 150 0 C 200 80 250 80 300 0",
  figureEight:
    "M 0 0 C 50 -50 100 0 50 50 C 0 100 -50 50 0 0 C -50 -50 -100 0 -50 50 C 0 100 50 50 0 0",
  line: "M 0 0 L 200 0",
  loop: "M 0 0 C 50 -80 100 0 50 80 C 0 160 -50 80 0 0",
};

// --- Helper functions ---

function resolvePresetId(options: AnimationOptions): number {
  if (options.mediaType) return 1;
  if (options.attributeName) return 0;
  const cls: AnimationClass = options.class ?? "entr";

  if (cls === "emph" && options.emphasisType) {
    return EMPH_PRESET_IDS[options.emphasisType] ?? 0;
  }

  if (options.pathType) {
    return PATH_PRESET_IDS[options.pathType] ?? 200;
  }

  const type = options.type ?? "appear";
  const map = cls === "exit" ? EXIT_PRESET_IDS : ENTR_PRESET_IDS;
  return map[type] ?? 0;
}

function resolvePresetClass(options: AnimationOptions): AnimationClass {
  if (options.mediaType) return "mediacall";
  if (options.attributeName) return "entr";
  if (options.pathType) return "emph";
  return options.class ?? "entr";
}

function resolvePresetSubtype(options: AnimationOptions): number {
  if (options.direction) {
    return DIRECTION_SUBTYPES[options.direction] ?? 0;
  }
  return 0;
}

function buildTargetElement(spid: number, options?: AnimationOptions): string {
  // Target element type selection
  if (options?.inkTargetShapeId !== undefined) {
    return buildXml("p:tgtEl", undefined, [`<p:inkTgt spid="${options.inkTargetShapeId}"/>`]);
  }

  if (options?.soundTarget !== undefined) {
    return buildXml("p:tgtEl", undefined, [`<p:sndTgt r:id="${options.soundTarget}"/>`]);
  }

  const spTgtChildren: string[] = [];
  if (options?.charRange) {
    spTgtChildren.push(
      buildXml("p:txEl", undefined, [
        `<p:charRg st="${options.charRange[0]}" end="${options.charRange[1]}"/>`,
      ]),
    );
  } else if (options?.paragraphRange) {
    spTgtChildren.push(
      buildXml("p:txEl", undefined, [
        `<p:pRg st="${options.paragraphRange[0]}" end="${options.paragraphRange[1]}"/>`,
      ]),
    );
  } else if (options?.subShapeId !== undefined) {
    spTgtChildren.push(`<p:subSp spid="${options.subShapeId}"/>`);
  } else if (options?.graphicElementType !== undefined) {
    spTgtChildren.push(
      buildXml("p:graphicEl", undefined, [`<a:graphicEl type="${options.graphicElementType}"/>`]),
    );
  } else if (options?.oleChartElementType !== undefined) {
    const oleChartAttrs: Record<string, string | number> = {
      type: options.oleChartElementType,
    };
    if (options.oleChartElementLevel !== undefined) {
      oleChartAttrs.lvl = options.oleChartElementLevel;
    }
    spTgtChildren.push(buildXml("p:oleChartEl", oleChartAttrs));
  }

  return buildXml("p:tgtEl", undefined, [
    buildXml("p:spTgt", { spid }, spTgtChildren.length > 0 ? spTgtChildren : undefined),
  ]);
}

function buildEntrOrExitEffects(
  options: AnimationOptions,
  spid: number,
  ids: { set: number; effect: number },
): string[] {
  const children: string[] = [];

  // p:set — make shape visible (entr) or hidden (exit)
  const cls = options.class ?? "entr";
  const visibility = cls === "exit" ? "hidden" : "visible";
  children.push(
    buildXml("p:set", undefined, [
      buildXml("p:cBhvr", undefined, [
        buildXml("p:cTn", { id: ids.set, dur: "1", fill: "hold" }, [
          buildXml("p:stCondLst", undefined, [`<p:cond delay="0"/>`]),
        ]),
        buildXml("p:tgtEl", undefined, [buildXml("p:spTgt", { spid })]),
        buildXml("p:attrNameLst", undefined, [`<p:attrName>style.visibility</p:attrName>`]),
      ]),
      buildXml("p:to", undefined, [`<p:strVal val="${visibility}"/>`]),
    ]),
  );

  // p:animEffect for non-appear types
  if (options.type !== "appear") {
    const filterBase = FILTER_MAP[options.type ?? "fade"];
    const dirMap = DIRECTION_FILTER[options.type ?? "fade"];
    const dirFilter = options.direction && dirMap ? dirMap[options.direction] : undefined;

    let filter = filterBase;
    if (options.type === "wheel") {
      filter = "wheel(4)";
    } else if (options.type === "split") {
      const dir = options.direction ?? "horizontal";
      filter = `split(${dir})`;
    } else if (dirFilter) {
      filter = `${filterBase}(${dirFilter})`;
    }

    const transition = cls === "exit" ? "out" : "in";

    const animEffectAttrs: Record<string, string | undefined> = {
      transition,
      filter,
    };
    if (options.propertyList) animEffectAttrs.prLst = options.propertyList;

    const animEffectChildren: string[] = [
      buildXml("p:cBhvr", undefined, [
        buildXml("p:cTn", { id: ids.effect, dur: String(options.duration ?? 500), fill: "hold" }),
        buildXml("p:tgtEl", undefined, [buildXml("p:spTgt", { spid })]),
      ]),
    ];

    // Add progress if specified (CT_TLAnimateEffectBehavior progress)
    if (options.effectProgress) {
      animEffectChildren.push(
        buildXml("p:progress", undefined, [buildVariantValue(options.effectProgress)]),
      );
    }

    children.push(buildXml("p:animEffect", animEffectAttrs, animEffectChildren));
  }

  return children;
}

function buildEmphasisEffects(
  options: AnimationOptions,
  spid: number,
  ids: { set: number; effect: number },
): string[] {
  const emphType = options.emphasisType ?? "growShrink";
  const children: string[] = [];
  const dur = String(options.duration ?? 500);

  // Common target element
  const tgtEl = buildXml("p:tgtEl", undefined, [buildXml("p:spTgt", { spid })]);

  switch (emphType) {
    case "growShrink": {
      const animScaleAttrs: Record<string, string | number | undefined> = {};
      if (options.zoomContents !== undefined)
        animScaleAttrs.zoomContents = options.zoomContents ? 1 : 0;
      children.push(
        buildXml("p:animScale", animScaleAttrs, [
          buildXml("p:cBhvr", undefined, [
            buildXml(
              "p:cTn",
              {
                id: ids.effect,
                dur,
                fill: "hold",
                ...(options.autoReverse ? { autoRev: 1 } : {}),
              },
              [buildXml("p:stCondLst", undefined, [`<p:cond delay="0"/>`])],
            ),
            tgtEl,
          ]),
          `<p:by x="150000" y="150000"/>`,
        ]),
      );
      break;
    }

    case "spin":
      children.push(
        buildXml("p:animRot", { by: "3600000" }, [
          buildXml("p:cBhvr", undefined, [
            buildXml(
              "p:cTn",
              {
                id: ids.effect,
                dur,
                fill: "hold",
              },
              [buildXml("p:stCondLst", undefined, [`<p:cond delay="0"/>`])],
            ),
            tgtEl,
          ]),
        ]),
      );
      break;

    case "colorChange": {
      const color = options.color ?? "FF0000";
      const animClrChildren: string[] = [
        buildXml("p:cBhvr", undefined, [
          buildXml(
            "p:cTn",
            {
              id: ids.effect,
              dur,
              fill: "hold",
            },
            [buildXml("p:stCondLst", undefined, [`<p:cond delay="0"/>`])],
          ),
          tgtEl,
        ]),
      ];
      // p:by (color animation by value — rgb or hsl)
      if (options.colorByRgb) {
        animClrChildren.push(
          buildXml("p:by", undefined, [
            `<p:rgb r="${options.colorByRgb.r}" g="${options.colorByRgb.g}" b="${options.colorByRgb.b}"/>`,
          ]),
        );
      } else if (options.colorByHsl) {
        animClrChildren.push(
          buildXml("p:by", undefined, [
            `<p:hsl h="${options.colorByHsl.h}" s="${options.colorByHsl.s}" l="${options.colorByHsl.l}"/>`,
          ]),
        );
      }
      // p:from (color animation from value)
      if (options.colorFrom) {
        animClrChildren.push(
          buildXml("p:from", undefined, [`<a:srgbClr val="${options.colorFrom}"/>`]),
        );
      }
      // p:to (color animation to value)
      animClrChildren.push(
        buildXml("p:to", undefined, [`<a:srgbClr val="${options.colorTo ?? color}"/>`]),
      );
      children.push(buildXml("p:animClr", undefined, animClrChildren));
      break;
    }

    case "transparency":
      children.push(
        buildXml(
          "p:anim",
          {
            calcmode: "lin",
            valueType: "num",
            from: "0",
            to: "1",
          },
          [
            buildXml("p:cBhvr", undefined, [
              buildXml("p:cTn", {
                id: ids.effect,
                dur,
                fill: "hold",
              }),
              tgtEl,
              buildXml("p:attrNameLst", undefined, [`<p:attrName>style.opacity</p:attrName>`]),
            ]),
          ],
        ),
      );
      break;

    default:
      // boldFlash, wave, pulse, growWithTurn — use p:animEffect with emph filter
      children.push(
        buildXml(
          "p:animEffect",
          {
            transition: "in",
            filter: emphType,
          },
          [
            buildXml("p:cBhvr", undefined, [
              buildXml("p:cTn", {
                id: ids.effect,
                dur,
                fill: "hold",
              }),
              tgtEl,
            ]),
          ],
        ),
      );
      break;
  }

  return children;
}

function buildPathEffects(
  options: AnimationOptions,
  spid: number,
  ids: { set: number; effect: number },
): string[] {
  const pathStr = options.path ?? PATH_STRINGS[options.pathType ?? "customPath"] ?? "";
  const dur = String(options.duration ?? 1000);

  const animMotionAttrs: Record<string, string | undefined> = {
    origin: "layout",
    path: pathStr,
  };
  if (options.pointsTypes) animMotionAttrs.ptsTypes = options.pointsTypes;
  if (options.pathEditMode) animMotionAttrs.pathEditMode = options.pathEditMode;
  if (options.rotationAngle !== undefined) animMotionAttrs.rAng = String(options.rotationAngle);

  const animMotionChildren: string[] = [
    buildXml("p:cBhvr", undefined, [
      buildXml(
        "p:cTn",
        {
          id: ids.effect,
          dur,
          fill: "hold",
        },
        [buildXml("p:stCondLst", undefined, [`<p:cond delay="0"/>`])],
      ),
      buildTargetElement(spid, options),
    ]),
  ];

  // Motion path from/rCtr (A5)
  if (options.motionFrom) {
    animMotionChildren.push(`<p:from x="${options.motionFrom.x}" y="${options.motionFrom.y}"/>`);
  }
  if (options.motionRotationCenter) {
    animMotionChildren.push(
      `<p:rCtr x="${options.motionRotationCenter.x}" y="${options.motionRotationCenter.y}"/>`,
    );
  }

  return [buildXml("p:animMotion", animMotionAttrs, animMotionChildren)];
}

/**
 * Build p:cmd for media play/pause/stop (goes inside the animation sequence).
 * PowerPoint uses p:cmd type="call" cmd="playFrom(0.0)" for video play.
 */
function buildMediaPlayCommand(
  options: AnimationOptions,
  spid: number,
  ids: { cmd: number },
): string {
  const cmdStr = options.fullScreen ? "playFrom(0.0,1.0)" : "playFrom(0.0)";

  return buildXml("p:cmd", { type: "call", cmd: cmdStr }, [
    buildXml("p:cBhvr", undefined, [
      buildXml("p:cTn", {
        id: ids.cmd,
        dur: String(options.duration ?? 10000),
        fill: "hold",
      }),
      buildTargetElement(spid),
    ]),
  ]);
}

/**
 * Build generic p:cmd with configurable type (call/evt/verb).
 */
function buildCommand(options: AnimationOptions, spid: number, ids: { cmd: number }): string {
  const cmdType = options.commandType ?? "call";
  const cmdStr = options.command ?? "";

  return buildXml("p:cmd", { type: cmdType, cmd: cmdStr }, [
    buildXml("p:cBhvr", undefined, [
      buildXml("p:cTn", {
        id: ids.cmd,
        dur: String(options.duration ?? 1000),
        fill: "hold",
      }),
      buildTargetElement(spid),
    ]),
  ]);
}

/**
 * Build p:set for an instant property change (setBehavior).
 */
function buildSetBehavior(
  setOptions: NonNullable<AnimationOptions["setBehavior"]>,
  spid: number,
  ids: { set: number },
): string {
  const toValName = setOptions.toType === "number" ? "p:numVal" : "p:strVal";

  return buildXml("p:set", undefined, [
    buildXml("p:cBhvr", undefined, [
      buildXml("p:cTn", { id: ids.set, dur: "1", fill: "hold" }, [
        buildXml("p:stCondLst", undefined, [`<p:cond delay="0"/>`]),
      ]),
      buildTargetElement(spid),
      buildXml("p:attrNameLst", undefined, [
        `<p:attrName>${escapeXml(setOptions.attributeName)}</p:attrName>`,
      ]),
    ]),
    buildXml("p:to", undefined, [`<${toValName} val="${setOptions.toValue}"/>`]),
  ]);
}

/**
 * Build p:iterate for text-level animation iteration.
 */
function buildIterate(iterate: NonNullable<AnimationOptions["iterate"]>): string {
  const attrs: Record<string, string | number | undefined> = {};
  if (iterate.type) attrs.type = iterate.type;
  if (iterate.backwards) attrs.backwards = 1;

  const iterChildren: string[] = [];
  if (iterate.iteratePercentage !== undefined) {
    iterChildren.push(`<p:tmPct val="${iterate.iteratePercentage}"/>`);
  } else if (iterate.interval !== undefined) {
    iterChildren.push(`<p:tmAbs val="${iterate.interval}"/>`);
  }

  return buildXml("p:iterate", attrs, iterChildren.length > 0 ? iterChildren : undefined);
}

/**
 * Build p:animVariant value child (boolVal|intVal|fltVal|strVal|clrVal).
 */
function buildVariantValue(variant: AnimationVariantOptions): string {
  if (variant.bool !== undefined) {
    return `<p:boolVal val="${variant.bool ? 1 : 0}"/>`;
  }
  if (variant.int !== undefined) {
    return `<p:intVal val="${variant.int}"/>`;
  }
  if (variant.float !== undefined) {
    return `<p:fltVal val="${variant.float}"/>`;
  }
  if (variant.color !== undefined) {
    return buildXml("p:clrVal", undefined, [`<a:srgbClr val="${variant.color}"/>`]);
  }
  // string fallback
  return `<p:strVal val="${variant.string ?? ""}"/>`;
}

/**
 * Build a condition element (p:cond) from EndConditionOptions.
 */
function buildCondition(cond: EndConditionOptions): string {
  const attrs: Record<string, string | number | undefined> = {};
  if (cond.delay !== undefined) attrs.delay = cond.delay;
  if (cond.event !== undefined) attrs.evt = cond.event;

  const children: string[] = [];
  if (cond.timeNodeId !== undefined) {
    children.push(`<p:tn val="${cond.timeNodeId}"/>`);
  } else if (cond.runtimeNode !== undefined) {
    children.push(`<p:rtn val="${cond.runtimeNode}"/>`);
  }

  return buildXml("p:cond", attrs, children.length > 0 ? children : undefined);
}

/**
 * Build p:video/p:audio with p:cMediaNode (goes as sibling of p:seq).
 * This is the media state controller, separate from the play command.
 */
function buildMediaStateNode(
  options: AnimationOptions,
  spid: number,
  ids: { mediaCtn: number },
): string {
  const isVideo = options.mediaType === "playVideo";
  const elementName = isVideo ? "p:video" : "p:audio";

  const mediaAttrs: Record<string, string | number | boolean | undefined> = {};
  if (!isVideo && options.isNarration) {
    mediaAttrs.isNarration = true;
  }
  if (isVideo && options.fullScreen) {
    mediaAttrs.fullScrn = true;
  }

  const cMediaNodeAttrs: Record<string, string | number | undefined> = {};
  if (options.volume !== undefined) {
    cMediaNodeAttrs.vol = options.volume * 1000;
  } else {
    cMediaNodeAttrs.vol = 80000;
  }
  if (options.mute) {
    cMediaNodeAttrs.mute = 1;
  }
  if (options.numberOfSlides !== undefined) {
    cMediaNodeAttrs.numSld = options.numberOfSlides;
  }
  if (options.showWhenStopped !== undefined) {
    cMediaNodeAttrs.showWhenStopped = options.showWhenStopped ? 1 : 0;
  }

  return buildXml(elementName, mediaAttrs, [
    buildXml("p:cMediaNode", cMediaNodeAttrs, [
      buildXml(
        "p:cTn",
        {
          id: ids.mediaCtn,
          fill: "hold",
          display: "0",
        },
        [buildXml("p:stCondLst", undefined, [`<p:cond delay="indefinite"/>`])],
      ),
      buildTargetElement(spid),
    ]),
  ]);
}

function buildPropertyAnimation(
  options: AnimationOptions,
  spid: number,
  ids: { cBhvr: number },
): string {
  const attrs: Record<string, string | undefined> = {};
  if (options.calcMode) attrs.calcmode = options.calcMode;
  if (options.valueType) attrs.valueType = options.valueType;
  if (options.from !== undefined) attrs.from = options.from;
  if (options.to !== undefined) attrs.to = options.to;
  if (options.animBy !== undefined) attrs.by = options.animBy;
  if (options.formula !== undefined) attrs.fmla = options.formula;
  if (options.colorSpace !== undefined) attrs.clrSpc = options.colorSpace;

  const cBhvrChildren: string[] = [
    buildXml("p:cTn", {
      id: ids.cBhvr,
      dur: String(options.duration ?? 500),
      fill: "hold",
    }),
    buildTargetElement(spid, options),
  ];

  if (options.attributeName) {
    cBhvrChildren.push(
      buildXml("p:attrNameLst", undefined, [
        `<p:attrName>${escapeXml(options.attributeName)}</p:attrName>`,
      ]),
    );
  }

  const cBhvrAttrs: Record<string, string | number | undefined> = {};
  if (options.additive !== undefined) cBhvrAttrs.additive = options.additive;
  if (options.accumulate !== undefined) cBhvrAttrs.accumulate = options.accumulate;
  if (options.transformType !== undefined) cBhvrAttrs.xfrmType = options.transformType;
  if (options.runtimeContext !== undefined) cBhvrAttrs.rctx = options.runtimeContext;
  if (options.override !== undefined) cBhvrAttrs.override = options.override;

  const animChildren: string[] = [buildXml("p:cBhvr", cBhvrAttrs, cBhvrChildren)];

  // Build tavLst if from/to are specified (support variant value types)
  if (
    options.from !== undefined ||
    options.to !== undefined ||
    options.colorFrom ||
    options.colorTo
  ) {
    const tavList: string[] = [];

    // Build variant child for p:val
    const buildVariantChild = (val: string, isColor?: boolean): string => {
      if (isColor) {
        return `<p:clrVal><a:srgbClr val="${val}"/></p:clrVal>`;
      }
      if (options.variantInt !== undefined) {
        return `<p:intVal val="${Number(val)}"/>`;
      }
      if (options.variantFloat !== undefined) {
        return `<p:fltVal val="${Number(val)}"/>`;
      }
      if (options.variantBool !== undefined) {
        return `<p:boolVal val="${val === "true" ? 1 : 0}"/>`;
      }
      return `<p:strVal val="${val}"/>`;
    };

    if (options.colorFrom) {
      tavList.push(
        buildXml("p:tav", { tm: "0" }, [
          buildXml("p:val", undefined, [buildVariantChild(options.colorFrom, true)]),
        ]),
      );
    } else if (options.from !== undefined) {
      tavList.push(
        buildXml("p:tav", { tm: "0" }, [
          buildXml("p:val", undefined, [buildVariantChild(options.from)]),
        ]),
      );
    }
    if (options.colorTo) {
      tavList.push(
        buildXml("p:tav", { tm: "100000" }, [
          buildXml("p:val", undefined, [buildVariantChild(options.colorTo, true)]),
        ]),
      );
    } else if (options.to !== undefined) {
      tavList.push(
        buildXml("p:tav", { tm: "100000" }, [
          buildXml("p:val", undefined, [buildVariantChild(options.to)]),
        ]),
      );
    }
    if (tavList.length > 0) {
      animChildren.push(buildXml("p:tavLst", undefined, tavList));
    }
  }

  return buildXml("p:anim", attrs, animChildren);
}

// --- Build list (bldLst) ---

function buildBuildList(builds: AnimationBuildOptions[], nextId: () => number): string {
  const bldChildren: string[] = [];

  for (const bld of builds) {
    const bldAttrs: Record<string, string | number | undefined> = {
      spid: bld.spid,
      grpId: bld.grpId,
    };
    if (bld.uiExpand) bldAttrs.uiExpand = 1;

    let elementName: string;
    const bldChildrenInner: string[] = [];

    switch (bld.type) {
      case "paragraph": {
        elementName = "p:bldP";
        if (bld.build) bldAttrs.build = bld.build;
        if (bld.bldLvl !== undefined) bldAttrs.bldLvl = bld.bldLvl;
        if (bld.animBg !== undefined) bldAttrs.animBg = bld.animBg ? 1 : 0;
        if (bld.autoUpdateAnimBg !== undefined)
          bldAttrs.autoUpdateAnimBg = bld.autoUpdateAnimBg ? 1 : 0;
        if (bld.rev) bldAttrs.rev = 1;
        if (bld.advAuto !== undefined) bldAttrs.advAuto = bld.advAuto;
        // Templates (tmplLst)
        if (bld.templates && bld.templates.length > 0) {
          const tmplChildren = bld.templates.map((tmpl) => {
            const tmplAttrs: Record<string, string | number | undefined> = {};
            if (tmpl.lvl !== undefined) tmplAttrs.lvl = tmpl.lvl;
            const tnLstChildren = tmpl.children.map(() => {
              const tid = nextId();
              return buildXml("p:par", undefined, [`<p:cTn id="${tid}"/>`]);
            });
            return buildXml("p:tmpl", tmplAttrs, [buildXml("p:tnLst", undefined, tnLstChildren)]);
          });
          bldChildrenInner.push(buildXml("p:tmplLst", undefined, tmplChildren));
        }
        break;
      }
      case "diagram": {
        elementName = "p:bldDgm";
        if (bld.diagramBuild) bldAttrs.bld = bld.diagramBuild;
        break;
      }
      case "oleChart": {
        elementName = "p:bldOleChart";
        if (bld.oleChartBuild) bldAttrs.bld = bld.oleChartBuild;
        if (bld.oleChartAnimBg !== undefined) bldAttrs.animBg = bld.oleChartAnimBg ? 1 : 0;
        break;
      }
      case "graphic": {
        elementName = "p:bldGraphic";
        if (bld.graphicBuildAsOne) {
          bldChildrenInner.push(`<p:bldAsOne/>`);
        } else {
          bldChildrenInner.push(`<p:bldSub/>`);
        }
        break;
      }
    }

    bldChildren.push(
      buildXml(elementName, bldAttrs, bldChildrenInner.length > 0 ? bldChildrenInner : undefined),
    );
  }

  return buildXml("p:bldLst", undefined, bldChildren);
}

// --- Main class ---

export interface AnimationEntry {
  spid: number;
  options: AnimationOptions;
}

/**
 * p:timing — Slide timing for shape animations.
 *
 * Pure collector — no XmlComponent inheritance.
 * Collects string children and serializes via `toXml()`.
 */
export class SlideTiming {
  private parts: string[] = [];

  public constructor(entries: AnimationEntry[]) {
    if (entries.length === 0) return;

    let id = 1;
    const rootCtnId = id++;
    const seqCtnId = id++;

    // Check for builds from first entry
    const builds = entries[0]?.options.builds;
    const previousAction = entries[0]?.options.previousAction;

    const animationNodes: string[] = [];
    const mediaStateNodes: string[] = [];
    let clickGroupDelay = 0;

    for (let i = 0; i < entries.length; i++) {
      const entry = entries[i];
      const { spid, options } = entry;

      const nodeType =
        options.trigger === "withPrevious"
          ? "withEffect"
          : options.trigger === "afterPrevious"
            ? "afterEffect"
            : "clickEffect";

      if (options.trigger === "afterPrevious" && i > 0) {
        clickGroupDelay += (options.duration ?? 500) + (options.delay ?? 0);
      }
      if (options.trigger === "onClick" || options.trigger === undefined) {
        clickGroupDelay = 0;
      }

      const groupCtnId = id++;
      const effectCtnId = id++;
      const setCtnId = id++;
      const effectCtnId2 = id++;

      const presetId = resolvePresetId(options);
      const presetClass = resolvePresetClass(options);
      const presetSubtype = resolvePresetSubtype(options);

      // Build effect children based on animation class
      let effectChildren: string[];

      if (options.commandType) {
        effectChildren = [buildCommand(options, spid, { cmd: effectCtnId2 })];
      } else if (options.mediaType) {
        effectChildren = [buildMediaPlayCommand(options, spid, { cmd: effectCtnId2 })];
        // Generate media state node (p:video/p:audio) as sibling of p:seq
        const mediaStateId = id++;
        mediaStateNodes.push(buildMediaStateNode(options, spid, { mediaCtn: mediaStateId }));
      } else if (options.attributeName) {
        effectChildren = [buildPropertyAnimation(options, spid, { cBhvr: effectCtnId2 })];
      } else if (options.pathType) {
        effectChildren = buildPathEffects(options, spid, {
          set: setCtnId,
          effect: effectCtnId2,
        });
      } else if (presetClass === "emph") {
        effectChildren = buildEmphasisEffects(options, spid, {
          set: setCtnId,
          effect: effectCtnId2,
        });
      } else {
        effectChildren = buildEntrOrExitEffects(options, spid, {
          set: setCtnId,
          effect: effectCtnId2,
        });
      }

      // Prepend a p:set for setBehavior (if specified)
      if (options.setBehavior) {
        const setBehId = id++;
        effectChildren = [
          buildSetBehavior(options.setBehavior, spid, { set: setBehId }),
          ...effectChildren,
        ];
      }

      // Build cTn attributes with optional speed/repeatCount/autoReverse
      const cTnAttrs: Record<string, string | number | undefined> = {
        id: effectCtnId,
        fill: "hold",
        nodeType,
      };
      if (!options.mediaType && !options.attributeName) {
        cTnAttrs.presetID = presetId;
        cTnAttrs.presetClass = presetClass;
        cTnAttrs.presetSubtype = presetSubtype;
      }
      if (options.mediaType) {
        cTnAttrs.presetID = presetId;
        cTnAttrs.presetClass = presetClass;
        cTnAttrs.presetSubtype = presetSubtype;
      }
      if (options.speed !== undefined) cTnAttrs.spd = String(options.speed);
      if (options.repeatCount !== undefined) cTnAttrs.repeatCount = String(options.repeatCount);
      if (options.autoReverse) cTnAttrs.autoRev = 1;
      if (options.repeatDuration !== undefined) cTnAttrs.repeatDur = options.repeatDuration;
      if (options.acceleration !== undefined) cTnAttrs.accel = options.acceleration;
      if (options.deceleration !== undefined) cTnAttrs.decel = options.deceleration;
      if (options.restart !== undefined) cTnAttrs.restart = options.restart;
      if (options.syncBehavior !== undefined) cTnAttrs.syncBehavior = options.syncBehavior;
      if (options.timeFilter !== undefined) cTnAttrs.tmFilter = options.timeFilter;
      if (options.eventFilter !== undefined) cTnAttrs.evtFilter = options.eventFilter;
      if (options.display !== undefined) cTnAttrs.display = options.display ? 1 : 0;
      if (options.masterRelation !== undefined) cTnAttrs.masterRel = options.masterRelation;
      if (options.buildLevel !== undefined) cTnAttrs.bldLvl = options.buildLevel;
      if (options.groupId !== undefined) cTnAttrs.grpId = options.groupId;
      if (options.afterEffect !== undefined) cTnAttrs.afterEffect = options.afterEffect ? 1 : 0;
      if (options.nodePlaceholder !== undefined) cTnAttrs.nodePh = options.nodePlaceholder ? 1 : 0;
      if (options.advanceAfterTime !== undefined) cTnAttrs.advAuto = options.advanceAfterTime;
      if (options.animateBackground !== undefined)
        cTnAttrs.animBg = options.animateBackground ? 1 : 0;
      if (options.autoUpdateAnimationBackground !== undefined)
        cTnAttrs.autoUpdateAnimBg = options.autoUpdateAnimationBackground ? 1 : 0;

      // Build effect cTn
      const effectCtnChildren: string[] = [
        buildXml("p:stCondLst", undefined, [
          buildXml("p:cond", { delay: String(options.delay ?? 0) }),
        ]),
      ];

      // Add iterate container if specified
      if (options.iterate) {
        effectCtnChildren.push(buildIterate(options.iterate));
      }

      // Add endCondLst (A2)
      if (options.endConditions && options.endConditions.length > 0) {
        effectCtnChildren.push(
          buildXml("p:endCondLst", undefined, options.endConditions.map(buildCondition)),
        );
      }

      // Add endSync (A2) — directly CT_TLTimeCondition, no cond wrapper
      if (options.endSyncCondition) {
        const syncAttrs: Record<string, string | undefined> = {};
        if (options.endSyncCondition.delay !== undefined)
          syncAttrs.delay = options.endSyncCondition.delay;
        if (options.endSyncCondition.event !== undefined)
          syncAttrs.evt = options.endSyncCondition.event;

        const syncChildren: string[] = [];
        if (options.endSyncCondition.timeNodeId !== undefined) {
          syncChildren.push(`<p:tn val="${options.endSyncCondition.timeNodeId}"/>`);
        } else if (options.endSyncCondition.runtimeNode !== undefined) {
          syncChildren.push(`<p:rtn val="${options.endSyncCondition.runtimeNode}"/>`);
        }

        effectCtnChildren.push(
          buildXml("p:endSync", syncAttrs, syncChildren.length > 0 ? syncChildren : undefined),
        );
      }

      // Add childTnLst with possible excl wrapper
      const childTnListChildren = options.exclusiveMode
        ? [
            buildXml("p:excl", undefined, [
              buildXml("p:cTn", { id: id++ }, [
                buildXml("p:childTnLst", undefined, effectChildren),
              ]),
            ]),
          ]
        : effectChildren;

      effectCtnChildren.push(buildXml("p:childTnLst", undefined, childTnListChildren));

      // Add subTnLst (A2)
      if (options.subTimeNodes && options.subTimeNodes.length > 0) {
        effectCtnChildren.push(
          buildXml(
            "p:subTnLst",
            undefined,
            options.subTimeNodes.map((subOpts) => {
              const subId = id++;
              return buildXml("p:par", undefined, [
                buildXml(
                  "p:cTn",
                  {
                    id: subId,
                    dur: String(subOpts.duration ?? 0),
                  },
                  [
                    buildXml("p:stCondLst", undefined, [
                      buildXml("p:cond", { delay: String(subOpts.delay ?? 0) }),
                    ]),
                  ],
                ),
              ]);
            }),
          ),
        );
      }

      const effectCtn = buildXml("p:cTn", cTnAttrs, effectCtnChildren);

      // Wrap in inner par
      const innerPar = buildXml("p:par", undefined, [effectCtn]);

      // Wrap in outer par (click group)
      const outerPar = buildXml("p:par", undefined, [
        buildXml(
          "p:cTn",
          {
            id: groupCtnId,
            fill: "hold",
          },
          [
            buildXml("p:stCondLst", undefined, [`<p:cond delay="0"/>`]),
            buildXml("p:childTnLst", undefined, [innerPar]),
          ],
        ),
      ]);

      animationNodes.push(outerPar);
    }

    // Assemble the full timing tree
    const seqAttrs: Record<string, string | number | undefined> = {
      concurrent: 1,
      nextAc: "seek",
    };
    if (previousAction !== undefined) seqAttrs.prevAc = previousAction;

    this.parts.push(
      buildXml("p:tnLst", undefined, [
        buildXml("p:par", undefined, [
          buildXml(
            "p:cTn",
            {
              id: rootCtnId,
              dur: "indefinite",
              restart: "never",
              nodeType: "tmRoot",
            },
            [
              buildXml("p:childTnLst", undefined, [
                buildXml("p:seq", seqAttrs, [
                  buildXml(
                    "p:cTn",
                    {
                      id: seqCtnId,
                      dur: "indefinite",
                      nodeType: "mainSeq",
                    },
                    [buildXml("p:childTnLst", undefined, animationNodes)],
                  ),
                  buildXml("p:prevCondLst", undefined, [
                    buildXml("p:cond", { evt: "onPrev", delay: "0" }, [
                      buildXml("p:tgtEl", undefined, [`<p:sldTgt/>`]),
                    ]),
                  ]),
                  buildXml("p:nextCondLst", undefined, [
                    buildXml("p:cond", { evt: "onNext", delay: "0" }, [
                      buildXml("p:tgtEl", undefined, [`<p:sldTgt/>`]),
                    ]),
                  ]),
                ]),
                ...mediaStateNodes,
              ]),
            ],
          ),
        ]),
      ]),
    );

    // bldLst — sibling of tnLst in p:timing (CT_SlideTiming: tnLst → bldLst → extLst)
    if (builds) {
      this.parts.push(buildBuildList(builds, () => id++));
    }
  }

  /** Serialize to p:timing XML string. */
  public toXml(): string {
    if (this.parts.length === 0) return "";
    const body = this.parts.join("");
    return body ? `<p:timing>${body}</p:timing>` : "";
  }
}
