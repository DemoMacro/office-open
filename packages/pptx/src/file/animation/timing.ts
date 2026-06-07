import { BuilderElement, XmlComponent, stringContainerObj } from "@file/xml-components";

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

function buildTargetElement(spid: number, options?: AnimationOptions): XmlComponent {
  // Target element type selection
  if (options?.inkTargetShapeId !== undefined) {
    return new BuilderElement({
      name: "p:tgtEl",
      children: [
        new BuilderElement({
          name: "p:inkTgt",
          attributes: { spid: { key: "spid", value: options.inkTargetShapeId } },
        }),
      ],
    });
  }

  if (options?.soundTarget !== undefined) {
    return new BuilderElement({
      name: "p:tgtEl",
      children: [
        new BuilderElement({
          name: "p:sndTgt",
          attributes: { "r:id": { key: "r:id", value: options.soundTarget } },
        }),
      ],
    });
  }

  const spTgtChildren: XmlComponent[] = [];
  if (options?.charRange) {
    spTgtChildren.push(
      new BuilderElement({
        name: "p:txEl",
        children: [
          new BuilderElement({
            name: "p:charRg",
            attributes: {
              st: { key: "st", value: options.charRange[0] },
              end: { key: "end", value: options.charRange[1] },
            },
          }),
        ],
      }),
    );
  } else if (options?.paragraphRange) {
    spTgtChildren.push(
      new BuilderElement({
        name: "p:txEl",
        children: [
          new BuilderElement({
            name: "p:pRg",
            attributes: {
              st: { key: "st", value: options.paragraphRange[0] },
              end: { key: "end", value: options.paragraphRange[1] },
            },
          }),
        ],
      }),
    );
  } else if (options?.subShapeId !== undefined) {
    spTgtChildren.push(
      new BuilderElement({
        name: "p:subSp",
        attributes: { spid: { key: "spid", value: options.subShapeId } },
      }),
    );
  } else if (options?.graphicElementType !== undefined) {
    spTgtChildren.push(
      new BuilderElement({
        name: "p:graphicEl",
        children: [
          new BuilderElement({
            name: "a:graphicEl",
            attributes: { type: { key: "type", value: options.graphicElementType } },
          }),
        ],
      }),
    );
  } else if (options?.oleChartElementType !== undefined) {
    const oleChartAttrs: Record<string, { key: string; value: string | number }> = {
      type: { key: "type", value: options.oleChartElementType },
    };
    if (options.oleChartElementLevel !== undefined) {
      oleChartAttrs.lvl = { key: "lvl", value: options.oleChartElementLevel };
    }
    spTgtChildren.push(
      new BuilderElement({
        name: "p:oleChartEl",
        attributes: oleChartAttrs,
      }),
    );
  }

  return new BuilderElement({
    name: "p:tgtEl",
    children: [
      new BuilderElement({
        name: "p:spTgt",
        attributes: { spid: { key: "spid", value: spid } },
        children: spTgtChildren.length > 0 ? spTgtChildren : undefined,
      }),
    ],
  });
}

function buildEntrOrExitEffects(
  options: AnimationOptions,
  spid: number,
  ids: { set: number; effect: number },
): XmlComponent[] {
  const children: XmlComponent[] = [];

  // p:set — make shape visible (entr) or hidden (exit)
  const cls = options.class ?? "entr";
  const visibility = cls === "exit" ? "hidden" : "visible";
  children.push(
    new BuilderElement({
      name: "p:set",
      children: [
        new BuilderElement({
          name: "p:cBhvr",
          children: [
            new BuilderElement({
              name: "p:cTn",
              attributes: {
                id: { key: "id", value: ids.set },
                dur: { key: "dur", value: "1" },
                fill: { key: "fill", value: "hold" },
              },
              children: [
                new BuilderElement({
                  name: "p:stCondLst",
                  children: [
                    new BuilderElement({
                      name: "p:cond",
                      attributes: { delay: { key: "delay", value: "0" } },
                    }),
                  ],
                }),
              ],
            }),
            new BuilderElement({
              name: "p:tgtEl",
              children: [
                new BuilderElement({
                  name: "p:spTgt",
                  attributes: { spid: { key: "spid", value: spid } },
                }),
              ],
            }),
            new BuilderElement({
              name: "p:attrNameLst",
              children: [stringContainerObj("p:attrName", "style.visibility")],
            }),
          ],
        }),
        new BuilderElement({
          name: "p:to",
          children: [
            new BuilderElement({
              name: "p:strVal",
              attributes: { val: { key: "val", value: visibility } },
            }),
          ],
        }),
      ],
    }),
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

    const animEffectAttrs: Record<string, { key: string; value: string }> = {
      transition: { key: "transition", value: transition },
      filter: { key: "filter", value: filter },
    };
    if (options.propertyList) animEffectAttrs.prLst = { key: "prLst", value: options.propertyList };

    const animEffectChildren: XmlComponent[] = [
      new BuilderElement({
        name: "p:cBhvr",
        children: [
          new BuilderElement({
            name: "p:cTn",
            attributes: {
              id: { key: "id", value: ids.effect },
              dur: { key: "dur", value: String(options.duration ?? 500) },
              fill: { key: "fill", value: "hold" },
            },
          }),
          new BuilderElement({
            name: "p:tgtEl",
            children: [
              new BuilderElement({
                name: "p:spTgt",
                attributes: { spid: { key: "spid", value: spid } },
              }),
            ],
          }),
        ],
      }),
    ];

    // Add progress if specified (CT_TLAnimateEffectBehavior progress)
    if (options.effectProgress) {
      animEffectChildren.push(
        new BuilderElement({
          name: "p:progress",
          children: [buildVariantValue(options.effectProgress)],
        }),
      );
    }

    children.push(
      new BuilderElement({
        name: "p:animEffect",
        attributes: animEffectAttrs,
        children: animEffectChildren,
      }),
    );
  }

  return children;
}

function buildEmphasisEffects(
  options: AnimationOptions,
  spid: number,
  ids: { set: number; effect: number },
): XmlComponent[] {
  const emphType = options.emphasisType ?? "growShrink";
  const children: XmlComponent[] = [];
  const dur = String(options.duration ?? 500);

  // Common target element
  const tgtEl = new BuilderElement({
    name: "p:tgtEl",
    children: [
      new BuilderElement({
        name: "p:spTgt",
        attributes: { spid: { key: "spid", value: spid } },
      }),
    ],
  });

  switch (emphType) {
    case "growShrink":
      children.push(
        new BuilderElement({
          name: "p:animScale",
          children: [
            new BuilderElement({
              name: "p:cBhvr",
              children: [
                new BuilderElement({
                  name: "p:cTn",
                  attributes: {
                    id: { key: "id", value: ids.effect },
                    dur: { key: "dur", value: dur },
                    fill: { key: "fill", value: "hold" },
                    ...(options.autoReverse ? { autoRev: { key: "autoRev", value: 1 } } : {}),
                  },
                  children: [
                    new BuilderElement({
                      name: "p:stCondLst",
                      children: [
                        new BuilderElement({
                          name: "p:cond",
                          attributes: {
                            delay: { key: "delay", value: "0" },
                          },
                        }),
                      ],
                    }),
                  ],
                }),
                tgtEl,
              ],
            }),
            new BuilderElement({
              name: "p:by",
              attributes: {
                x: { key: "x", value: "150000" },
                y: { key: "y", value: "150000" },
              },
            }),
          ],
        }),
      );
      break;

    case "spin":
      children.push(
        new BuilderElement({
          name: "p:animRot",
          attributes: { by: { key: "by", value: "3600000" } },
          children: [
            new BuilderElement({
              name: "p:cBhvr",
              children: [
                new BuilderElement({
                  name: "p:cTn",
                  attributes: {
                    id: { key: "id", value: ids.effect },
                    dur: { key: "dur", value: dur },
                    fill: { key: "fill", value: "hold" },
                  },
                  children: [
                    new BuilderElement({
                      name: "p:stCondLst",
                      children: [
                        new BuilderElement({
                          name: "p:cond",
                          attributes: {
                            delay: { key: "delay", value: "0" },
                          },
                        }),
                      ],
                    }),
                  ],
                }),
                tgtEl,
              ],
            }),
          ],
        }),
      );
      break;

    case "colorChange": {
      const color = options.color ?? "FF0000";
      const animClrChildren: XmlComponent[] = [
        new BuilderElement({
          name: "p:cBhvr",
          children: [
            new BuilderElement({
              name: "p:cTn",
              attributes: {
                id: { key: "id", value: ids.effect },
                dur: { key: "dur", value: dur },
                fill: { key: "fill", value: "hold" },
              },
              children: [
                new BuilderElement({
                  name: "p:stCondLst",
                  children: [
                    new BuilderElement({
                      name: "p:cond",
                      attributes: {
                        delay: { key: "delay", value: "0" },
                      },
                    }),
                  ],
                }),
              ],
            }),
            tgtEl,
          ],
        }),
      ];
      // p:by (color animation by value — rgb or hsl)
      if (options.colorByRgb) {
        animClrChildren.push(
          new BuilderElement({
            name: "p:by",
            children: [
              new BuilderElement({
                name: "p:rgb",
                attributes: {
                  r: { key: "r", value: options.colorByRgb.r },
                  g: { key: "g", value: options.colorByRgb.g },
                  b: { key: "b", value: options.colorByRgb.b },
                },
              }),
            ],
          }),
        );
      } else if (options.colorByHsl) {
        animClrChildren.push(
          new BuilderElement({
            name: "p:by",
            children: [
              new BuilderElement({
                name: "p:hsl",
                attributes: {
                  h: { key: "h", value: options.colorByHsl.h },
                  s: { key: "s", value: options.colorByHsl.s },
                  l: { key: "l", value: options.colorByHsl.l },
                },
              }),
            ],
          }),
        );
      }
      // p:from (color animation from value)
      if (options.colorFrom) {
        animClrChildren.push(
          new BuilderElement({
            name: "p:from",
            children: [
              new BuilderElement({
                name: "a:srgbClr",
                attributes: { val: { key: "val", value: options.colorFrom } },
              }),
            ],
          }),
        );
      }
      // p:to (color animation to value)
      animClrChildren.push(
        new BuilderElement({
          name: "p:to",
          children: [
            new BuilderElement({
              name: "a:srgbClr",
              attributes: { val: { key: "val", value: options.colorTo ?? color } },
            }),
          ],
        }),
      );
      children.push(new BuilderElement({ name: "p:animClr", children: animClrChildren }));
      break;
    }

    case "transparency":
      children.push(
        new BuilderElement({
          name: "p:anim",
          attributes: {
            calcmode: { key: "calcmode", value: "lin" },
            valueType: { key: "valueType", value: "num" },
            from: { key: "from", value: "0" },
            to: { key: "to", value: "1" },
          },
          children: [
            new BuilderElement({
              name: "p:cBhvr",
              children: [
                new BuilderElement({
                  name: "p:cTn",
                  attributes: {
                    id: { key: "id", value: ids.effect },
                    dur: { key: "dur", value: dur },
                    fill: { key: "fill", value: "hold" },
                  },
                }),
                tgtEl,
                new BuilderElement({
                  name: "p:attrNameLst",
                  children: [stringContainerObj("p:attrName", "style.opacity")],
                }),
              ],
            }),
          ],
        }),
      );
      break;

    default:
      // boldFlash, wave, pulse, growWithTurn — use p:animEffect with emph filter
      children.push(
        new BuilderElement({
          name: "p:animEffect",
          attributes: {
            transition: { key: "transition", value: "in" },
            filter: { key: "filter", value: emphType },
          },
          children: [
            new BuilderElement({
              name: "p:cBhvr",
              children: [
                new BuilderElement({
                  name: "p:cTn",
                  attributes: {
                    id: { key: "id", value: ids.effect },
                    dur: { key: "dur", value: dur },
                    fill: { key: "fill", value: "hold" },
                  },
                }),
                tgtEl,
              ],
            }),
          ],
        }),
      );
      break;
  }

  return children;
}

function buildPathEffects(
  options: AnimationOptions,
  spid: number,
  ids: { set: number; effect: number },
): XmlComponent[] {
  const pathStr = options.path ?? PATH_STRINGS[options.pathType ?? "customPath"] ?? "";
  const dur = String(options.duration ?? 1000);

  const animMotionAttrs: Record<string, { key: string; value: string }> = {
    origin: { key: "origin", value: "layout" },
    path: { key: "path", value: pathStr },
  };
  if (options.pointsTypes)
    animMotionAttrs.ptsTypes = { key: "ptsTypes", value: options.pointsTypes };

  const animMotionChildren: XmlComponent[] = [
    new BuilderElement({
      name: "p:cBhvr",
      children: [
        new BuilderElement({
          name: "p:cTn",
          attributes: {
            id: { key: "id", value: ids.effect },
            dur: { key: "dur", value: dur },
            fill: { key: "fill", value: "hold" },
          },
          children: [
            new BuilderElement({
              name: "p:stCondLst",
              children: [
                new BuilderElement({
                  name: "p:cond",
                  attributes: { delay: { key: "delay", value: "0" } },
                }),
              ],
            }),
          ],
        }),
        buildTargetElement(spid, options),
      ],
    }),
  ];

  // Motion path from/rCtr (A5)
  if (options.motionFrom) {
    animMotionChildren.push(
      new BuilderElement({
        name: "p:from",
        attributes: {
          x: { key: "x", value: options.motionFrom.x },
          y: { key: "y", value: options.motionFrom.y },
        },
      }),
    );
  }
  if (options.motionRotationCenter) {
    animMotionChildren.push(
      new BuilderElement({
        name: "p:rCtr",
        attributes: {
          x: { key: "x", value: options.motionRotationCenter.x },
          y: { key: "y", value: options.motionRotationCenter.y },
        },
      }),
    );
  }

  return [
    new BuilderElement({
      name: "p:animMotion",
      attributes: animMotionAttrs,
      children: animMotionChildren,
    }),
  ];
}

/**
 * Build p:cmd for media play/pause/stop (goes inside the animation sequence).
 * PowerPoint uses p:cmd type="call" cmd="playFrom(0.0)" for video play.
 */
function buildMediaPlayCommand(
  options: AnimationOptions,
  spid: number,
  ids: { cmd: number },
): XmlComponent {
  const cmdStr = options.fullScreen ? "playFrom(0.0,1.0)" : "playFrom(0.0)";

  return new BuilderElement({
    name: "p:cmd",
    attributes: {
      type: { key: "type", value: "call" },
      cmd: { key: "cmd", value: cmdStr },
    },
    children: [
      new BuilderElement({
        name: "p:cBhvr",
        children: [
          new BuilderElement({
            name: "p:cTn",
            attributes: {
              id: { key: "id", value: ids.cmd },
              dur: { key: "dur", value: String(options.duration ?? 10000) },
              fill: { key: "fill", value: "hold" },
            },
          }),
          buildTargetElement(spid),
        ],
      }),
    ],
  });
}

/**
 * Build generic p:cmd with configurable type (call/evt/verb).
 */
function buildCommand(options: AnimationOptions, spid: number, ids: { cmd: number }): XmlComponent {
  const cmdType = options.commandType ?? "call";
  const cmdStr = options.command ?? "";

  return new BuilderElement({
    name: "p:cmd",
    attributes: {
      type: { key: "type", value: cmdType },
      cmd: { key: "cmd", value: cmdStr },
    },
    children: [
      new BuilderElement({
        name: "p:cBhvr",
        children: [
          new BuilderElement({
            name: "p:cTn",
            attributes: {
              id: { key: "id", value: ids.cmd },
              dur: { key: "dur", value: String(options.duration ?? 1000) },
              fill: { key: "fill", value: "hold" },
            },
          }),
          buildTargetElement(spid),
        ],
      }),
    ],
  });
}

/**
 * Build p:set for an instant property change (setBehavior).
 */
function buildSetBehavior(
  setOptions: NonNullable<AnimationOptions["setBehavior"]>,
  spid: number,
  ids: { set: number },
): XmlComponent {
  const toValName = setOptions.toType === "number" ? "p:numVal" : "p:strVal";

  return new BuilderElement({
    name: "p:set",
    children: [
      new BuilderElement({
        name: "p:cBhvr",
        children: [
          new BuilderElement({
            name: "p:cTn",
            attributes: {
              id: { key: "id", value: ids.set },
              dur: { key: "dur", value: "1" },
              fill: { key: "fill", value: "hold" },
            },
            children: [
              new BuilderElement({
                name: "p:stCondLst",
                children: [
                  new BuilderElement({
                    name: "p:cond",
                    attributes: { delay: { key: "delay", value: "0" } },
                  }),
                ],
              }),
            ],
          }),
          buildTargetElement(spid),
          new BuilderElement({
            name: "p:attrNameLst",
            children: [stringContainerObj("p:attrName", setOptions.attributeName)],
          }),
        ],
      }),
      new BuilderElement({
        name: "p:to",
        children: [
          new BuilderElement({
            name: toValName,
            attributes: { val: { key: "val", value: setOptions.toValue } },
          }),
        ],
      }),
    ],
  });
}

/**
 * Build p:iterate for text-level animation iteration.
 */
function buildIterate(iterate: NonNullable<AnimationOptions["iterate"]>): XmlComponent {
  const attrs: Record<string, { key: string; value: string | number }> = {};
  if (iterate.type) attrs.type = { key: "type", value: iterate.type };
  if (iterate.backwards) attrs.backwards = { key: "backwards", value: 1 };

  const iterChildren: XmlComponent[] = [];
  if (iterate.iteratePercentage !== undefined) {
    iterChildren.push(
      new BuilderElement({
        name: "p:tmPct",
        attributes: { val: { key: "val", value: iterate.iteratePercentage } },
      }),
    );
  } else if (iterate.interval !== undefined) {
    iterChildren.push(
      new BuilderElement({
        name: "p:tmAbs",
        attributes: { val: { key: "val", value: iterate.interval } },
      }),
    );
  }

  return new BuilderElement({
    name: "p:iterate",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
    children: iterChildren.length > 0 ? iterChildren : undefined,
  });
}

/**
 * Build p:animVariant value child (boolVal|intVal|fltVal|strVal|clrVal).
 */
function buildVariantValue(variant: AnimationVariantOptions): XmlComponent {
  if (variant.bool !== undefined) {
    return new BuilderElement({
      name: "p:boolVal",
      attributes: { val: { key: "val", value: variant.bool ? 1 : 0 } },
    });
  }
  if (variant.int !== undefined) {
    return new BuilderElement({
      name: "p:intVal",
      attributes: { val: { key: "val", value: variant.int } },
    });
  }
  if (variant.float !== undefined) {
    return new BuilderElement({
      name: "p:fltVal",
      attributes: { val: { key: "val", value: variant.float } },
    });
  }
  if (variant.color !== undefined) {
    return new BuilderElement({
      name: "p:clrVal",
      children: [
        new BuilderElement({
          name: "a:srgbClr",
          attributes: { val: { key: "val", value: variant.color } },
        }),
      ],
    });
  }
  // string fallback
  return new BuilderElement({
    name: "p:strVal",
    attributes: { val: { key: "val", value: variant.string ?? "" } },
  });
}

/**
 * Build a condition element (p:cond) from EndConditionOptions.
 */
function buildCondition(cond: EndConditionOptions): XmlComponent {
  const attrs: Record<string, { key: string; value: string | number }> = {};
  if (cond.delay !== undefined) attrs.delay = { key: "delay", value: cond.delay };
  if (cond.event !== undefined) attrs.evt = { key: "evt", value: cond.event };

  const children: XmlComponent[] = [];
  if (cond.timeNodeId !== undefined) {
    children.push(
      new BuilderElement({
        name: "p:tn",
        attributes: { val: { key: "val", value: cond.timeNodeId } },
      }),
    );
  } else if (cond.runtimeNode !== undefined) {
    children.push(
      new BuilderElement({
        name: "p:rtn",
        attributes: { val: { key: "val", value: cond.runtimeNode } },
      }),
    );
  }

  return new BuilderElement({
    name: "p:cond",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
    children: children.length > 0 ? children : undefined,
  });
}

/**
 * Build p:video/p:audio with p:cMediaNode (goes as sibling of p:seq).
 * This is the media state controller, separate from the play command.
 */
function buildMediaStateNode(
  options: AnimationOptions,
  spid: number,
  ids: { mediaCtn: number },
): XmlComponent {
  const isVideo = options.mediaType === "playVideo";
  const elementName = isVideo ? "p:video" : "p:audio";

  const mediaAttrs: Record<string, { key: string; value: string | number | boolean }> = {};
  if (!isVideo && options.isNarration) {
    mediaAttrs.isNarration = { key: "isNarration", value: true };
  }
  if (isVideo && options.fullScreen) {
    mediaAttrs.fullScrn = { key: "fullScrn", value: true };
  }

  const cMediaNodeAttrs: Record<string, { key: string; value: string | number }> = {};
  if (options.volume !== undefined) {
    cMediaNodeAttrs.vol = { key: "vol", value: options.volume * 1000 };
  } else {
    cMediaNodeAttrs.vol = { key: "vol", value: 80000 };
  }
  if (options.mute) {
    cMediaNodeAttrs.mute = { key: "mute", value: 1 };
  }
  if (options.numberOfSlides !== undefined) {
    cMediaNodeAttrs.numSld = { key: "numSld", value: options.numberOfSlides };
  }

  return new BuilderElement({
    name: elementName,
    attributes: Object.keys(mediaAttrs).length > 0 ? mediaAttrs : undefined,
    children: [
      new BuilderElement({
        name: "p:cMediaNode",
        attributes: cMediaNodeAttrs,
        children: [
          new BuilderElement({
            name: "p:cTn",
            attributes: {
              id: { key: "id", value: ids.mediaCtn },
              fill: { key: "fill", value: "hold" },
              display: { key: "display", value: "0" },
            },
            children: [
              new BuilderElement({
                name: "p:stCondLst",
                children: [
                  new BuilderElement({
                    name: "p:cond",
                    attributes: {
                      delay: { key: "delay", value: "indefinite" },
                    },
                  }),
                ],
              }),
            ],
          }),
          buildTargetElement(spid),
        ],
      }),
    ],
  });
}

function buildPropertyAnimation(
  options: AnimationOptions,
  spid: number,
  ids: { cBhvr: number },
): XmlComponent {
  const attrs: Record<string, { key: string; value: string }> = {};
  if (options.calcMode) attrs.calcmode = { key: "calcmode", value: options.calcMode };
  if (options.valueType) attrs.valueType = { key: "valueType", value: options.valueType };
  if (options.from !== undefined) attrs.from = { key: "from", value: options.from };
  if (options.to !== undefined) attrs.to = { key: "to", value: options.to };
  if (options.animBy !== undefined) attrs.by = { key: "by", value: options.animBy };
  if (options.formula !== undefined) attrs.fmla = { key: "fmla", value: options.formula };
  if (options.colorSpace !== undefined) attrs.clrSpc = { key: "clrSpc", value: options.colorSpace };
  if (options.rotationAngle !== undefined)
    attrs.rAng = { key: "rAng", value: String(options.rotationAngle) };

  const cBhvrChildren: XmlComponent[] = [
    new BuilderElement({
      name: "p:cTn",
      attributes: {
        id: { key: "id", value: ids.cBhvr },
        dur: { key: "dur", value: String(options.duration ?? 500) },
        fill: { key: "fill", value: "hold" },
      },
    }),
    buildTargetElement(spid, options),
  ];

  if (options.attributeName) {
    cBhvrChildren.push(
      new BuilderElement({
        name: "p:attrNameLst",
        children: [stringContainerObj("p:attrName", options.attributeName)],
      }),
    );
  }

  const cBhvrAttrs: Record<string, { key: string; value: string | number }> = {};
  if (options.additive !== undefined)
    cBhvrAttrs.additive = { key: "additive", value: options.additive };
  if (options.accumulate !== undefined)
    cBhvrAttrs.accumulate = { key: "accumulate", value: options.accumulate };
  if (options.transformType !== undefined)
    cBhvrAttrs.xfrmType = { key: "xfrmType", value: options.transformType };
  if (options.runtimeContext !== undefined)
    cBhvrAttrs.rctx = { key: "rctx", value: options.runtimeContext };
  if (options.override !== undefined)
    cBhvrAttrs.override = { key: "override", value: options.override };

  const animChildren: XmlComponent[] = [
    new BuilderElement({
      name: "p:cBhvr",
      attributes: Object.keys(cBhvrAttrs).length > 0 ? cBhvrAttrs : undefined,
      children: cBhvrChildren,
    }),
  ];

  // Build tavLst if from/to are specified (support variant value types)
  if (
    options.from !== undefined ||
    options.to !== undefined ||
    options.colorFrom ||
    options.colorTo
  ) {
    const tavList: XmlComponent[] = [];

    // Build variant child for p:val
    const buildVariantChild = (val: string, isColor?: boolean): XmlComponent => {
      if (isColor) {
        return new BuilderElement({
          name: "p:clrVal",
          children: [
            new BuilderElement({
              name: "a:srgbClr",
              attributes: { val: { key: "val", value: val } },
            }),
          ],
        });
      }
      if (options.variantInt !== undefined) {
        return new BuilderElement({
          name: "p:intVal",
          attributes: { val: { key: "val", value: Number(val) } },
        });
      }
      if (options.variantFloat !== undefined) {
        return new BuilderElement({
          name: "p:fltVal",
          attributes: { val: { key: "val", value: Number(val) } },
        });
      }
      if (options.variantBool !== undefined) {
        return new BuilderElement({
          name: "p:boolVal",
          attributes: { val: { key: "val", value: val === "true" ? 1 : 0 } },
        });
      }
      return new BuilderElement({
        name: "p:strVal",
        attributes: { val: { key: "val", value: val } },
      });
    };

    if (options.colorFrom) {
      tavList.push(
        new BuilderElement({
          name: "p:tav",
          attributes: { tm: { key: "tm", value: "0" } },
          children: [
            new BuilderElement({
              name: "p:val",
              children: [buildVariantChild(options.colorFrom, true)],
            }),
          ],
        }),
      );
    } else if (options.from !== undefined) {
      tavList.push(
        new BuilderElement({
          name: "p:tav",
          attributes: { tm: { key: "tm", value: "0" } },
          children: [
            new BuilderElement({ name: "p:val", children: [buildVariantChild(options.from)] }),
          ],
        }),
      );
    }
    if (options.colorTo) {
      tavList.push(
        new BuilderElement({
          name: "p:tav",
          attributes: { tm: { key: "tm", value: "100000" } },
          children: [
            new BuilderElement({
              name: "p:val",
              children: [buildVariantChild(options.colorTo, true)],
            }),
          ],
        }),
      );
    } else if (options.to !== undefined) {
      tavList.push(
        new BuilderElement({
          name: "p:tav",
          attributes: { tm: { key: "tm", value: "100000" } },
          children: [
            new BuilderElement({ name: "p:val", children: [buildVariantChild(options.to)] }),
          ],
        }),
      );
    }
    if (tavList.length > 0) {
      animChildren.push(new BuilderElement({ name: "p:tavLst", children: tavList }));
    }
  }

  return new BuilderElement({
    name: "p:anim",
    attributes: Object.keys(attrs).length > 0 ? attrs : undefined,
    children: animChildren,
  });
}

// --- Build list (bldLst) ---

function buildBuildList(
  builds: readonly AnimationBuildOptions[],
  nextId: () => number,
): XmlComponent {
  const bldChildren: XmlComponent[] = [];

  for (const bld of builds) {
    const bldAttrs: Record<string, { key: string; value: string | number }> = {
      spid: { key: "spid", value: bld.spid },
      grpId: { key: "grpId", value: bld.grpId },
    };
    if (bld.uiExpand) bldAttrs.uiExpand = { key: "uiExpand", value: 1 };

    let elementName: string;
    const bldChildrenInner: XmlComponent[] = [];

    switch (bld.type) {
      case "paragraph": {
        elementName = "p:bldP";
        if (bld.build) bldAttrs.build = { key: "build", value: bld.build };
        if (bld.bldLvl !== undefined) bldAttrs.bldLvl = { key: "bldLvl", value: bld.bldLvl };
        if (bld.animBg !== undefined)
          bldAttrs.animBg = { key: "animBg", value: bld.animBg ? 1 : 0 };
        if (bld.autoUpdateAnimBg !== undefined)
          bldAttrs.autoUpdateAnimBg = {
            key: "autoUpdateAnimBg",
            value: bld.autoUpdateAnimBg ? 1 : 0,
          };
        if (bld.rev) bldAttrs.rev = { key: "rev", value: 1 };
        if (bld.advAuto !== undefined) bldAttrs.advAuto = { key: "advAuto", value: bld.advAuto };
        // Templates (tmplLst)
        if (bld.templates && bld.templates.length > 0) {
          const tmplChildren = bld.templates.map((tmpl) => {
            const tmplAttrs: Record<string, { key: string; value: string | number }> = {};
            if (tmpl.lvl !== undefined) tmplAttrs.lvl = { key: "lvl", value: tmpl.lvl };
            const tnLstChildren = tmpl.children.map(() => {
              const tid = nextId();
              return new BuilderElement({
                name: "p:par",
                children: [
                  new BuilderElement({
                    name: "p:cTn",
                    attributes: { id: { key: "id", value: tid } },
                  }),
                ],
              });
            });
            return new BuilderElement({
              name: "p:tmpl",
              attributes: Object.keys(tmplAttrs).length > 0 ? tmplAttrs : undefined,
              children: [new BuilderElement({ name: "p:tnLst", children: tnLstChildren })],
            });
          });
          bldChildrenInner.push(new BuilderElement({ name: "p:tmplLst", children: tmplChildren }));
        }
        break;
      }
      case "diagram": {
        elementName = "p:bldDgm";
        if (bld.diagramBuild) bldAttrs.bld = { key: "bld", value: bld.diagramBuild };
        break;
      }
      case "oleChart": {
        elementName = "p:bldOleChart";
        if (bld.oleChartBuild) bldAttrs.bld = { key: "bld", value: bld.oleChartBuild };
        if (bld.oleChartAnimBg !== undefined)
          bldAttrs.animBg = { key: "animBg", value: bld.oleChartAnimBg ? 1 : 0 };
        break;
      }
      case "graphic": {
        elementName = "p:bldGraphic";
        if (bld.graphicBuildAsOne) {
          bldChildrenInner.push(new BuilderElement({ name: "p:bldAsOne" }));
        } else {
          bldChildrenInner.push(new BuilderElement({ name: "p:bldSub" }));
        }
        break;
      }
    }

    bldChildren.push(
      new BuilderElement({
        name: elementName,
        attributes: bldAttrs,
        children: bldChildrenInner.length > 0 ? bldChildrenInner : undefined,
      }),
    );
  }

  return new BuilderElement({ name: "p:bldLst", children: bldChildren });
}

// --- Main class ---

export interface AnimationEntry {
  readonly spid: number;
  readonly options: AnimationOptions;
}

/**
 * p:timing — Slide timing for shape animations.
 */
export class SlideTiming extends XmlComponent {
  public constructor(entries: readonly AnimationEntry[]) {
    super("p:timing");

    if (entries.length === 0) return;

    let id = 1;
    const rootCtnId = id++;
    const seqCtnId = id++;

    // Check for builds from first entry
    const builds = entries[0]?.options.builds;
    const previousAction = entries[0]?.options.previousAction;

    const animationNodes: XmlComponent[] = [];
    const mediaStateNodes: XmlComponent[] = [];
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
      let effectChildren: XmlComponent[];

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
      const cTnAttrs: Record<string, { key: string; value: string | number }> = {
        id: { key: "id", value: effectCtnId },
        fill: { key: "fill", value: "hold" },
        nodeType: { key: "nodeType", value: nodeType },
      };
      if (!options.mediaType && !options.attributeName) {
        cTnAttrs.presetID = { key: "presetID", value: presetId };
        cTnAttrs.presetClass = { key: "presetClass", value: presetClass };
        cTnAttrs.presetSubtype = { key: "presetSubtype", value: presetSubtype };
      }
      if (options.mediaType) {
        cTnAttrs.presetID = { key: "presetID", value: presetId };
        cTnAttrs.presetClass = { key: "presetClass", value: presetClass };
        cTnAttrs.presetSubtype = { key: "presetSubtype", value: presetSubtype };
      }
      if (options.speed !== undefined) cTnAttrs.spd = { key: "spd", value: String(options.speed) };
      if (options.repeatCount !== undefined)
        cTnAttrs.repeatCount = { key: "repeatCount", value: String(options.repeatCount) };
      if (options.autoReverse) cTnAttrs.autoRev = { key: "autoRev", value: 1 };
      if (options.repeatDuration !== undefined)
        cTnAttrs.repeatDur = { key: "repeatDur", value: options.repeatDuration };
      if (options.acceleration !== undefined)
        cTnAttrs.accel = { key: "accel", value: options.acceleration };
      if (options.deceleration !== undefined)
        cTnAttrs.decel = { key: "decel", value: options.deceleration };
      if (options.restart !== undefined)
        cTnAttrs.restart = { key: "restart", value: options.restart };
      if (options.syncBehavior !== undefined)
        cTnAttrs.syncBehavior = { key: "syncBehavior", value: options.syncBehavior };
      if (options.timeFilter !== undefined)
        cTnAttrs.tmFilter = { key: "tmFilter", value: options.timeFilter };
      if (options.eventFilter !== undefined)
        cTnAttrs.evtFilter = { key: "evtFilter", value: options.eventFilter };
      if (options.display !== undefined)
        cTnAttrs.display = { key: "display", value: options.display ? 1 : 0 };
      if (options.masterRelation !== undefined)
        cTnAttrs.masterRel = { key: "masterRel", value: options.masterRelation };
      if (options.buildLevel !== undefined)
        cTnAttrs.bldLvl = { key: "bldLvl", value: options.buildLevel };
      if (options.groupId !== undefined) cTnAttrs.grpId = { key: "grpId", value: options.groupId };
      if (options.afterEffect !== undefined)
        cTnAttrs.afterEffect = { key: "afterEffect", value: options.afterEffect ? 1 : 0 };
      if (options.nodePlaceholder !== undefined)
        cTnAttrs.nodePh = { key: "nodePh", value: options.nodePlaceholder ? 1 : 0 };
      if (options.advanceAfterTime !== undefined)
        cTnAttrs.advAuto = { key: "advAuto", value: options.advanceAfterTime };
      if (options.animateBackground !== undefined)
        cTnAttrs.animBg = { key: "animBg", value: options.animateBackground ? 1 : 0 };
      if (options.autoUpdateAnimationBackground !== undefined)
        cTnAttrs.autoUpdateAnimBg = {
          key: "autoUpdateAnimBg",
          value: options.autoUpdateAnimationBackground ? 1 : 0,
        };

      // Build effect cTn
      const effectCtnChildren: XmlComponent[] = [
        new BuilderElement({
          name: "p:stCondLst",
          children: [
            new BuilderElement({
              name: "p:cond",
              attributes: {
                delay: { key: "delay", value: String(options.delay ?? 0) },
              },
            }),
          ],
        }),
      ];

      // Add iterate container if specified
      if (options.iterate) {
        effectCtnChildren.push(buildIterate(options.iterate));
      }

      // Add endCondLst (A2)
      if (options.endConditions && options.endConditions.length > 0) {
        effectCtnChildren.push(
          new BuilderElement({
            name: "p:endCondLst",
            children: options.endConditions.map(buildCondition),
          }),
        );
      }

      // Add endSync (A2) — directly CT_TLTimeCondition, no cond wrapper
      if (options.endSyncCondition) {
        const syncAttrs: Record<string, { key: string; value: string }> = {};
        if (options.endSyncCondition.delay !== undefined)
          syncAttrs.delay = { key: "delay", value: options.endSyncCondition.delay };
        if (options.endSyncCondition.event !== undefined)
          syncAttrs.evt = { key: "evt", value: options.endSyncCondition.event };

        const syncChildren: XmlComponent[] = [];
        if (options.endSyncCondition.timeNodeId !== undefined) {
          syncChildren.push(
            new BuilderElement({
              name: "p:tn",
              attributes: { val: { key: "val", value: options.endSyncCondition.timeNodeId } },
            }),
          );
        } else if (options.endSyncCondition.runtimeNode !== undefined) {
          syncChildren.push(
            new BuilderElement({
              name: "p:rtn",
              attributes: { val: { key: "val", value: options.endSyncCondition.runtimeNode } },
            }),
          );
        }

        effectCtnChildren.push(
          new BuilderElement({
            name: "p:endSync",
            attributes: Object.keys(syncAttrs).length > 0 ? syncAttrs : undefined,
            children: syncChildren.length > 0 ? syncChildren : undefined,
          }),
        );
      }

      // Add childTnLst with possible excl wrapper
      const childTnListChildren = options.exclusiveMode
        ? [
            new BuilderElement({
              name: "p:excl",
              children: [
                new BuilderElement({
                  name: "p:cTn",
                  attributes: { id: { key: "id", value: id++ } },
                  children: [
                    new BuilderElement({
                      name: "p:childTnLst",
                      children: effectChildren,
                    }),
                  ],
                }),
              ],
            }),
          ]
        : effectChildren;

      effectCtnChildren.push(
        new BuilderElement({
          name: "p:childTnLst",
          children: childTnListChildren,
        }),
      );

      // Add subTnLst (A2)
      if (options.subTimeNodes && options.subTimeNodes.length > 0) {
        effectCtnChildren.push(
          new BuilderElement({
            name: "p:subTnLst",
            children: options.subTimeNodes.map((subOpts) => {
              const subId = id++;
              return new BuilderElement({
                name: "p:par",
                children: [
                  new BuilderElement({
                    name: "p:cTn",
                    attributes: {
                      id: { key: "id", value: subId },
                      dur: { key: "dur", value: String(subOpts.duration ?? 0) },
                    },
                    children: [
                      new BuilderElement({
                        name: "p:stCondLst",
                        children: [
                          new BuilderElement({
                            name: "p:cond",
                            attributes: {
                              delay: { key: "delay", value: String(subOpts.delay ?? 0) },
                            },
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              });
            }),
          }),
        );
      }

      const effectCtn = new BuilderElement({
        name: "p:cTn",
        attributes: cTnAttrs,
        children: effectCtnChildren,
      });

      // Wrap in inner par
      const innerPar = new BuilderElement({
        name: "p:par",
        children: [effectCtn],
      });

      // Wrap in outer par (click group)
      const outerPar = new BuilderElement({
        name: "p:par",
        children: [
          new BuilderElement({
            name: "p:cTn",
            attributes: {
              id: { key: "id", value: groupCtnId },
              fill: { key: "fill", value: "hold" },
            },
            children: [
              new BuilderElement({
                name: "p:stCondLst",
                children: [
                  new BuilderElement({
                    name: "p:cond",
                    attributes: { delay: { key: "delay", value: "0" } },
                  }),
                ],
              }),
              new BuilderElement({
                name: "p:childTnLst",
                children: [innerPar],
              }),
            ],
          }),
        ],
      });

      animationNodes.push(outerPar);
    }

    // Assemble the full timing tree
    this.root.push(
      new BuilderElement({
        name: "p:tnLst",
        children: [
          new BuilderElement({
            name: "p:par",
            children: [
              new BuilderElement({
                name: "p:cTn",
                attributes: {
                  id: { key: "id", value: rootCtnId },
                  dur: { key: "dur", value: "indefinite" },
                  restart: { key: "restart", value: "never" },
                  nodeType: { key: "nodeType", value: "tmRoot" },
                },
                children: [
                  new BuilderElement({
                    name: "p:childTnLst",
                    children: [
                      new BuilderElement({
                        name: "p:seq",
                        attributes: {
                          concurrent: { key: "concurrent", value: 1 },
                          nextAc: { key: "nextAc", value: "seek" },
                          ...(previousAction !== undefined
                            ? { prevAc: { key: "prevAc", value: previousAction } }
                            : {}),
                        },
                        children: [
                          new BuilderElement({
                            name: "p:cTn",
                            attributes: {
                              id: { key: "id", value: seqCtnId },
                              dur: {
                                key: "dur",
                                value: "indefinite",
                              },
                              nodeType: {
                                key: "nodeType",
                                value: "mainSeq",
                              },
                            },
                            children: [
                              new BuilderElement({
                                name: "p:childTnLst",
                                children: animationNodes,
                              }),
                            ],
                          }),
                          new BuilderElement({
                            name: "p:prevCondLst",
                            children: [
                              new BuilderElement({
                                name: "p:cond",
                                attributes: {
                                  evt: {
                                    key: "evt",
                                    value: "onPrev",
                                  },
                                  delay: {
                                    key: "delay",
                                    value: "0",
                                  },
                                },
                                children: [
                                  new BuilderElement({
                                    name: "p:tgtEl",
                                    children: [
                                      new BuilderElement({
                                        name: "p:sldTgt",
                                      }),
                                    ],
                                  }),
                                ],
                              }),
                            ],
                          }),
                          new BuilderElement({
                            name: "p:nextCondLst",
                            children: [
                              new BuilderElement({
                                name: "p:cond",
                                attributes: {
                                  evt: {
                                    key: "evt",
                                    value: "onNext",
                                  },
                                  delay: {
                                    key: "delay",
                                    value: "0",
                                  },
                                },
                                children: [
                                  new BuilderElement({
                                    name: "p:tgtEl",
                                    children: [
                                      new BuilderElement({
                                        name: "p:sldTgt",
                                      }),
                                    ],
                                  }),
                                ],
                              }),
                            ],
                          }),
                        ],
                      }),
                      ...mediaStateNodes,
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    );

    // bldLst — sibling of tnLst in p:timing (CT_SlideTiming: tnLst → bldLst → extLst)
    if (builds) {
      this.root.push(buildBuildList(builds, () => id++));
    }
  }
}
