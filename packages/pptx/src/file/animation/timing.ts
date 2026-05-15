import { BuilderElement, StringContainer, XmlComponent } from "@file/xml-components";

import type {
    AnimationClass,
    AnimationType,
    EmphasisType,
    IAnimationOptions,
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

function resolvePresetId(options: IAnimationOptions): number {
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

function resolvePresetClass(options: IAnimationOptions): AnimationClass {
    if (options.pathType) return "emph";
    return options.class ?? "entr";
}

function resolvePresetSubtype(options: IAnimationOptions): number {
    if (options.direction) {
        return DIRECTION_SUBTYPES[options.direction] ?? 0;
    }
    return 0;
}

function buildEntrOrExitEffects(
    options: IAnimationOptions,
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
                            children: [new StringContainer("p:attrName", "style.visibility")],
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

        children.push(
            new BuilderElement({
                name: "p:animEffect",
                attributes: {
                    transition: { key: "transition", value: transition },
                    filter: { key: "filter", value: filter },
                },
                children: [
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
                ],
            }),
        );
    }

    return children;
}

function buildEmphasisEffects(
    options: IAnimationOptions,
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
                                        ...(options.autoReverse
                                            ? { autoRev: { key: "autoRev", value: 1 } }
                                            : {}),
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
            children.push(
                new BuilderElement({
                    name: "p:animClr",
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
                        new BuilderElement({
                            name: "p:to",
                            children: [
                                new BuilderElement({
                                    name: "a:srgbClr",
                                    attributes: { val: { key: "val", value: color } },
                                }),
                            ],
                        }),
                    ],
                }),
            );
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
                                    children: [new StringContainer("p:attrName", "style.opacity")],
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
    options: IAnimationOptions,
    spid: number,
    ids: { set: number; effect: number },
): XmlComponent[] {
    const pathStr = options.path ?? PATH_STRINGS[options.pathType ?? "customPath"] ?? "";
    const dur = String(options.duration ?? 1000);

    return [
        new BuilderElement({
            name: "p:animMotion",
            attributes: {
                origin: { key: "origin", value: "layout" },
                path: { key: "path", value: pathStr },
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
                    ],
                }),
            ],
        }),
    ];
}

// --- Main class ---

interface IAnimationEntry {
    readonly spid: number;
    readonly options: IAnimationOptions;
}

/**
 * p:timing — Slide timing for shape animations.
 */
export class SlideTiming extends XmlComponent {
    public constructor(entries: readonly IAnimationEntry[]) {
        super("p:timing");

        if (entries.length === 0) return;

        let id = 1;
        const rootCtnId = id++;
        const seqCtnId = id++;

        const animationNodes: XmlComponent[] = [];
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

            if (options.pathType) {
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

            // Build cTn attributes with optional speed/repeatCount/autoReverse
            const cTnAttrs: Record<string, { key: string; value: string | number }> = {
                id: { key: "id", value: effectCtnId },
                presetID: { key: "presetID", value: presetId },
                presetClass: { key: "presetClass", value: presetClass },
                presetSubtype: { key: "presetSubtype", value: presetSubtype },
                fill: { key: "fill", value: "hold" },
                nodeType: { key: "nodeType", value: nodeType },
            };
            if (options.speed !== undefined)
                cTnAttrs.spd = { key: "spd", value: String(options.speed) };
            if (options.repeatCount !== undefined)
                cTnAttrs.repeatCount = { key: "repeatCount", value: String(options.repeatCount) };
            if (options.autoReverse) cTnAttrs.autoRev = { key: "autoRev", value: 1 };

            // Build effect cTn
            const effectCtn = new BuilderElement({
                name: "p:cTn",
                attributes: cTnAttrs,
                children: [
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
                    new BuilderElement({
                        name: "p:childTnLst",
                        children: effectChildren,
                    }),
                ],
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
                                        ],
                                    }),
                                ],
                            }),
                        ],
                    }),
                ],
            }),
        );
    }
}
