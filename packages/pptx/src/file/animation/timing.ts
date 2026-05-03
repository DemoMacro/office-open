import { BuilderElement, StringContainer, XmlComponent } from "@file/xml-components";

import type { IAnimationOptions, AnimationType } from "./types";

const PRESET_IDS: Record<AnimationType, number> = {
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

        // Build animation nodes grouped by click (onClick triggers create new click groups)
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

            // For afterEffect, add delay from previous click group
            if (options.trigger === "afterPrevious" && i > 0) {
                clickGroupDelay += (options.duration ?? 500) + (options.delay ?? 0);
            }
            // For new clickEffect, reset delay
            if (options.trigger === "onClick" || options.trigger === undefined) {
                clickGroupDelay = 0;
            }

            const groupCtnId = id++;
            const effectCtnId = id++;
            const setCtnId = id++;
            const effectCtnId2 = id++;

            const presetId = PRESET_IDS[options.type];
            const presetSubtype = options.direction
                ? (DIRECTION_SUBTYPES[options.direction] ?? 0)
                : 0;

            // Build the animation effect children
            const effectChildren: XmlComponent[] = [];

            // <p:set> to make shape visible
            effectChildren.push(
                new BuilderElement({
                    name: "p:set",
                    children: [
                        new BuilderElement({
                            name: "p:cBhvr",
                            children: [
                                new BuilderElement({
                                    name: "p:cTn",
                                    attributes: {
                                        id: { key: "id", value: setCtnId },
                                        dur: { key: "dur", value: "1" },
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
                                    children: [
                                        new StringContainer("p:attrName", "style.visibility"),
                                    ],
                                }),
                            ],
                        }),
                        new BuilderElement({
                            name: "p:to",
                            children: [
                                new BuilderElement({
                                    name: "p:strVal",
                                    attributes: { val: { key: "val", value: "visible" } },
                                }),
                            ],
                        }),
                    ],
                }),
            );

            // Add animEffect for non-appear types
            if (options.type !== "appear") {
                const filterBase = FILTER_MAP[options.type];
                const dirMap = DIRECTION_FILTER[options.type];
                const dirFilter =
                    options.direction && dirMap ? dirMap[options.direction] : undefined;

                let filter = filterBase;
                if (options.type === "wheel") {
                    filter = "wheel(4)";
                } else if (options.type === "split") {
                    const dir = options.direction ?? "horizontal";
                    filter = `split(${dir})`;
                } else if (dirFilter) {
                    filter = `${filterBase}(${dirFilter})`;
                }

                effectChildren.push(
                    new BuilderElement({
                        name: "p:animEffect",
                        attributes: {
                            transition: { key: "transition", value: "in" },
                            filter: { key: "filter", value: filter },
                        },
                        children: [
                            new BuilderElement({
                                name: "p:cBhvr",
                                children: [
                                    new BuilderElement({
                                        name: "p:cTn",
                                        attributes: {
                                            id: { key: "id", value: effectCtnId2 },
                                            dur: {
                                                key: "dur",
                                                value: String(options.duration ?? 500),
                                            },
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

            // Build effect cTn
            const effectCtn = new BuilderElement({
                name: "p:cTn",
                attributes: {
                    id: { key: "id", value: effectCtnId },
                    presetID: { key: "presetID", value: presetId },
                    presetClass: { key: "presetClass", value: "entr" },
                    presetSubtype: { key: "presetSubtype", value: presetSubtype },
                    fill: { key: "fill", value: "hold" },
                    nodeType: { key: "nodeType", value: nodeType },
                },
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
