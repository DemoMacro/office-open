import type { IContext, IXmlableObject } from "@file/xml-components";
import { BaseXmlComponent } from "@file/xml-components";

const ORIENT_TYPES = new Set(["blinds", "checker", "comb", "randomBar"]);

const SIDE_DIR_TYPES = new Set(["push", "wipe"]);

const EIGHT_DIR_TYPES = new Set(["cover", "pull"]);

export type TransitionType =
    | "fade"
    | "push"
    | "wipe"
    | "split"
    | "blinds"
    | "checker"
    | "comb"
    | "randomBar"
    | "cover"
    | "pull"
    | "strips"
    | "wheel"
    | "zoom"
    | "circle"
    | "dissolve"
    | "diamond"
    | "newsflash"
    | "plus"
    | "wedge"
    | "random"
    | "cut";

export interface ITransitionOptions {
    readonly type?: TransitionType;
    readonly speed?: "slow" | "med" | "fast";
    readonly advanceOnClick?: boolean;
    readonly advanceAfterTime?: number;
    readonly dir?: string;
    readonly orient?: "horz" | "vert";
    readonly thruBlk?: boolean;
    readonly spokes?: number;
}

function buildTransitionElement(
    type: TransitionType,
    dir?: string,
    orient?: string,
    thruBlk?: boolean,
    spokes?: number,
): IXmlableObject {
    const attrs: Record<string, string | number> = {};

    if (ORIENT_TYPES.has(type) && dir) {
        attrs.dir = dir;
    } else if (SIDE_DIR_TYPES.has(type) && dir) {
        attrs.dir = dir;
    } else if (EIGHT_DIR_TYPES.has(type) && dir) {
        attrs.dir = dir;
    } else if (type === "strips" && dir) {
        attrs.dir = dir;
    } else if ((type === "fade" || type === "cut") && thruBlk !== undefined) {
        attrs.thruBlk = thruBlk ? 1 : 0;
    } else if (type === "split") {
        attrs.orient = orient ?? "horz";
        attrs.dir = dir ?? "out";
    } else if (type === "wheel") {
        attrs.spokes = spokes ?? 4;
    } else if (type === "zoom" && dir) {
        attrs.dir = dir;
    }

    return { [`p:${type}`]: Object.keys(attrs).length > 0 ? { _attr: attrs } : {} };
}

export function buildTransition(options: ITransitionOptions): IXmlableObject {
    const children: IXmlableObject[] = [];
    const attrs: Record<string, string | number> = {};
    if (options.speed) attrs.spd = options.speed;
    if (options.advanceOnClick !== undefined) attrs.advClick = options.advanceOnClick ? 1 : 0;
    if (options.advanceAfterTime !== undefined) attrs.advTm = options.advanceAfterTime;
    if (Object.keys(attrs).length > 0) children.push({ _attr: attrs });

    if (options.type) {
        children.push(
            buildTransitionElement(
                options.type,
                options.dir,
                options.orient,
                options.thruBlk,
                options.spokes,
            ),
        );
    }

    return { "p:transition": children.length === 0 ? {} : children };
}

/**
 * p:transition — Slide transition effect.
 * Lazy: stores options, builds XML object in prepForXml.
 */
export class Transition extends BaseXmlComponent {
    private readonly options: ITransitionOptions;

    public constructor(options: ITransitionOptions = {}) {
        super("p:transition");
        this.options = options;
    }

    public override prepForXml(_context: IContext): IXmlableObject {
        return buildTransition(this.options);
    }
}
