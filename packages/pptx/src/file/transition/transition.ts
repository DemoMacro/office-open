import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";

const EMPTY_TYPES = new Set([
    "circle",
    "dissolve",
    "diamond",
    "newsflash",
    "plus",
    "wedge",
    "random",
]);

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

/**
 * p:transition — Slide transition effect.
 */
export class Transition extends XmlComponent {
    public constructor(options: ITransitionOptions = {}) {
        super("p:transition");

        const attrs: Record<string, { readonly key: string; readonly value: string | number | boolean }> = {};
        if (options.speed) attrs.spd = { key: "spd", value: options.speed };
        if (options.advanceOnClick !== undefined)
            attrs.advClick = { key: "advClick", value: options.advanceOnClick ? 1 : 0 };
        if (options.advanceAfterTime !== undefined)
            attrs.advTm = { key: "advTm", value: options.advanceAfterTime };

        this.root.push(new NextAttributeComponent(attrs));

        if (options.type) {
            const child = this.buildTransitionElement(options);
            if (child) {
                this.root.push(child);
            }
        }
    }

    private buildTransitionElement(options: ITransitionOptions): XmlComponent | undefined {
        const { type, dir, orient, thruBlk, spokes } = options;

        if (!type) return undefined;

        // Empty transitions (no attributes)
        if (EMPTY_TYPES.has(type)) {
            return new BuilderElement({ name: `p:${type}` });
        }

        // Orientation transitions: blinds, checker, comb, randomBar
        if (ORIENT_TYPES.has(type) && dir) {
            return new BuilderElement({
                name: `p:${type}`,
                attributes: { dir: { key: "dir", value: dir } },
            });
        }
        if (ORIENT_TYPES.has(type)) {
            return new BuilderElement({ name: `p:${type}` });
        }

        // Side direction transitions: push, wipe
        if (SIDE_DIR_TYPES.has(type)) {
            const elementAttrs: Record<string, { readonly key: string; readonly value: string }> = {};
            if (dir) elementAttrs.dir = { key: "dir", value: dir };
            return new BuilderElement({ name: `p:${type}`, attributes: elementAttrs });
        }

        // Eight direction transitions: cover, pull
        if (EIGHT_DIR_TYPES.has(type)) {
            const elementAttrs: Record<string, { readonly key: string; readonly value: string }> = {};
            if (dir) elementAttrs.dir = { key: "dir", value: dir };
            return new BuilderElement({ name: `p:${type}`, attributes: elementAttrs });
        }

        // Corner direction: strips
        if (type === "strips") {
            const elementAttrs: Record<string, { readonly key: string; readonly value: string }> = {};
            if (dir) elementAttrs.dir = { key: "dir", value: dir };
            return new BuilderElement({ name: "p:strips", attributes: elementAttrs });
        }

        // Fade / Cut: thruBlk
        if (type === "fade" || type === "cut") {
            const elementAttrs: Record<string, { readonly key: string; readonly value: string | number | boolean }> = {};
            if (thruBlk !== undefined)
                elementAttrs.thruBlk = { key: "thruBlk", value: thruBlk ? 1 : 0 };
            return new BuilderElement({ name: `p:${type}`, attributes: elementAttrs });
        }

        // Split: orient + dir
        if (type === "split") {
            const elementAttrs: Record<string, { readonly key: string; readonly value: string | number | boolean }> = {
                orient: { key: "orient", value: orient ?? "horz" },
                dir: { key: "dir", value: dir ?? "out" },
            };
            return new BuilderElement({ name: "p:split", attributes: elementAttrs });
        }

        // Wheel: spokes
        if (type === "wheel") {
            const elementAttrs: Record<string, { readonly key: string; readonly value: string | number }> = {
                spokes: { key: "spokes", value: spokes ?? 4 },
            };
            return new BuilderElement({ name: "p:wheel", attributes: elementAttrs });
        }

        // Zoom: dir (in/out)
        if (type === "zoom") {
            const elementAttrs: Record<string, { readonly key: string; readonly value: string }> = {};
            if (dir) elementAttrs.dir = { key: "dir", value: dir };
            return new BuilderElement({ name: "p:zoom", attributes: elementAttrs });
        }

        return new BuilderElement({ name: `p:${type}` });
    }
}
