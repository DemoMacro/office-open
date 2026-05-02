import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";

import type { ITransform2DOptions } from "./transform-2d";

export interface IGroupTransform2DOptions extends ITransform2DOptions {
    readonly childOffsetX?: number;
    readonly childOffsetY?: number;
    readonly childExtentWidth?: number;
    readonly childExtentHeight?: number;
}

/**
 * a:xfrm — Group transform (CT_GroupTransform2D).
 * Extends regular Transform2D with child offset (chOff) and child extent (chExt).
 */
export class GroupTransform2D extends XmlComponent {
    public constructor(options: IGroupTransform2DOptions, prefix: "a" | "p" = "a") {
        super(`${prefix}:xfrm`);

        const attrs: Record<
            string,
            { readonly key: string; readonly value: string | boolean | number | undefined }
        > = {};
        if (options.flipH !== undefined) attrs.flipH = { key: "flipH", value: options.flipH };
        if (options.rotation !== undefined) attrs.rot = { key: "rot", value: options.rotation };
        if (Object.keys(attrs).length > 0) {
            this.root.push(new NextAttributeComponent(attrs));
        }

        this.root.push(
            new BuilderElement({
                name: "a:off",
                attributes: {
                    x: { key: "x", value: options.x ?? 0 },
                    y: { key: "y", value: options.y ?? 0 },
                },
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "a:ext",
                attributes: {
                    cx: { key: "cx", value: options.width ?? 0 },
                    cy: { key: "cy", value: options.height ?? 0 },
                },
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "a:chOff",
                attributes: {
                    x: { key: "x", value: options.childOffsetX ?? 0 },
                    y: { key: "y", value: options.childOffsetY ?? 0 },
                },
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "a:chExt",
                attributes: {
                    cx: { key: "cx", value: options.childExtentWidth ?? 0 },
                    cy: { key: "cy", value: options.childExtentHeight ?? 0 },
                },
            }),
        );
    }
}
