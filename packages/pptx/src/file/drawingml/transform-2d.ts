import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";

export interface ITransform2DOptions {
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly flipH?: boolean;
    readonly rotation?: number;
}

/**
 * a:xfrm — 2D transform for shapes (position + size in EMUs).
 */
export class Transform2D extends XmlComponent {
    public constructor(options: ITransform2DOptions) {
        super("a:xfrm");

        const attrs: Record<
            string,
            { readonly key: string; readonly value: string | boolean | number | undefined }
        > = {};
        if (options.flipH !== undefined) attrs.flipH = { key: "flipH", value: options.flipH };
        if (options.rotation !== undefined) attrs.rot = { key: "rot", value: options.rotation };
        if (Object.keys(attrs).length > 0) {
            this.root.push(new NextAttributeComponent(attrs));
        }

        if (options.x !== undefined || options.y !== undefined) {
            this.root.push(
                new BuilderElement({
                    name: "a:off",
                    attributes: {
                        x: { key: "x", value: options.x ?? 0 },
                        y: { key: "y", value: options.y ?? 0 },
                    },
                }),
            );
        }

        if (options.width !== undefined || options.height !== undefined) {
            this.root.push(
                new BuilderElement({
                    name: "a:ext",
                    attributes: {
                        cx: { key: "cx", value: options.width ?? 0 },
                        cy: { key: "cy", value: options.height ?? 0 },
                    },
                }),
            );
        }
    }
}
