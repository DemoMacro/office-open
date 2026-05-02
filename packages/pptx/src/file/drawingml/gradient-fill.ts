import { BuilderElement, XmlComponent } from "@file/xml-components";

export interface GradientStopOptions {
    readonly position: number;
    readonly color: string;
}

export interface GradientFillOptions {
    readonly angle?: number;
    readonly stops: readonly GradientStopOptions[];
}

/**
 * a:gradFill — Gradient fill for shapes and backgrounds.
 */
export class GradientFill extends XmlComponent {
    public constructor(options: GradientFillOptions) {
        super("a:gradFill");

        const stops = options.stops.map(
            (stop) =>
                new BuilderElement({
                    name: "a:gs",
                    attributes: { pos: { key: "pos", value: stop.position } },
                    children: [
                        new BuilderElement({
                            name: "a:srgbClr",
                            attributes: {
                                val: { key: "val", value: stop.color.replace("#", "") },
                            },
                        }),
                    ],
                }),
        );

        this.root.push(
            new BuilderElement({
                name: "a:gsLst",
                children: stops,
            }),
        );

        if (options.angle !== undefined) {
            this.root.push(
                new BuilderElement({
                    name: "a:lin",
                    attributes: {
                        ang: { key: "ang", value: options.angle * 60000 },
                        scaled: { key: "scaled", value: 1 },
                    },
                }),
            );
        }
    }
}
