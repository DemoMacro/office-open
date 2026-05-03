import type { FillOptions } from "@file/drawingml/fill";
import { buildFill } from "@file/drawingml/fill";
import { NonVisualShapeProperties } from "@file/drawingml/non-visual-shape-props";
import { createOutlineCompat, type OutlineOptions } from "@file/drawingml/outline";
import { PresetGeometry } from "@file/drawingml/preset-geometry";
import { BuilderElement, XmlComponent as Xc } from "@file/xml-components";
import { pixelsToEmus } from "@util/types";

/**
 * p:sp — A line shape on a slide.
 */
export class LineShape extends Xc {
    private static nextId = 2;
    private readonly shapeId: number;

    public constructor(options: ILineShapeOptions = {}) {
        super("p:sp");

        const id = options.id ?? LineShape.nextId++;
        this.shapeId = id;
        const name = options.name ?? `Line ${id}`;

        const x1 = pixelsToEmus(options.x1 ?? 0);
        const y1 = pixelsToEmus(options.y1 ?? 0);
        const x2 = pixelsToEmus(options.x2 ?? 100);
        const y2 = pixelsToEmus(options.y2 ?? 100);

        this.root.push(
            new BuilderElement({
                name: "p:nvSpPr",
                children: [
                    new BuilderElement({
                        name: "p:cNvPr",
                        attributes: {
                            id: { key: "id", value: id },
                            name: { key: "name", value: name },
                        },
                    }),
                    new NonVisualShapeProperties(),
                    new BuilderElement({ name: "p:nvPr" }),
                ],
            }),
        );

        const offX = Math.min(x1, x2);
        const offY = Math.min(y1, y2);

        const xfrmAttrs:
            | Record<string, { readonly key: string; readonly value: string | number }>
            | undefined =
            x1 > x2 || y1 > y2
                ? (() => {
                      const a: Record<
                          string,
                          { readonly key: string; readonly value: string | number }
                      > = {};
                      if (x1 > x2) a.flipH = { key: "flipH", value: 1 };
                      if (y1 > y2) a.flipV = { key: "flipV", value: 1 };
                      return a;
                  })()
                : undefined;

        const spPrChildren: BuilderElement<{}>[] = [
            new BuilderElement({
                name: "a:xfrm",
                children: [
                    new BuilderElement({
                        name: "a:off",
                        attributes: { x: { key: "x", value: offX }, y: { key: "y", value: offY } },
                    }),
                    new BuilderElement({
                        name: "a:ext",
                        attributes: {
                            cx: { key: "cx", value: Math.abs(x2 - x1) },
                            cy: { key: "cy", value: Math.abs(y2 - y1) },
                        },
                    }),
                ],
                attributes: xfrmAttrs,
            }),
            new PresetGeometry("line"),
        ];

        if (options.fill !== undefined) {
            spPrChildren.push(buildFill(options.fill));
        }
        if (options.outline) {
            spPrChildren.push(createOutlineCompat(options.outline));
        }

        this.root.push(
            new BuilderElement({
                name: "p:spPr",
                children: spPrChildren,
            }),
        );

        this.root.push(
            new BuilderElement({
                name: "p:txBody",
                children: [
                    new BuilderElement({
                        name: "a:bodyPr",
                        attributes: { wrap: { key: "wrap", value: "square" } },
                    }),
                    new BuilderElement({ name: "a:lstStyle" }),
                    new BuilderElement({ name: "a:p" }),
                ],
            }),
        );
    }

    public get ShapeId(): number {
        return this.shapeId;
    }
}

export interface ILineShapeOptions {
    readonly id?: number;
    readonly name?: string;
    readonly x1?: number;
    readonly y1?: number;
    readonly x2?: number;
    readonly y2?: number;
    readonly fill?: FillOptions;
    readonly outline?: OutlineOptions;
}

export type ArrowheadType = "triangle" | "stealth" | "diamond" | "oval" | "open" | "none";

/**
 * p:cxnSp — A connector shape on a slide (line with optional arrowheads).
 */
export class ConnectorShape extends Xc {
    private static nextId = 2;
    private readonly shapeId: number;

    public constructor(options: IConnectorShapeOptions = {}) {
        super("p:cxnSp");

        const id = options.id ?? ConnectorShape.nextId++;
        this.shapeId = id;
        const name = options.name ?? `Connector ${id}`;

        const x1 = pixelsToEmus(options.x1 ?? 0);
        const y1 = pixelsToEmus(options.y1 ?? 0);
        const x2 = pixelsToEmus(options.x2 ?? 100);
        const y2 = pixelsToEmus(options.y2 ?? 100);

        this.root.push(
            new BuilderElement({
                name: "p:nvCxnSpPr",
                children: [
                    new BuilderElement({
                        name: "p:cNvPr",
                        attributes: {
                            id: { key: "id", value: id },
                            name: { key: "name", value: name },
                        },
                    }),
                    new BuilderElement({ name: "p:cNvCxnSpPr" }),
                    new BuilderElement({ name: "p:nvPr" }),
                ],
            }),
        );

        const offX = Math.min(x1, x2);
        const offY = Math.min(y1, y2);

        const xfrmAttrs:
            | Record<string, { readonly key: string; readonly value: string | number }>
            | undefined =
            x1 > x2 || y1 > y2
                ? (() => {
                      const a: Record<
                          string,
                          { readonly key: string; readonly value: string | number }
                      > = {};
                      if (x1 > x2) a.flipH = { key: "flipH", value: 1 };
                      if (y1 > y2) a.flipV = { key: "flipV", value: 1 };
                      return a;
                  })()
                : undefined;

        const spPrChildren: BuilderElement<{}>[] = [
            new BuilderElement({
                name: "a:xfrm",
                children: [
                    new BuilderElement({
                        name: "a:off",
                        attributes: { x: { key: "x", value: offX }, y: { key: "y", value: offY } },
                    }),
                    new BuilderElement({
                        name: "a:ext",
                        attributes: {
                            cx: { key: "cx", value: Math.abs(x2 - x1) },
                            cy: { key: "cy", value: Math.abs(y2 - y1) },
                        },
                    }),
                ],
                attributes: xfrmAttrs,
            }),
            new PresetGeometry("line"),
        ];

        if (options.fill !== undefined) {
            spPrChildren.push(buildFill(options.fill));
        }

        // Arrowheads + outline: headEnd/tailEnd must be inside a:ln per XSD
        const hasArrowheads = options.beginArrowhead || options.endArrowhead;
        if (options.outline || hasArrowheads) {
            spPrChildren.push(
                createOutlineCompat(
                    options.outline ?? {},
                    hasArrowheads
                        ? {
                              beginType: options.beginArrowhead,
                              endType: options.endArrowhead,
                              width: options.arrowheadWidth,
                              length: options.arrowheadLength,
                          }
                        : undefined,
                ),
            );
        }

        this.root.push(
            new BuilderElement({
                name: "p:spPr",
                children: spPrChildren,
            }),
        );
    }

    public get ShapeId(): number {
        return this.shapeId;
    }
}

export interface IConnectorShapeOptions {
    readonly id?: number;
    readonly name?: string;
    readonly x1?: number;
    readonly y1?: number;
    readonly x2?: number;
    readonly y2?: number;
    readonly fill?: FillOptions;
    readonly outline?: OutlineOptions;
    readonly beginArrowhead?: ArrowheadType;
    readonly endArrowhead?: ArrowheadType;
    readonly arrowheadWidth?: "sm" | "med" | "lg";
    readonly arrowheadLength?: "sm" | "med" | "lg";
}
