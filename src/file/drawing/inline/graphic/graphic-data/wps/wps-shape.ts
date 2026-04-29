import type { IMediaDataTransformation } from "@file/media";
import type { Paragraph } from "@file/paragraph";
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import type { EffectListOptions } from "../pic/shape-properties/effects/effect-list";
import type { GradientFillOptions } from "../pic/shape-properties/fill/gradient-fill";
import type { PatternFillOptions } from "../pic/shape-properties/fill/pattern-fill";
import type { OutlineOptions } from "../pic/shape-properties/outline/outline";
import type { SolidFillOptions } from "../pic/shape-properties/outline/solid-fill";
import { ShapeProperties } from "../pic/shape-properties/shape-properties";
import type { Shape3DOptions } from "../pic/shape-properties/three-d/shape-3d";
import { createBodyProperties } from "./body-properties";
import type { IBodyPropertiesOptions } from "./body-properties";
import { createNonVisualShapeProperties } from "./non-visual-shape-properties";
import type { INonVisualShapePropertiesOptions } from "./non-visual-shape-properties";
import { createWpsTextBox } from "./wps-text-box";

export interface WpsShapeCoreOptions {
    readonly children: readonly Paragraph[];
    readonly nonVisualProperties?: INonVisualShapePropertiesOptions;
    readonly bodyProperties?: IBodyPropertiesOptions;
    readonly outline?: OutlineOptions;
    readonly solidFill?: SolidFillOptions;
    readonly gradientFill?: GradientFillOptions;
    readonly patternFill?: PatternFillOptions;
    readonly noFill?: boolean;
    readonly effects?: EffectListOptions;
    readonly shape3d?: Shape3DOptions;
}

export type WpsShapeOptions = WpsShapeCoreOptions & {
    readonly transformation: IMediaDataTransformation;
};

export const createWpsShape = (options: WpsShapeOptions): XmlComponent =>
    new BuilderElement({
        children: [
            createNonVisualShapeProperties(options.nonVisualProperties),
            new ShapeProperties({
                element: "wps",
                effects: options.effects,
                gradientFill: options.gradientFill,
                noFill: options.noFill,
                outline: options.outline,
                patternFill: options.patternFill,
                shape3d: options.shape3d,
                solidFill: options.solidFill,
                transform: options.transformation,
            }),
            createWpsTextBox(options.children),
            createBodyProperties(options.bodyProperties),
        ],
        name: "wps:wsp",
    });
