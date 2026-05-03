import type { IMediaDataTransformation } from "@file/media";
import type { Paragraph } from "@file/paragraph";
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";
import type { FillOptions } from "@office-open/core/drawingml";

import type { CustomGeometryOptions } from "../pic/shape-properties/custom-geometry/custom-geometry";
import type { EffectDagOptions } from "../pic/shape-properties/effects/effect-dag";
import type { EffectListOptions } from "../pic/shape-properties/effects/effect-list";
import type { OutlineOptions } from "../pic/shape-properties/outline/outline";
import { ShapeProperties } from "../pic/shape-properties/shape-properties";
import type { Scene3DOptions } from "../pic/shape-properties/three-d/scene-3d";
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
    readonly fill?: FillOptions;
    readonly customGeometry?: CustomGeometryOptions;
    readonly effectDag?: EffectDagOptions;
    readonly effects?: EffectListOptions;
    readonly scene3d?: Scene3DOptions;
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
                customGeometry: options.customGeometry,
                effectDag: options.effectDag,
                effects: options.effects,
                fill: options.fill,
                outline: options.outline,
                scene3d: options.scene3d,
                shape3d: options.shape3d,
                transform: options.transformation,
            }),
            createWpsTextBox(options.children),
            createBodyProperties(options.bodyProperties),
        ],
        name: "wps:wsp",
    });
