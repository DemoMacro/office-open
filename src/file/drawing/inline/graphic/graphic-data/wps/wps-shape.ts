import type { IMediaDataTransformation } from "@file/media";
import type { Paragraph } from "@file/paragraph";
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import type { OutlineOptions } from "../pic/shape-properties/outline/outline";
import type { SolidFillOptions } from "../pic/shape-properties/outline/solid-fill";
import { ShapeProperties } from "../pic/shape-properties/shape-properties";
import { createBodyProperties } from "./body-properties";
import type { IBodyPropertiesOptions } from "./body-properties";
import { createNonVisualShapeProperties } from "./non-visual-shape-properties";
import type { INonVisualShapePropertiesOptions } from "./non-visual-shape-properties";
import { createWpsTextBox } from "./wps-text-box";

export interface WpsShapeCoreOptions {
    readonly children: readonly Paragraph[];
    readonly nonVisualProperties?: INonVisualShapePropertiesOptions;
    readonly bodyProperties?: IBodyPropertiesOptions;
}

export type WpsShapeOptions = WpsShapeCoreOptions & {
    readonly transformation: IMediaDataTransformation;
    readonly outline?: OutlineOptions;
    readonly solidFill?: SolidFillOptions;
};

export const createWpsShape = (options: WpsShapeOptions): XmlComponent =>
    new BuilderElement({
        children: [
            createNonVisualShapeProperties(options.nonVisualProperties),
            new ShapeProperties({
                element: "wps",
                outline: options.outline,
                solidFill: options.solidFill,
                transform: options.transformation,
            }),
            createWpsTextBox(options.children),
            createBodyProperties(options.bodyProperties),
        ],
        name: "wps:wsp",
    });
