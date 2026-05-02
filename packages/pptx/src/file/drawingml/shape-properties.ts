import { XmlComponent as Xc } from "@file/xml-components";

import { GradientFill } from "./gradient-fill";
import { NoFill } from "./no-fill";
import { Outline } from "./outline";
import { PresetGeometry } from "./preset-geometry";
import { SolidFill } from "./solid-fill";
import type { ITransform2DOptions } from "./transform-2d";
import { Transform2D } from "./transform-2d";

export type ShapeFill = SolidFill | NoFill | GradientFill;

export interface IShapePropertiesOptions extends ITransform2DOptions {
    readonly geometry?: string;
    readonly fill?: ShapeFill;
    readonly outline?: Outline;
}

/**
 * p:spPr — Shape properties (transform, geometry, fill, outline).
 * Uses p: prefix in PresentationML context, though type is a:CT_ShapeProperties.
 */
export class ShapeProperties extends Xc {
    public constructor(options: IShapePropertiesOptions) {
        super("p:spPr");

        if (
            options.x !== undefined ||
            options.y !== undefined ||
            options.width !== undefined ||
            options.height !== undefined ||
            options.flipH !== undefined ||
            options.rotation !== undefined
        ) {
            this.root.push(new Transform2D(options));
        }

        if (options.geometry) {
            this.root.push(new PresetGeometry(options.geometry));
        }

        if (options.fill) {
            this.root.push(options.fill);
        } else {
            this.root.push(new NoFill());
        }

        if (options.outline) {
            this.root.push(options.outline);
        }
    }
}
