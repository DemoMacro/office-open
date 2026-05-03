/**
 * Shape properties for DrawingML pictures.
 *
 * This module provides the shape properties element which defines visual
 * characteristics of a picture including transformation, geometry, fill, outline,
 * effects, and 3D properties.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_ShapeProperties
 *
 * @module
 */
import type { IMediaDataTransformation } from "@file/media";
import type { IContext, IXmlableObject } from "@file/xml-components";
import { XmlComponent } from "@file/xml-components";
import { buildFill, extractBlipFillMedia, type FillOptions } from "@office-open/core/drawingml";

import { createCustomGeometry } from "./custom-geometry/custom-geometry";
import type { CustomGeometryOptions } from "./custom-geometry/custom-geometry";
import { createEffectDag } from "./effects/effect-dag";
import type { EffectDagOptions } from "./effects/effect-dag";
import { createEffectList } from "./effects/effect-list";
import type { EffectListOptions } from "./effects/effect-list";
import { Form } from "./form";
import { createOutline } from "./outline/outline";
import type { OutlineOptions } from "./outline/outline";
import { PresetGeometry } from "./preset-geometry/preset-geometry";
import type { PresetGeometryOptions } from "./preset-geometry/preset-geometry";
import { ShapePropertiesAttributes } from "./shape-properties-attributes";
import { createScene3D } from "./three-d/scene-3d";
import type { Scene3DOptions } from "./three-d/scene-3d";
import { createShape3D } from "./three-d/shape-3d";
import type { Shape3DOptions } from "./three-d/shape-3d";

/**
 * Represents shape properties for a DrawingML picture.
 *
 * This element defines the visual formatting of a picture, including
 * its transform (size, position, rotation, flip), geometry preset,
 * fill, outline, effects, and 3D properties.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_ShapeProperties">
 *   <xsd:sequence>
 *     <xsd:element name="xfrm" type="CT_Transform2D" minOccurs="0"/>
 *     <xsd:group ref="EG_Geometry" minOccurs="0"/>
 *     <xsd:group ref="EG_FillProperties" minOccurs="0"/>
 *     <xsd:element name="ln" type="CT_LineProperties" minOccurs="0"/>
 *     <xsd:group ref="EG_EffectProperties" minOccurs="0"/>
 *     <xsd:element name="scene3d" type="CT_Scene3D" minOccurs="0"/>
 *     <xsd:element name="sp3d" type="CT_Shape3D" minOccurs="0"/>
 *     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * const shapeProps = new ShapeProperties({
 *   element: "pic",
 *   transform: { emus: { x: 914400, y: 914400 } },
 *   effects: {
 *     glow: { rad: 50800, color: { value: "FF0000" } },
 *   },
 * });
 * ```
 */
export class ShapeProperties extends XmlComponent {
    private readonly form: Form;
    private readonly fillOptions?: FillOptions;

    public constructor({
        element,
        customGeometry,
        effectDag,
        effects,
        fill,
        outline,
        presetGeometry,
        scene3d,
        shape3d,
        transform,
    }: {
        readonly element: string;
        readonly outline?: OutlineOptions;
        readonly fill?: FillOptions;
        readonly presetGeometry?: PresetGeometryOptions;
        /** Custom geometry (mutually exclusive with presetGeometry). */
        readonly customGeometry?: CustomGeometryOptions;
        /** Effect DAG container (mutually exclusive with effects). */
        readonly effectDag?: EffectDagOptions;
        readonly effects?: EffectListOptions;
        readonly scene3d?: Scene3DOptions;
        readonly shape3d?: Shape3DOptions;
        readonly transform: IMediaDataTransformation;
    }) {
        super(`${element}:spPr`);

        this.fillOptions = fill;

        this.root.push(
            new ShapePropertiesAttributes({
                bwMode: "auto",
            }),
        );

        this.form = new Form(transform);

        this.root.push(this.form);

        // EG_Geometry: custGeom and prstGeom are mutually exclusive
        if (customGeometry) {
            this.root.push(createCustomGeometry(customGeometry));
        } else {
            this.root.push(new PresetGeometry(presetGeometry));
        }

        if (fill !== undefined) {
            this.root.push(buildFill(fill));
        }

        if (outline) {
            this.root.push(createOutline(outline));
        }

        // EG_EffectProperties: effectDag and effectLst are mutually exclusive
        if (effectDag) {
            this.root.push(createEffectDag(effectDag));
        } else if (effects) {
            this.root.push(createEffectList(effects));
        }

        if (scene3d) {
            this.root.push(createScene3D(scene3d));
        }

        if (shape3d) {
            this.root.push(createShape3D(shape3d));
        }
    }

    public override prepForXml(context: IContext): IXmlableObject | undefined {
        const media = this.fillOptions ? extractBlipFillMedia(this.fillOptions) : undefined;
        if (media) {
            context.file.Media.addImage(media.fileName, {
                data: media.data,
                fileName: media.fileName,
                type: media.type as "png",
                transformation: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
            });
        }
        return super.prepForXml(context);
    }
}
