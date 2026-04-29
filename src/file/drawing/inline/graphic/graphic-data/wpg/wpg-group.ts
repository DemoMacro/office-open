import type { IMediaDataTransformation } from "@file/media";
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import { createEffectList } from "../pic/shape-properties/effects/effect-list";
import type { EffectListOptions } from "../pic/shape-properties/effects/effect-list";
import { Extents } from "../pic/shape-properties/form/extents/extents";
import { Offset } from "../pic/shape-properties/form/offset/off";
import { createNoFill } from "../pic/shape-properties/outline/no-fill";
import { createSolidFill } from "../pic/shape-properties/outline/solid-fill";
import type { SolidFillOptions } from "../pic/shape-properties/outline/solid-fill";
import { createPatternFill } from "../pic/shape-properties/fill/pattern-fill";
import type { PatternFillOptions } from "../pic/shape-properties/fill/pattern-fill";

export type GroupChild = XmlComponent;

/**
 * Child coordinate offset (CT_Point2D).
 */
export interface IChildOffset {
    readonly x: number;
    readonly y: number;
}

/**
 * Child coordinate extent (CT_PositiveSize2D).
 */
export interface IChildExtent {
    readonly cx: number;
    readonly cy: number;
}

export interface WpgGroupCoreOptions {
    readonly children: readonly GroupChild[];
}

export type WpgGroupOptions = WpgGroupCoreOptions & {
    readonly transformation: IMediaDataTransformation;
    /** Child coordinate offset (chOff) */
    readonly chOff?: IChildOffset;
    /** Child coordinate extent (chExt) */
    readonly chExt?: IChildExtent;
    /** Group fill */
    readonly solidFill?: SolidFillOptions;
    /** Group pattern fill */
    readonly patternFill?: PatternFillOptions;
    /** Group no fill */
    readonly noFill?: boolean;
    /** Group effects */
    readonly effects?: EffectListOptions;
};

/**
 * Creates a group transform element (a:xfrm) with optional chOff/chExt.
 *
 * This uses CT_GroupTransform2D which extends CT_Transform2D with
 * child offset and extent elements.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_GroupTransform2D">
 *   <xsd:sequence>
 *     <xsd:element name="off" type="CT_Point2D" minOccurs="0"/>
 *     <xsd:element name="ext" type="CT_PositiveSize2D" minOccurs="0"/>
 *     <xsd:element name="chOff" type="CT_Point2D" minOccurs="0"/>
 *     <xsd:element name="chExt" type="CT_PositiveSize2D" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="rot" type="ST_Angle" default="0"/>
 *   <xsd:attribute name="flipH" type="xsd:boolean" default="false"/>
 *   <xsd:attribute name="flipV" type="xsd:boolean" default="false"/>
 * </xsd:complexType>
 * ```
 */
const createGroupForm = (
    transform: IMediaDataTransformation,
    chOff?: IChildOffset,
    chExt?: IChildExtent,
): XmlComponent => {
    const children: XmlComponent[] = [
        new Offset(transform.offset?.emus?.x, transform.offset?.emus?.y),
        new Extents(transform.emus.x, transform.emus.y),
    ];

    if (chOff) {
        children.push(
            new BuilderElement<{ readonly x: number; readonly y: number }>({
                attributes: {
                    x: { key: "x", value: chOff.x },
                    y: { key: "y", value: chOff.y },
                },
                name: "a:chOff",
            }),
        );
    }

    if (chExt) {
        children.push(
            new BuilderElement<{ readonly cx: number; readonly cy: number }>({
                attributes: {
                    cx: { key: "cx", value: chExt.cx },
                    cy: { key: "cy", value: chExt.cy },
                },
                name: "a:chExt",
            }),
        );
    }

    const hasAttributes =
        transform.flip?.horizontal !== undefined ||
        transform.flip?.vertical !== undefined ||
        transform.rotation !== undefined;

    const attributes = hasAttributes
        ? {
              ...(transform.flip?.horizontal !== undefined && {
                  flipH: { key: "flipH", value: transform.flip.horizontal },
              }),
              ...(transform.flip?.vertical !== undefined && {
                  flipV: { key: "flipV", value: transform.flip.vertical },
              }),
              ...(transform.rotation !== undefined && {
                  rot: { key: "rot", value: transform.rotation },
              }),
          }
        : undefined;

    return new BuilderElement({
        attributes: attributes as never,
        children,
        name: "a:xfrm",
    });
};

const createGroupProperties = (options: WpgGroupOptions): XmlComponent => {
    const children: XmlComponent[] = [
        createGroupForm(options.transformation, options.chOff, options.chExt),
    ];

    if (options.noFill) {
        children.push(createNoFill());
    } else if (options.solidFill) {
        children.push(createSolidFill(options.solidFill));
    } else if (options.patternFill) {
        children.push(createPatternFill(options.patternFill));
    }

    if (options.effects) {
        children.push(createEffectList(options.effects));
    }

    return new BuilderElement({
        children,
        name: "wpg:grpSpPr",
    });
};

const createNonVisualGroupProperties = (): XmlComponent =>
    new BuilderElement({
        name: "wpg:cNvGrpSpPr",
    });

export const createWpgGroup = (options: WpgGroupOptions): XmlComponent =>
    new BuilderElement({
        children: [
            createNonVisualGroupProperties(),
            createGroupProperties(options),
            ...options.children,
        ],
        name: "wpg:wgp",
    });
