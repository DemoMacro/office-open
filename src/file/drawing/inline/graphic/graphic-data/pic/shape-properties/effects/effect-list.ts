/**
 * Effect list container for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_EffectList, EG_EffectProperties
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import type { GlowEffectOptions } from "./glow";
import { createGlowEffect } from "./glow";
import type { InnerShadowEffectOptions } from "./inner-shdw";
import { createInnerShadowEffect } from "./inner-shdw";
import type { OuterShadowEffectOptions } from "./outer-shdw";
import { createOuterShadowEffect } from "./outer-shdw";
import type { PresetShadowEffectOptions } from "./prst-shdw";
import { createPresetShadowEffect } from "./prst-shdw";
import type { ReflectionEffectOptions } from "./reflection";
import { createReflectionEffect } from "./reflection";
import { createSoftEdgeEffect } from "./soft-edge";

/**
 * Blur effect options.
 */
export interface BlurEffectOptions {
    /** Blur radius in EMUs */
    readonly rad?: number;
    /** Whether to grow the shape boundary */
    readonly grow?: boolean;
}

/**
 * Options for the effect list container.
 *
 * Each property corresponds to an optional child effect element.
 * Only specified effects will be included in the output.
 */
export interface EffectListOptions {
    /** Blur effect */
    readonly blur?: BlurEffectOptions;
    /** Glow effect */
    readonly glow?: GlowEffectOptions;
    /** Inner shadow effect */
    readonly innerShdw?: InnerShadowEffectOptions;
    /** Outer shadow effect */
    readonly outerShdw?: OuterShadowEffectOptions;
    /** Preset shadow effect */
    readonly prstShdw?: PresetShadowEffectOptions;
    /** Reflection effect (pass object for attributes, or true for defaults) */
    readonly reflection?: ReflectionEffectOptions | true;
    /** Soft edge radius in EMUs */
    readonly softEdge?: number;
}

/**
 * Creates a blur effect element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_BlurEffect">
 *   <xsd:attribute name="rad" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="grow" type="xsd:boolean" default="true"/>
 * </xsd:complexType>
 * ```
 */
const createBlurEffect = (options: BlurEffectOptions): XmlComponent => {
    const hasAttributes = options.rad !== undefined || options.grow === false;

    if (!hasAttributes) {
        return new BuilderElement({ name: "a:blur" });
    }

    const attributePayload = {
        ...(options.rad !== undefined && {
            rad: { key: "rad", value: options.rad },
        }),
        ...(options.grow === false && {
            grow: { key: "grow", value: 0 },
        }),
    };

    return new BuilderElement({
        attributes: attributePayload as never,
        name: "a:blur",
    });
};

/**
 * Creates an effect list element (a:effectLst).
 *
 * This is the EG_EffectProperties choice for a flat list of effects.
 * Effects are emitted in XSD order: blur, glow, innerShdw, outerShdw, prstShdw, reflection, softEdge.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_EffectList">
 *   <xsd:sequence>
 *     <xsd:element name="blur" type="CT_BlurEffect" minOccurs="0"/>
 *     <xsd:element name="fillOverlay" type="CT_FillOverlayEffect" minOccurs="0"/>
 *     <xsd:element name="glow" type="CT_GlowEffect" minOccurs="0"/>
 *     <xsd:element name="innerShdw" type="CT_InnerShadowEffect" minOccurs="0"/>
 *     <xsd:element name="outerShdw" type="CT_OuterShadowEffect" minOccurs="0"/>
 *     <xsd:element name="prstShdw" type="CT_PresetShadowEffect" minOccurs="0"/>
 *     <xsd:element name="reflection" type="CT_ReflectionEffect" minOccurs="0"/>
 *     <xsd:element name="softEdge" type="CT_SoftEdgesEffect" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createEffectList = (options: EffectListOptions): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options.blur) {
        children.push(createBlurEffect(options.blur));
    }
    if (options.glow) {
        children.push(createGlowEffect(options.glow));
    }
    if (options.innerShdw) {
        children.push(createInnerShadowEffect(options.innerShdw));
    }
    if (options.outerShdw) {
        children.push(createOuterShadowEffect(options.outerShdw));
    }
    if (options.prstShdw) {
        children.push(createPresetShadowEffect(options.prstShdw));
    }
    if (options.reflection) {
        children.push(
            createReflectionEffect(options.reflection === true ? undefined : options.reflection),
        );
    }
    if (options.softEdge !== undefined) {
        children.push(createSoftEdgeEffect(options.softEdge));
    }

    return new BuilderElement({
        children,
        name: "a:effectLst",
    });
};
