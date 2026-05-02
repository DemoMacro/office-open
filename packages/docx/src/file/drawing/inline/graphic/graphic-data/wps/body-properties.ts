/**
 * Text body properties for DrawingML shapes.
 *
 * Provides CT_TextBodyProperties — defines how text is laid out within a shape,
 * including margins, alignment, autofit, vertical text, overflow, columns, and 3D.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_TextBodyProperties
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";
import { OnOffElement } from "@file/xml-components";

import type { Scene3DOptions } from "../pic/shape-properties/three-d/scene-3d";
import { createScene3D } from "../pic/shape-properties/three-d/scene-3d";
import type { Shape3DOptions } from "../pic/shape-properties/three-d/shape-3d";
import { createShape3D } from "../pic/shape-properties/three-d/shape-3d";

// ─── Constants ──────────────────────────────────────────────────────────────

/**
 * Text anchoring type (ST_TextAnchoringType).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_TextAnchoringType">
 *   <xsd:restriction base="xsd:token">
 *     <xsd:enumeration value="t"/>
 *     <xsd:enumeration value="ctr"/>
 *     <xsd:enumeration value="b"/>
 *     <xsd:enumeration value="just"/>
 *     <xsd:enumeration value="dist"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export enum VerticalAnchor {
    TOP = "t",
    CENTER = "ctr",
    BOTTOM = "b",
    JUSTIFY = "just",
    DISTRIBUTED = "dist",
}

/**
 * Text vertical overflow type (ST_TextVertOverflowType).
 */
export const TextVertOverflowType = {
    OVERFLOW: "overflow",
    ELLIPSIS: "ellipsis",
    CLIP: "clip",
} as const;

/**
 * Text horizontal overflow type (ST_TextHorzOverflowType).
 */
export const TextHorzOverflowType = {
    OVERFLOW: "overflow",
    CLIP: "clip",
} as const;

/**
 * Text vertical type (ST_TextVerticalType).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_TextVerticalType">
 *   <xsd:restriction base="xsd:token">
 *     <xsd:enumeration value="horz"/>
 *     <xsd:enumeration value="vert"/>
 *     <xsd:enumeration value="vert270"/>
 *     <xsd:enumeration value="wordArtVert"/>
 *     <xsd:enumeration value="eaVert"/>
 *     <xsd:enumeration value="mongolianVert"/>
 *     <xsd:enumeration value="wordArtVertRtl"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 */
export const TextVerticalType = {
    HORIZONTAL: "horz",
    VERTICAL: "vert",
    VERTICAL_270: "vert270",
    WORD_ART_VERTICAL: "wordArtVert",
    EAST_ASIAN_VERTICAL: "eaVert",
    MONGOLIAN_VERTICAL: "mongolianVert",
    WORD_ART_VERTICAL_RTL: "wordArtVertRtl",
} as const;

/**
 * Text body wrapping type (ST_TextWrappingType).
 *
 * This is different from text wrapping around shapes (ST_WrapText).
 */
export const TextBodyWrappingType = {
    NONE: "none",
    SQUARE: "square",
} as const;

// ─── Sub-element Options ────────────────────────────────────────────────────

/**
 * Normal autofit options (CT_TextNormalAutofit).
 *
 * Scales text and line spacing to fit the shape.
 */
export interface NormalAutofitOptions {
    /** Font scale percentage (e.g., 100000 = 100%) */
    readonly fontScale?: number;
    /** Line spacing reduction percentage */
    readonly lnSpcReduction?: number;
}

/**
 * Preset text warp options (CT_PresetTextShape).
 */
export interface PresetTextShapeOptions {
    /** Preset shape type (e.g., "textArchUp", "textCircle") */
    readonly preset: string;
    /** Optional adjustment values */
    readonly adjustments?: readonly { readonly name: string; readonly formula: string }[];
}

/**
 * Flat text options (CT_FlatText).
 */
export interface FlatTextOptions {
    /** Z-offset in EMUs */
    readonly z?: number;
}

// ─── Main Options ───────────────────────────────────────────────────────────

/**
 * Options for text body properties (CT_TextBodyProperties).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_TextBodyProperties">
 *   <xsd:sequence>
 *     <xsd:element name="prstTxWarp" type="CT_PresetTextShape" minOccurs="0"/>
 *     <xsd:group ref="EG_TextAutofit" minOccurs="0"/>
 *     <xsd:element name="scene3d" type="CT_Scene3D" minOccurs="0"/>
 *     <xsd:group ref="EG_Text3D" minOccurs="0"/>
 *   </xsd:sequence>
 *   <!-- 19 attributes -->
 * </xsd:complexType>
 * ```
 */
export interface IBodyPropertiesOptions {
    // ── Attributes ──

    /** Text rotation angle in 60,000ths of a degree */
    readonly rotation?: number;
    /** Whether to use spcFirstLastPara behavior */
    readonly spcFirstLastPara?: boolean;
    /** Vertical text overflow behavior */
    readonly vertOverflow?: (typeof TextVertOverflowType)[keyof typeof TextVertOverflowType];
    /** Horizontal text overflow behavior */
    readonly horzOverflow?: (typeof TextHorzOverflowType)[keyof typeof TextHorzOverflowType];
    /** Text vertical direction */
    readonly vert?: (typeof TextVerticalType)[keyof typeof TextVerticalType];
    /** Text wrapping type */
    readonly wrap?: (typeof TextBodyWrappingType)[keyof typeof TextBodyWrappingType];
    /** Left inset in EMUs */
    readonly lIns?: number;
    /** Top inset in EMUs */
    readonly tIns?: number;
    /** Right inset in EMUs */
    readonly rIns?: number;
    /** Bottom inset in EMUs */
    readonly bIns?: number;
    /** Number of text columns (1-16) */
    readonly numCol?: number;
    /** Spacing between columns in EMUs */
    readonly spcCol?: number;
    /** Whether columns are right-to-left */
    readonly rtlCol?: boolean;
    /** Whether text is from WordArt */
    readonly fromWordArt?: boolean;
    /** Text anchor position */
    readonly anchor?: (typeof VerticalAnchor)[keyof typeof VerticalAnchor];
    /** Whether to anchor at center */
    readonly anchorCtr?: boolean;
    /** Whether to force anti-aliasing */
    readonly forceAA?: boolean;
    /** Whether text is upright (default false) */
    readonly upright?: boolean;
    /** Whether to use compatible line spacing */
    readonly compatLnSpc?: boolean;

    // ── Convenience aliases (backward compatible) ──

    /** Vertical anchor position (alias for `anchor`) */
    readonly verticalAnchor?: VerticalAnchor;
    /** Margins shorthand */
    readonly margins?: {
        readonly top?: number;
        readonly bottom?: number;
        readonly left?: number;
        readonly right?: number;
    };

    // ── Child elements ──

    /** Preset text warp shape */
    readonly prstTxWarp?: PresetTextShapeOptions;
    /** Disable autofit (EG_TextAutofit choice) */
    readonly noAutoFit?: boolean;
    /** Normal autofit (EG_TextAutofit choice) */
    readonly normAutofit?: NormalAutofitOptions;
    /** Shape autofit (EG_TextAutofit choice) */
    readonly spAutoFit?: boolean;
    /** 3D scene */
    readonly scene3d?: Scene3DOptions;
    /** 3D shape properties (EG_Text3D choice) */
    readonly sp3d?: Shape3DOptions;
    /** Flat text (EG_Text3D choice) */
    readonly flatTx?: FlatTextOptions;
}

// ─── Internal Helpers ───────────────────────────────────────────────────────

const buildOptionalAttributes = (
    options: Record<string, unknown>,
): Record<string, { readonly key: string; readonly value: string | number | boolean }> => {
    const attrs: Record<
        string,
        { readonly key: string; readonly value: string | number | boolean }
    > = {};
    for (const [key, value] of Object.entries(options)) {
        if (value !== undefined) {
            attrs[key] = { key, value: value as string | number | boolean };
        }
    }
    return attrs;
};

// ─── Preset Text Warp ───────────────────────────────────────────────────────

const createPresetTextShape = (options: PresetTextShapeOptions): XmlComponent => {
    const children: XmlComponent[] = [];
    if (options.adjustments) {
        for (const adj of options.adjustments) {
            children.push(
                new BuilderElement<{ readonly name: string; readonly fmla: string }>({
                    name: "a:avLst",
                    children: [
                        new BuilderElement<{ readonly name: string; readonly fmla: string }>({
                            name: "a:gd",
                            attributes: {
                                name: { key: "name", value: adj.name },
                                fmla: { key: "fmla", value: adj.formula },
                            },
                        }),
                    ],
                }),
            );
        }
    }

    return new BuilderElement<{ readonly prst: string }>({
        name: "a:prstTxWarp",
        attributes: { prst: { key: "prst", value: options.preset } },
        children: children.length > 0 ? children : undefined,
    });
};

// ─── Main Factory ───────────────────────────────────────────────────────────

/**
 * Creates a text body properties element (wps:bodyPr).
 *
 * Supports all 19 CT_TextBodyProperties attributes and all child elements
 * including preset text warp, autofit variants, 3D scene, and text 3D.
 *
 * @example
 * ```typescript
 * // Simple margins and anchor
 * createBodyProperties({
 *   margins: { top: 100, bottom: 100, left: 200, right: 200 },
 *   verticalAnchor: VerticalAnchor.CENTER,
 * });
 *
 * // Full options with vertical text and autofit
 * createBodyProperties({
 *   vert: TextVerticalType.VERTICAL,
 *   wrap: TextBodyWrappingType.NONE,
 *   normAutofit: { fontScale: 80000 },
 *   numCol: 2,
 *   anchor: VerticalAnchor.TOP,
 * });
 * ```
 */
export const createBodyProperties = (options: IBodyPropertiesOptions = {}): XmlComponent => {
    // Resolve anchor (direct `anchor` takes precedence over `verticalAnchor`)
    const anchor = options.anchor ?? options.verticalAnchor;

    // Resolve margins
    const lIns = options.lIns ?? options.margins?.left;
    const tIns = options.tIns ?? options.margins?.top;
    const rIns = options.rIns ?? options.margins?.right;
    const bIns = options.bIns ?? options.margins?.bottom;

    const attrs = buildOptionalAttributes({
        rot: options.rotation,
        spcFirstLastPara: options.spcFirstLastPara,
        vertOverflow: options.vertOverflow,
        horzOverflow: options.horzOverflow,
        vert: options.vert,
        wrap: options.wrap,
        lIns,
        tIns,
        rIns,
        bIns,
        numCol: options.numCol,
        spcCol: options.spcCol,
        rtlCol: options.rtlCol,
        fromWordArt: options.fromWordArt,
        anchor,
        anchorCtr: options.anchorCtr,
        forceAA: options.forceAA,
        upright: options.upright,
        compatLnSpc: options.compatLnSpc,
    });

    // Build children in XSD sequence order
    const children: XmlComponent[] = [];

    // a:prstTxWarp
    if (options.prstTxWarp) {
        children.push(createPresetTextShape(options.prstTxWarp));
    }

    // EG_TextAutofit (mutually exclusive)
    if (options.noAutoFit) {
        children.push(new OnOffElement("a:noAutofit", options.noAutoFit));
    } else if (options.normAutofit) {
        const normAttrs = buildOptionalAttributes({
            fontScale: options.normAutofit.fontScale,
            lnSpcReduction: options.normAutofit.lnSpcReduction,
        });
        children.push(
            new BuilderElement({
                name: "a:normAutofit",
                attributes: Object.keys(normAttrs).length > 0 ? normAttrs : undefined,
            }),
        );
    } else if (options.spAutoFit) {
        children.push(new OnOffElement("a:spAutoFit", options.spAutoFit));
    }

    // a:scene3d
    if (options.scene3d) {
        children.push(createScene3D(options.scene3d));
    }

    // EG_Text3D (mutually exclusive)
    if (options.sp3d) {
        children.push(createShape3D(options.sp3d));
    } else if (options.flatTx) {
        const flatAttrs = buildOptionalAttributes({ z: options.flatTx.z });
        children.push(
            new BuilderElement({
                name: "a:flatTx",
                attributes: Object.keys(flatAttrs).length > 0 ? flatAttrs : undefined,
            }),
        );
    }

    return new BuilderElement({
        name: "wps:bodyPr",
        attributes: Object.keys(attrs).length > 0 ? (attrs as never) : undefined,
        children: children.length > 0 ? children : undefined,
    });
};
