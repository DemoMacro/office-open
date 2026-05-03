/**
 * Pattern fill element for DrawingML shapes.
 *
 * This module provides pattern fill support with preset patterns and
 * optional foreground/background colors.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_PatternFillProperties
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";
import type { XmlComponent } from "../../xml-components";
import type { SolidFillOptions } from "../color/solid-fill";
import { createColorElement } from "../color/solid-fill";

/**
 * Preset pattern values for pattern fill.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_PresetPatternVal">
 *   <xsd:restriction base="xsd:token">
 *     <xsd:enumeration value="pct5"/> ... <xsd:enumeration value="zigZag"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const PresetPattern = {
    /** 5% pattern */
    PCT5: "pct5",
    /** 10% pattern */
    PCT10: "pct10",
    /** 20% pattern */
    PCT20: "pct20",
    /** 25% pattern */
    PCT25: "pct25",
    /** 30% pattern */
    PCT30: "pct30",
    /** 40% pattern */
    PCT40: "pct40",
    /** 50% pattern */
    PCT50: "pct50",
    /** 60% pattern */
    PCT60: "pct60",
    /** 70% pattern */
    PCT70: "pct70",
    /** 75% pattern */
    PCT75: "pct75",
    /** 80% pattern */
    PCT80: "pct80",
    /** 90% pattern */
    PCT90: "pct90",
    /** Horizontal lines */
    HORZ: "horz",
    /** Vertical lines */
    VERT: "vert",
    /** Light horizontal lines */
    LT_HORZ: "ltHorz",
    /** Light vertical lines */
    LT_VERT: "ltVert",
    /** Dark horizontal lines */
    DK_HORZ: "dkHorz",
    /** Dark vertical lines */
    DK_VERT: "dkVert",
    /** Narrow horizontal lines */
    NAR_HORZ: "narHorz",
    /** Narrow vertical lines */
    NAR_VERT: "narVert",
    /** Dashed horizontal lines */
    DASH_HORZ: "dashHorz",
    /** Dashed vertical lines */
    DASH_VERT: "dashVert",
    /** Cross pattern (+) */
    CROSS: "cross",
    /** Downward diagonal lines (\) */
    DN_DIAG: "dnDiag",
    /** Upward diagonal lines (/) */
    UP_DIAG: "upDiag",
    /** Light downward diagonal lines */
    LT_DN_DIAG: "ltDnDiag",
    /** Light upward diagonal lines */
    LT_UP_DIAG: "ltUpDiag",
    /** Dark downward diagonal lines */
    DK_DN_DIAG: "dkDnDiag",
    /** Dark upward diagonal lines */
    DK_UP_DIAG: "dkUpDiag",
    /** Wide downward diagonal lines */
    WD_DN_DIAG: "wdDnDiag",
    /** Wide upward diagonal lines */
    WD_UP_DIAG: "wdUpDiag",
    /** Dashed downward diagonal lines */
    DASH_DN_DIAG: "dashDnDiag",
    /** Dashed upward diagonal lines */
    DASH_UP_DIAG: "dashUpDiag",
    /** Diagonal cross pattern (X) */
    DIAG_CROSS: "diagCross",
    /** Small checkerboard */
    SM_CHECK: "smCheck",
    /** Large checkerboard */
    LG_CHECK: "lgCheck",
    /** Small grid */
    SM_GRID: "smGrid",
    /** Large grid */
    LG_GRID: "lgGrid",
    /** Dot grid */
    DOT_GRID: "dotGrid",
    /** Small confetti */
    SM_CONFETTI: "smConfetti",
    /** Large confetti */
    LG_CONFETTI: "lgConfetti",
    /** Horizontal brick pattern */
    HORZ_BRICK: "horzBrick",
    /** Diagonal brick pattern */
    DIAG_BRICK: "diagBrick",
    /** Solid diamond */
    SOLID_DMND: "solidDmnd",
    /** Open diamond */
    OPEN_DMND: "openDmnd",
    /** Dotted diamond */
    DOT_DMND: "dotDmnd",
    /** Plaid pattern */
    PLAID: "plaid",
    /** Sphere pattern */
    SPHERE: "sphere",
    /** Weave pattern */
    WEAVE: "weave",
    /** Divot pattern */
    DIVOT: "divot",
    /** Shingle pattern */
    SHINGLE: "shingle",
    /** Wave pattern */
    WAVE: "wave",
    /** Trellis pattern */
    TRELLIS: "trellis",
    /** Zigzag pattern */
    ZIG_ZAG: "zigZag",
} as const;

/**
 * Options for pattern fill.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_PatternFillProperties">
 *   <xsd:sequence>
 *     <xsd:element name="fgClr" type="CT_Color" minOccurs="0"/>
 *     <xsd:element name="bgClr" type="CT_Color" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="prst" type="ST_PresetPatternVal" use="optional"/>
 * </xsd:complexType>
 * ```
 */
export interface PatternFillOptions {
    /** Preset pattern type */
    readonly pattern: (typeof PresetPattern)[keyof typeof PresetPattern];
    /** Foreground color */
    readonly foregroundColor?: SolidFillOptions;
    /** Background color */
    readonly backgroundColor?: SolidFillOptions;
}

/**
 * Creates a pattern fill element (a:pattFill).
 *
 * Specifies a pattern fill using preset patterns with optional
 * foreground and background colors.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_PatternFillProperties">
 *   <xsd:sequence>
 *     <xsd:element name="fgClr" type="CT_Color" minOccurs="0"/>
 *     <xsd:element name="bgClr" type="CT_Color" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="prst" type="ST_PresetPatternVal" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Simple crosshatch pattern
 * createPatternFill({ pattern: PresetPattern.CROSS });
 * // Pattern with foreground color
 * createPatternFill({
 *   pattern: PresetPattern.DIAG_CROSS,
 *   foregroundColor: { value: "FF0000" },
 * });
 * // Pattern with foreground and background colors
 * createPatternFill({
 *   pattern: PresetPattern.HORZ,
 *   foregroundColor: { value: "0000FF" },
 *   backgroundColor: { value: "FFFF00" },
 * });
 * ```
 */
export const createPatternFill = (options: PatternFillOptions): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options.foregroundColor) {
        children.push(
            new BuilderElement({
                children: [createColorElement(options.foregroundColor)],
                name: "a:fgClr",
            }),
        );
    }

    if (options.backgroundColor) {
        children.push(
            new BuilderElement({
                children: [createColorElement(options.backgroundColor)],
                name: "a:bgClr",
            }),
        );
    }

    return new BuilderElement<{ readonly prst?: string }>({
        attributes: {
            prst: { key: "prst", value: options.pattern },
        },
        children,
        name: "a:pattFill",
    });
};
