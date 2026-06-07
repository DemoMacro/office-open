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
import { xsdPattern } from "../../xsd-mappings";
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
  PCT5: "percent5",
  /** 10% pattern */
  PCT10: "percent10",
  /** 20% pattern */
  PCT20: "percent20",
  /** 25% pattern */
  PCT25: "percent25",
  /** 30% pattern */
  PCT30: "percent30",
  /** 40% pattern */
  PCT40: "percent40",
  /** 50% pattern */
  PCT50: "percent50",
  /** 60% pattern */
  PCT60: "percent60",
  /** 70% pattern */
  PCT70: "percent70",
  /** 75% pattern */
  PCT75: "percent75",
  /** 80% pattern */
  PCT80: "percent80",
  /** 90% pattern */
  PCT90: "percent90",
  /** Horizontal lines */
  HORZ: "horizontal",
  /** Vertical lines */
  VERT: "vertical",
  /** Light horizontal lines */
  LT_HORZ: "lightHorizontal",
  /** Light vertical lines */
  LT_VERT: "lightVertical",
  /** Dark horizontal lines */
  DK_HORZ: "darkHorizontal",
  /** Dark vertical lines */
  DK_VERT: "darkVertical",
  /** Narrow horizontal lines */
  NAR_HORZ: "narrowHorizontal",
  /** Narrow vertical lines */
  NAR_VERT: "narrowVertical",
  /** Dashed horizontal lines */
  DASH_HORZ: "dashedHorizontal",
  /** Dashed vertical lines */
  DASH_VERT: "dashedVertical",
  /** Cross pattern (+) */
  CROSS: "cross",
  /** Downward diagonal lines (\) */
  DN_DIAG: "downDiagonal",
  /** Upward diagonal lines (/) */
  UP_DIAG: "upDiagonal",
  /** Light downward diagonal lines */
  LT_DN_DIAG: "lightDownDiagonal",
  /** Light upward diagonal lines */
  LT_UP_DIAG: "lightUpDiagonal",
  /** Dark downward diagonal lines */
  DK_DN_DIAG: "darkDownDiagonal",
  /** Dark upward diagonal lines */
  DK_UP_DIAG: "darkUpDiagonal",
  /** Wide downward diagonal lines */
  WD_DN_DIAG: "wideDownDiagonal",
  /** Wide upward diagonal lines */
  WD_UP_DIAG: "wideUpDiagonal",
  /** Dashed downward diagonal lines */
  DASH_DN_DIAG: "dashedDownDiagonal",
  /** Dashed upward diagonal lines */
  DASH_UP_DIAG: "dashedUpDiagonal",
  /** Diagonal cross pattern (X) */
  DIAG_CROSS: "diagonalCross",
  /** Small checkerboard */
  SM_CHECK: "smallChecker",
  /** Large checkerboard */
  LG_CHECK: "largeChecker",
  /** Small grid */
  SM_GRID: "smallGrid",
  /** Large grid */
  LG_GRID: "largeGrid",
  /** Dot grid */
  DOT_GRID: "dotGrid",
  /** Small confetti */
  SM_CONFETTI: "smallConfetti",
  /** Large confetti */
  LG_CONFETTI: "largeConfetti",
  /** Horizontal brick pattern */
  HORZ_BRICK: "horizontalBrick",
  /** Diagonal brick pattern */
  DIAG_BRICK: "diagonalBrick",
  /** Solid diamond */
  SOLID_DMND: "solidDiamond",
  /** Open diamond */
  OPEN_DMND: "openDiamond",
  /** Dotted diamond */
  DOT_DMND: "dottedDiamond",
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
  pattern: (typeof PresetPattern)[keyof typeof PresetPattern];
  /** Foreground color */
  foregroundColor?: SolidFillOptions;
  /** Background color */
  backgroundColor?: SolidFillOptions;
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
      prst: { key: "prst", value: xsdPattern.to(options.pattern) },
    },
    children,
    name: "a:pattFill",
  });
};
