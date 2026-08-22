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
import { element } from "@office-open/xml";

import { xsdPattern } from "../../util/mappings";
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
/**
 * Preset pattern values (a:pattFill `@prst`, ST_PresetPatternVal) — full
 * words; xsdPattern maps them to the abbreviated XSD tokens.
 */
export type PresetPattern =
  | "percent5"
  | "percent10"
  | "percent20"
  | "percent25"
  | "percent30"
  | "percent40"
  | "percent50"
  | "percent60"
  | "percent70"
  | "percent75"
  | "percent80"
  | "percent90"
  | "horizontal"
  | "vertical"
  | "lightHorizontal"
  | "lightVertical"
  | "darkHorizontal"
  | "darkVertical"
  | "narrowHorizontal"
  | "narrowVertical"
  | "dashedHorizontal"
  | "dashedVertical"
  | "cross"
  | "downDiagonal"
  | "upDiagonal"
  | "lightDownDiagonal"
  | "lightUpDiagonal"
  | "darkDownDiagonal"
  | "darkUpDiagonal"
  | "wideDownDiagonal"
  | "wideUpDiagonal"
  | "dashedDownDiagonal"
  | "dashedUpDiagonal"
  | "diagonalCross"
  | "smallChecker"
  | "largeChecker"
  | "smallGrid"
  | "largeGrid"
  | "dotGrid"
  | "smallConfetti"
  | "largeConfetti"
  | "horizontalBrick"
  | "diagonalBrick"
  | "solidDiamond"
  | "openDiamond"
  | "dottedDiamond"
  | "plaid"
  | "sphere"
  | "weave"
  | "divot"
  | "shingle"
  | "wave"
  | "trellis"
  | "zigZag";

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
  pattern: PresetPattern;
  foregroundColor?: SolidFillOptions;
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
export const createPatternFill = (options: PatternFillOptions): string => {
  const children: string[] = [];

  if (options.foregroundColor) {
    children.push(element("a:fgClr", undefined, [createColorElement(options.foregroundColor)]));
  }

  if (options.backgroundColor) {
    children.push(element("a:bgClr", undefined, [createColorElement(options.backgroundColor)]));
  }

  return element("a:pattFill", { prst: xsdPattern.to(options.pattern) }, children);
};
