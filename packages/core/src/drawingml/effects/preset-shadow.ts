/**
 * Preset shadow effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_PresetShadowEffect
 *
 * @module
 */
import { element } from "@office-open/xml";

import { xsdPresetShadow } from "../../util/mappings";
import { createColorElement } from "../color/solid-fill";
import type { SolidFillOptions } from "../color/solid-fill";

/**
 * Preset shadow types (20 variations).
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, ST_PresetShadowVal
 */
export type PresetShadow =
  | "shadow1"
  | "shadow2"
  | "shadow3"
  | "shadow4"
  | "shadow5"
  | "shadow6"
  | "shadow7"
  | "shadow8"
  | "shadow9"
  | "shadow10"
  | "shadow11"
  | "shadow12"
  | "shadow13"
  | "shadow14"
  | "shadow15"
  | "shadow16"
  | "shadow17"
  | "shadow18"
  | "shadow19"
  | "shadow20";

/**
 * Options for preset shadow effect.
 */
export interface PresetShadowEffectOptions {
  /** Preset shadow type (required) */
  preset: PresetShadow;
  /** Distance from shape in EMUs */
  distance?: number;
  /** Direction angle in degrees (0–360). */
  direction?: number;
  /** Shadow color */
  color: SolidFillOptions;
}

/**
 * Creates a preset shadow effect element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_PresetShadowEffect">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_ColorChoice" minOccurs="1" maxOccurs="1"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="prst" type="ST_PresetShadowVal" use="required"/>
 *   <xsd:attribute name="dist" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="dir" type="ST_PositiveFixedAngle" default="0"/>
 * </xsd:complexType>
 * ```
 */
export const createPresetShadowEffect = (options: PresetShadowEffectOptions): string => {
  const attrs: Record<string, string | number> = {
    prst: xsdPresetShadow.to(options.preset),
  };

  if (options.distance !== undefined) attrs.dist = options.distance;
  if (options.direction !== undefined) attrs.dir = Math.round(options.direction * 60000);

  return element("a:prstShdw", attrs, [createColorElement(options.color)]);
};
