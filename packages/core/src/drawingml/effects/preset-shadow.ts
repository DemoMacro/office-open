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
export const PresetShadowVal = {
  SHDW1: "shadow1",
  SHDW2: "shadow2",
  SHDW3: "shadow3",
  SHDW4: "shadow4",
  SHDW5: "shadow5",
  SHDW6: "shadow6",
  SHDW7: "shadow7",
  SHDW8: "shadow8",
  SHDW9: "shadow9",
  SHDW10: "shadow10",
  SHDW11: "shadow11",
  SHDW12: "shadow12",
  SHDW13: "shadow13",
  SHDW14: "shadow14",
  SHDW15: "shadow15",
  SHDW16: "shadow16",
  SHDW17: "shadow17",
  SHDW18: "shadow18",
  SHDW19: "shadow19",
  SHDW20: "shadow20",
} as const;

/**
 * Options for preset shadow effect.
 */
export interface PresetShadowEffectOptions {
  /** Preset shadow type (required) */
  preset: (typeof PresetShadowVal)[keyof typeof PresetShadowVal];
  /** Distance from shape in EMUs */
  distance?: number;
  /** Direction angle in 60,000ths of a degree */
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
  if (options.direction !== undefined) attrs.dir = options.direction;

  return element("a:prstShdw", attrs, [createColorElement(options.color)]);
};
