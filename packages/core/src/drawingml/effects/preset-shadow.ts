/**
 * Preset shadow effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_PresetShadowEffect
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";
import type { XmlComponent } from "../../xml-components";
import { xsdPresetShadow } from "../../xsd-mappings";
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
  readonly preset: (typeof PresetShadowVal)[keyof typeof PresetShadowVal];
  /** Distance from shape in EMUs */
  readonly distance?: number;
  /** Direction angle in 60,000ths of a degree */
  readonly direction?: number;
  /** Shadow color */
  readonly color: SolidFillOptions;
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
export const createPresetShadowEffect = (options: PresetShadowEffectOptions): XmlComponent => {
  const attributes: Record<string, { readonly key: string; readonly value: string | number }> = {
    prst: { key: "prst", value: xsdPresetShadow.to(options.preset) },
  };

  if (options.distance !== undefined) {
    attributes.dist = { key: "dist", value: options.distance };
  }
  if (options.direction !== undefined) {
    attributes.dir = { key: "dir", value: options.direction };
  }

  return new BuilderElement({
    attributes: attributes as never,
    children: [createColorElement(options.color)],
    name: "a:prstShdw",
  });
};
