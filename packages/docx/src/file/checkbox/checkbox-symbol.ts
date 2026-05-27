/**
 * CheckBox symbol module for WordprocessingML documents.
 *
 * Provides XML components for checkbox symbol states.
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";
import { shortHexNumber } from "@util/values";

/**
 * Represents a checkbox symbol element (checked or unchecked state).
 *
 * This element defines the appearance of a checkbox in a particular state,
 * specifying the Unicode character and font to render.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_SdtCheckboxSymbol">
 *   <xsd:attribute name="font" type="w:ST_String"/>
 *   <xsd:attribute name="val" type="w:ST_ShortHexNumber"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Checked state symbol
 * new CheckBoxSymbolElement("w14:checkedState", "2612", "MS Gothic");
 *
 * // Unchecked state symbol
 * new CheckBoxSymbolElement("w14:uncheckedState", "2610", "MS Gothic");
 *
 * // Symbol without explicit font
 * new CheckBoxSymbolElement("w14:checked", "1");
 * ```
 */
export class CheckBoxSymbolElement extends XmlComponent {
  public constructor(name: string, val: string, font?: string) {
    super(name);
    if (font) {
      this.root.push({ _attr: { "w14:font": font, "w14:val": shortHexNumber(val) } });
    } else {
      this.root.push({ _attr: { "w14:val": val } });
    }
  }
}
