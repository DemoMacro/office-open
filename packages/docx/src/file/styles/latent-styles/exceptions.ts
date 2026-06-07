import { XmlComponent } from "@file/xml-components";

/**
 * Properties for latent style exception attributes.
 *
 * @property name - The name of the style for this exception
 * @property uiPriority - UI priority for displaying the style
 * @property qFormat - Whether this style should appear in the quick format gallery
 * @property semiHidden - Whether the style is semi-hidden in the UI
 * @property unhideWhenUsed - Whether the style should unhide when used
 */
export interface LatentStyleExceptionAttributesProperties {
  /** The name of the style for this exception */
  name?: string;
  /** UI priority for displaying the style */
  uiPriority?: string;
  /** Whether this style should appear in the quick format gallery */
  qFormat?: string;
  /** Whether the style is semi-hidden in the UI */
  semiHidden?: string;
  /** Whether the style should unhide when used */
  unhideWhenUsed?: string;
}

/**
 * Represents an exception to the default latent style properties.
 *
 * This element allows specific latent styles to have different properties
 * from the defaults specified in the parent LatentStyles element.
 *
 * Reference: http://officeopenxml.com/WPstyles.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_LsdException">
 *   <xsd:attribute name="name" type="s:ST_String" use="required"/>
 *   <xsd:attribute name="locked" type="s:ST_OnOff"/>
 *   <xsd:attribute name="uiPriority" type="ST_DecimalNumber"/>
 *   <xsd:attribute name="semiHidden" type="s:ST_OnOff"/>
 *   <xsd:attribute name="unhideWhenUsed" type="s:ST_OnOff"/>
 *   <xsd:attribute name="qFormat" type="s:ST_OnOff"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Create an exception for Heading 1
 * new LatentStyleException({
 *   name: "Heading 1",
 *   uiPriority: "9",
 *   qFormat: "1"
 * });
 * ```
 */
export class LatentStyleException extends XmlComponent {
  public constructor(attributes: LatentStyleExceptionAttributesProperties) {
    super("w:lsdException");
    const attrs: Record<string, string> = {};
    if (attributes.name !== undefined) {
      attrs["w:name"] = attributes.name;
    }
    if (attributes.qFormat !== undefined) {
      attrs["w:qFormat"] = attributes.qFormat;
    }
    if (attributes.semiHidden !== undefined) {
      attrs["w:semiHidden"] = attributes.semiHidden;
    }
    if (attributes.uiPriority !== undefined) {
      attrs["w:uiPriority"] = attributes.uiPriority;
    }
    if (attributes.unhideWhenUsed !== undefined) {
      attrs["w:unhideWhenUsed"] = attributes.unhideWhenUsed;
    }
    this.root.push({ _attr: attrs });
  }
}
