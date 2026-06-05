/**
 * Latent styles module for WordprocessingML documents.
 *
 * Latent styles define styles that are not yet used in the document but can be
 * made available when needed. This reduces file size while maintaining style definitions.
 *
 * Reference: http://officeopenxml.com/WPstyles.php
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import type { LatentStyleException } from "./exceptions";

/**
 * Represents latent style definitions in a WordprocessingML document.
 *
 * Latent styles are style definitions that exist in the application's style gallery
 * but are not yet instantiated in the document. This element specifies default
 * properties for latent styles and can define exceptions for specific styles.
 *
 * Reference: http://officeopenxml.com/WPstyles.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_LatentStyles">
 *   <xsd:sequence>
 *     <xsd:element name="lsdException" type="CT_LsdException" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="defLockedState" type="s:ST_OnOff"/>
 *   <xsd:attribute name="defUIPriority" type="ST_DecimalNumber"/>
 *   <xsd:attribute name="defSemiHidden" type="s:ST_OnOff"/>
 *   <xsd:attribute name="defUnhideWhenUsed" type="s:ST_OnOff"/>
 *   <xsd:attribute name="defQFormat" type="s:ST_OnOff"/>
 *   <xsd:attribute name="count" type="ST_DecimalNumber"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Define latent styles with exceptions
 * new LatentStyles(
 *   new LatentStyleException({ name: "Heading 1", uiPriority: "9", qFormat: "1" })
 * );
 * ```
 */
export interface LatentStylesOptions {
  readonly defaultLockedState?: boolean;
  readonly defaultUIPriority?: number;
  readonly defaultSemiHidden?: boolean;
  readonly defaultUnhideWhenUsed?: boolean;
  readonly defaultQuickFormat?: boolean;
}

export class LatentStyles extends XmlComponent {
  public constructor(latentException?: LatentStyleException, options?: LatentStylesOptions) {
    super("w:latentStyles");

    if (options) {
      const attr: Record<string, string | number> = {};
      if (options.defaultLockedState !== undefined)
        attr["w:defLockedState"] = options.defaultLockedState ? "1" : "0";
      if (options.defaultUIPriority !== undefined)
        attr["w:defUIPriority"] = options.defaultUIPriority;
      if (options.defaultSemiHidden !== undefined)
        attr["w:defSemiHidden"] = options.defaultSemiHidden ? "1" : "0";
      if (options.defaultUnhideWhenUsed !== undefined)
        attr["w:defUnhideWhenUsed"] = options.defaultUnhideWhenUsed ? "1" : "0";
      if (options.defaultQuickFormat !== undefined)
        attr["w:defQFormat"] = options.defaultQuickFormat ? "1" : "0";
      if (Object.keys(attr).length > 0) this.root.push({ _attr: attr });
    }

    if (latentException) {
      this.root.push(latentException);
    }
  }
}
