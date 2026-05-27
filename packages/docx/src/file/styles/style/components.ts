/**
 * Style components module for WordprocessingML documents.
 *
 * Provides common elements used in style definitions.
 *
 * Reference: http://officeopenxml.com/WPstyleGenProps.php
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";
import { decimalNumber } from "@util/values";

/**
 * Represents the name element of a style.
 *
 * This element specifies the display name of a style as shown in the user interface.
 *
 * Reference: http://officeopenxml.com/WPstyleGenProps.php
 *
 * @example
 * ```typescript
 * // Typically used internally by Style class
 * new Name("Heading 1");
 * ```
 */
export class Name extends XmlComponent {
  public constructor(value: string) {
    super("w:name");
    this.root.push({ _attr: { "w:val": value } });
  }
}

/**
 * Represents the UI priority of a style.
 *
 * This element specifies the sort order priority for displaying the style in the user interface.
 * Lower numbers appear first in style lists.
 *
 * Reference: http://officeopenxml.com/WPstyleGenProps.php
 *
 * @example
 * ```typescript
 * // Typically used internally by Style class
 * new UiPriority(99);
 * ```
 */
export class UiPriority extends XmlComponent {
  public constructor(value: number) {
    super("w:uiPriority");
    this.root.push({ _attr: { "w:val": decimalNumber(value) } });
  }
}

/**
 * Represents table properties for table styles.
 *
 * @internal
 */
export class TableProperties extends XmlComponent {}

/**
 * Represents revision save ID for tracking document changes.
 *
 * @internal
 */
export class RsId extends XmlComponent {}
