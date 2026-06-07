/**
 * Font Wrapper module for WordprocessingML documents.
 *
 * Manages font table and relationships for embedded fonts.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-w_fonts.html
 *
 * @module
 */
import type { ViewWrapper } from "@file/document-wrapper";
import { Relationships } from "@file/relationships";
import type { XmlComponent } from "@file/xml-components";
import { uniqueUuid } from "@util/convenience-functions";

import { createFontTable } from "./font-table";
import type { EmbeddedFontOptions } from "./font-table";

/**
 * Font options extended with a unique font key.
 */
export type EmbeddedFontOptionsWithKey = EmbeddedFontOptions & { fontKey: string };

/**
 * Wrapper class for managing the font table and its relationships.
 *
 * Creates a font table with embedded font files and manages the relationships
 * required for font embedding. Each font is assigned a unique key for obfuscation.
 *
 * @example
 * ```typescript
 * const fontWrapper = new FontWrapper([
 *   { name: "CustomFont", data: fontBuffer }
 * ]);
 * ```
 */
export class FontWrapper implements ViewWrapper {
  private fontTable: XmlComponent;
  public relationships: Relationships;
  public fontOptionsWithKey: EmbeddedFontOptionsWithKey[] = [];

  public constructor(public options: EmbeddedFontOptions[]) {
    this.fontOptionsWithKey = options.map((o) => ({ ...o, fontKey: uniqueUuid() }));
    this.fontTable = createFontTable(this.fontOptionsWithKey);
    this.relationships = new Relationships();

    for (let i = 0; i < options.length; i++) {
      this.relationships.addRelationship(
        i + 1,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font",
        `fonts/${options[i].name}.odttf`,
      );
    }
  }

  public get view(): XmlComponent {
    return this.fontTable;
  }
}
