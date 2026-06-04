/**
 * Custom XML range markers for WordprocessingML track changes.
 *
 * These elements mark custom XML insertion, deletion, and move ranges.
 * They parallel the standard track change markers but operate on
 * custom XML markup.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_CustomXmlRun
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

// ── Insert range ──

/**
 * Marks the start of a custom XML insertion range.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, customXmlInsRangeStart (CT_TrackChange)
 */
export class CustomXmlInsRangeStart extends XmlComponent {
  public constructor(id: number, author: string, date?: string) {
    super("w:customXmlInsRangeStart");
    const attrs: Record<string, string | number> = { "w:id": id, "w:author": author };
    if (date !== undefined) attrs["w:date"] = date;
    this.root.push({ _attr: attrs });
  }
}

/**
 * Marks the end of a custom XML insertion range.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, customXmlInsRangeEnd (CT_Markup)
 */
export class CustomXmlInsRangeEnd extends XmlComponent {
  public constructor(id: number) {
    super("w:customXmlInsRangeEnd");
    this.root.push({ _attr: { "w:id": id } });
  }
}

// ── Delete range ──

/**
 * Marks the start of a custom XML deletion range.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, customXmlDelRangeStart (CT_TrackChange)
 */
export class CustomXmlDelRangeStart extends XmlComponent {
  public constructor(id: number, author: string, date?: string) {
    super("w:customXmlDelRangeStart");
    const attrs: Record<string, string | number> = { "w:id": id, "w:author": author };
    if (date !== undefined) attrs["w:date"] = date;
    this.root.push({ _attr: attrs });
  }
}

/**
 * Marks the end of a custom XML deletion range.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, customXmlDelRangeEnd (CT_Markup)
 */
export class CustomXmlDelRangeEnd extends XmlComponent {
  public constructor(id: number) {
    super("w:customXmlDelRangeEnd");
    this.root.push({ _attr: { "w:id": id } });
  }
}

// ── Move from range ──

/**
 * Marks the start of a custom XML move-from range.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, customXmlMoveFromRangeStart (CT_TrackChange)
 */
export class CustomXmlMoveFromRangeStart extends XmlComponent {
  public constructor(id: number, author: string, date?: string) {
    super("w:customXmlMoveFromRangeStart");
    const attrs: Record<string, string | number> = { "w:id": id, "w:author": author };
    if (date !== undefined) attrs["w:date"] = date;
    this.root.push({ _attr: attrs });
  }
}

/**
 * Marks the end of a custom XML move-from range.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, customXmlMoveFromRangeEnd (CT_Markup)
 */
export class CustomXmlMoveFromRangeEnd extends XmlComponent {
  public constructor(id: number) {
    super("w:customXmlMoveFromRangeEnd");
    this.root.push({ _attr: { "w:id": id } });
  }
}

// ── Move to range ──

/**
 * Marks the start of a custom XML move-to range.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, customXmlMoveToRangeStart (CT_TrackChange)
 */
export class CustomXmlMoveToRangeStart extends XmlComponent {
  public constructor(id: number, author: string, date?: string) {
    super("w:customXmlMoveToRangeStart");
    const attrs: Record<string, string | number> = { "w:id": id, "w:author": author };
    if (date !== undefined) attrs["w:date"] = date;
    this.root.push({ _attr: attrs });
  }
}

/**
 * Marks the end of a custom XML move-to range.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, customXmlMoveToRangeEnd (CT_Markup)
 */
export class CustomXmlMoveToRangeEnd extends XmlComponent {
  public constructor(id: number) {
    super("w:customXmlMoveToRangeEnd");
    this.root.push({ _attr: { "w:id": id } });
  }
}
