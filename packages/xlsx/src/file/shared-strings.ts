/**
 * Shared Strings Table — generates xl/sharedStrings.xml.
 *
 * XLSX stores repeated string values in a central table to reduce file size.
 * Cells reference strings by index into this table.
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";

export class SharedStrings extends BaseXmlComponent {
  private readonly strings: string[] = [];
  private readonly indexMap = new Map<string, number>();

  public constructor() {
    super("sst");
  }

  /**
   * Register a string and return its index.
   * Returns existing index if the string is already registered.
   */
  public register(s: string): number {
    const existing = this.indexMap.get(s);
    if (existing !== undefined) return existing;

    const idx = this.strings.length;
    this.strings.push(s);
    this.indexMap.set(s, idx);
    return idx;
  }

  public get count(): number {
    return this.strings.length;
  }

  public override prepForXml(_context: Context): IXmlableObject {
    const children: IXmlableObject[] = [
      {
        _attr: {
          xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          count: this.strings.length,
          uniqueCount: this.indexMap.size,
        },
      },
    ];

    for (const s of this.strings) {
      children.push({ si: [{ t: [s] }] });
    }

    return { sst: children };
  }
}
