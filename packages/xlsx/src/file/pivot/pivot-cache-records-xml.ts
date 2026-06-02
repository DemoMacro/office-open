/**
 * PivotCacheRecords XML generator.
 *
 * Generates xl/pivotCache/pivotCacheRecords{N}.xml.
 * Follows CT_PivotCacheRecords from sml.xsd.
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";

import type { PivotSourceData } from "./pivot-utils";
import { collectUniqueValues, isNumericField } from "./pivot-utils";

export class PivotCacheRecordsXml extends BaseXmlComponent {
  private readonly sourceData: PivotSourceData;
  private readonly numericFields: readonly boolean[];
  /** For string fields: value → sharedItems index */
  private readonly fieldIndexMaps: readonly Map<string, number>[];

  public constructor(sourceData: PivotSourceData) {
    super("pivotCacheRecords");
    this.sourceData = sourceData;
    this.numericFields = sourceData.fieldNames.map((_, i) => isNumericField(sourceData.records, i));
    this.fieldIndexMaps = sourceData.fieldNames.map((_, i) => {
      if (this.numericFields[i]) return new Map<string, number>();
      const unique = collectUniqueValues(sourceData.records, i);
      const map = new Map<string, number>();
      for (let j = 0; j < unique.length; j++) {
        map.set(String(unique[j]), j);
      }
      return map;
    });
  }

  public override toXml(_context: Context): string {
    const p: string[] = [];

    p.push(
      '<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ` count="${this.sourceData.records.length}">`,
    );

    for (const row of this.sourceData.records) {
      p.push("<r>");
      for (let i = 0; i < row.length; i++) {
        const val = row[i];
        if (this.numericFields[i]) {
          p.push(`<n v="${val}"/>`);
        } else {
          const idx = this.fieldIndexMaps[i].get(String(val)) ?? 0;
          p.push(`<x v="${idx}"/>`);
        }
      }
      p.push("</r>");
    }

    p.push("</pivotCacheRecords>");
    return p.join("");
  }
}
