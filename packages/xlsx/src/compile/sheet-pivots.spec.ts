import { describe, expect, it } from "vite-plus/test";

import type { SourcePivotTableOptions } from "../parts/pivot/pivot-utils";
import { SharedStrings } from "../parts/shared-strings";
import { renderPivotSheetData } from "./sheet-pivots";

describe("renderPivotSheetData", () => {
  it("reads sparse A1-addressed records and renders every column/data field", () => {
    const source = {
      name: "Data",
      rows: [
        {
          rowNumber: 1,
          cells: [
            { reference: "A1", value: "Region" },
            { reference: "B1", value: "Product" },
            { reference: "C1", value: "Revenue" },
            { reference: "D1", value: "Cost" },
          ],
        },
        {
          rowNumber: 2,
          cells: [
            { reference: "A2", value: "East" },
            { reference: "B2", value: "A" },
            { reference: "C2", value: 10 },
            { reference: "D2", value: 2 },
          ],
        },
        {
          rowNumber: 4,
          cells: [
            { reference: "A4", value: "West" },
            { reference: "B4", value: "B" },
            { reference: "C4", value: 30 },
            { reference: "D4", value: 5 },
          ],
        },
      ],
    };
    const pivot: SourcePivotTableOptions = {
      mode: "source",
      source: "A1:D4",
      sourceSheet: "Data",
      rows: ["Region"],
      columns: ["Product"],
      data: [
        { field: "Revenue", summarize: "sum" },
        { field: "Cost", summarize: "sum" },
      ],
    };
    const output = renderPivotSheetData(
      [pivot],
      [source, { name: "Pivot" }],
      new SharedStrings(),
      "Pivot",
    );

    expect(output.sheetData).toContain('<c r="B4"><v>10</v></c>');
    expect(output.sheetData).toContain('<c r="C4"><v>2</v></c>');
    expect(output.sheetData).toContain('<c r="D5"><v>30</v></c>');
    expect(output.sheetData).toContain('<c r="E5"><v>5</v></c>');
    expect(output.sheetData).toContain('<c r="F4"><v>10</v></c>');
    expect(output.sheetData).toContain('<c r="G4"><v>2</v></c>');
    expect(output.sheetData).toContain('<c r="F6"><v>40</v></c>');
    expect(output.sheetData).toContain('<c r="G6"><v>7</v></c>');
  });
});
