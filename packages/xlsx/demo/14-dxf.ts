import { writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const buffer = await generateWorkbook({
  dxfs: [
    // DXF 0: red text for values > 200
    { font: { color: "FF0000", bold: true } },
    // DXF 1: green fill for values > 100
    { fill: { color: "C6EFCE" } },
    // DXF 2: yellow fill for values < 50
    { fill: { color: "FFEB9C" } },
  ],
  worksheets: [
    {
      name: "Conditional DXFs",
      rows: [
        {
          cells: [
            { value: "Product" },
            { value: "Q1" },
            { value: "Q2" },
            { value: "Q3" },
            { value: "Q4" },
          ],
        },
        {
          cells: [
            { value: "Widget A" },
            { value: 50 },
            { value: 150 },
            { value: 250 },
            { value: 30 },
          ],
        },
        {
          cells: [
            { value: "Widget B" },
            { value: 80 },
            { value: 220 },
            { value: 40 },
            { value: 180 },
          ],
        },
        {
          cells: [
            { value: "Widget C" },
            { value: 120 },
            { value: 90 },
            { value: 300 },
            { value: 60 },
          ],
        },
      ],
      conditionalFormats: [
        {
          sqref: "B2:E4",
          rules: [
            { type: "cellIs", operator: "greaterThan", formulas: ["200"], dxfId: 0, priority: 1 },
            { type: "cellIs", operator: "greaterThan", formulas: ["100"], dxfId: 1, priority: 2 },
            { type: "cellIs", operator: "lessThan", formulas: ["50"], dxfId: 2, priority: 3 },
          ],
        },
      ],
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
