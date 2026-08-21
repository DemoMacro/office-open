import { mkdirSync, writeFileSync } from "node:fs";

import { generateWorkbook, PivotFilterTypeValue } from "@office-open/xlsx";

const funcs = [
  "sum",
  "average",
  "count",
  "countNumbers",
  "max",
  "min",
  "product",
  "standardDeviation",
  "standardDeviationPopulation",
  "variance",
  "variancePopulation",
] as const;

const pivotTables = funcs.map((f, i) => {
  const col = (i % 3) * 3;
  const row = Math.floor(i / 3) * 8 + 3;
  const colLetter = String.fromCharCode(65 + col);
  const label = f.charAt(0).toUpperCase() + f.slice(1);
  return {
    name: `PivotTable_${label}`,
    source: "A1:C9",
    sourceSheet: "Data",
    location: `${colLetter}${row}`,
    rows: ["City"],
    data: [{ field: "Revenue", summarize: f, name: `${label} of Revenue` }],
  };
});

const buffer = await generateWorkbook({
  worksheets: [
    {
      name: "Data",
      rows: [
        { cells: [{ value: "City" }, { value: "Category" }, { value: "Revenue" }] },
        { cells: [{ value: "Beijing" }, { value: "Food" }, { value: 320 }] },
        { cells: [{ value: "Beijing" }, { value: "Tech" }, { value: 580 }] },
        { cells: [{ value: "Shanghai" }, { value: "Food" }, { value: 410 }] },
        { cells: [{ value: "Shanghai" }, { value: "Tech" }, { value: 720 }] },
        { cells: [{ value: "Guangzhou" }, { value: "Food" }, { value: 260 }] },
        { cells: [{ value: "Guangzhou" }, { value: "Tech" }, { value: 390 }] },
        { cells: [{ value: "Shenzhen" }, { value: "Food" }, { value: 195 }] },
        { cells: [{ value: "Shenzhen" }, { value: "Tech" }, { value: 850 }] },
      ],
    },
    {
      name: "Pivot",
      rows: [],
      pivotTables,
    },
    {
      name: "Filtered",
      rows: [],
      pivotTables: [
        {
          name: "FilteredPivot",
          source: "A1:C9",
          sourceSheet: "Data",
          location: "A3",
          rows: ["City"],
          data: [{ field: "Revenue", summarize: "sum" }],
          filters: [
            {
              fld: 0,
              type: PivotFilterTypeValue.CAPTION_NOT_EQUAL,
              id: 1,
              stringValue1: "Guangzhou",
              name: "ExcludeGuangzhou",
            },
          ],
        },
      ],
    },
    {
      name: "AdvancedPivot",
      rows: [],
      pivotTables: [
        {
          name: "Pivot_Advanced",
          source: "A1:C9",
          sourceSheet: "Data",
          location: "A3",
          rows: ["City"],
          columns: ["Category"],
          data: [
            {
              field: "Revenue",
              summarize: "sum",
              name: "Total Revenue",
              sortByTupleItems: [0],
            },
          ],
          // Per-field overrides: hide detail for City field
          fieldOverrides: [{ field: "City", defaultItemSd: false }],
          // Additional definition attributes
          asteriskTotals: true,
          immersive: true,
          // pivotConditionalFormats is intentionally not exercised here: Excel
          // rejects the pivot-view conditional format in this multi-pivot layout
          // (it repairs the record away on open) even with byte-identical
          // copies of Excel's own output, so the demo leaves it out until the
          // trigger is understood.
        },
      ],
      sheetView: {
        tabSelected: true,
      },
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/17-pivot-table.xlsx", buffer);
