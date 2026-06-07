import { writeFileSync } from "node:fs";

import { generate, PivotFilterTypeValue } from "@office-open/xlsx";

const funcs = [
  "sum",
  "average",
  "count",
  "countNums",
  "max",
  "min",
  "product",
  "stdDev",
  "stdDevp",
  "var",
  "varp",
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

const buffer = await generate({
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
          pivotConditionalFormats: [
            {
              priority: 1,
              scope: "data",
              type: "all",
              pivotAreas: [
                {
                  field: 2,
                  type: "data",
                  outline: true,
                },
              ],
            },
          ],
        },
      ],
      sheetView: {
        tabSelected: true,
        pivotSelections: [
          {
            pane: "topLeft",
            axis: "axisRow",
            activeRow: 3,
          },
        ],
      },
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
