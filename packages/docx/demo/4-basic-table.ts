// Example of how you would create a table and add data to it

import { writeFileSync } from "node:fs";

import { generateDocument, WidthType } from "@office-open/docx";

const table = {
  table: {
    columnWidths: [3505, 5505],
    rows: [
      {
        cells: [
          {
            children: [{ paragraph: "Hello" }],
            width: {
              size: 3505,
              type: WidthType.DXA,
            },
          },
          {
            children: [],
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
          },
        ],
      },
      {
        cells: [
          {
            children: [],
            width: {
              size: 3505,
              type: WidthType.DXA,
            },
          },
          {
            children: [{ paragraph: "World" }],
            width: {
              size: 5505,
              type: WidthType.DXA,
            },
          },
        ],
      },
    ],
  },
};

const table2 = {
  table: {
    columnWidths: [4505, 4505],
    rows: [
      {
        cells: [
          {
            children: [{ paragraph: "Hello" }],
            width: {
              size: 4505,
              type: WidthType.DXA,
            },
          },
          {
            children: [],
            width: {
              size: 4505,
              type: WidthType.DXA,
            },
          },
        ],
      },
      {
        cells: [
          {
            children: [],
            width: {
              size: 4505,
              type: WidthType.DXA,
            },
          },
          {
            children: [{ paragraph: "World" }],
            width: {
              size: 4505,
              type: WidthType.DXA,
            },
          },
        ],
      },
    ],
  },
};

const table3 = {
  table: {
    rows: [
      {
        cells: [
          {
            children: [{ paragraph: "Hello" }],
          },
          {
            children: [],
          },
        ],
      },
      {
        cells: [
          {
            children: [],
          },
          {
            children: [{ paragraph: "World" }],
          },
        ],
      },
    ],
  },
};

// Table with new properties: row cnfStyle, gridBefore/gridAfter, rowAlignment,
// cell noWrap, fitText, horizontalMerge, headers
const table4 = {
  table: {
    styleRowBandSize: 3,
    styleColBandSize: 2,
    caption: "Sales Report",
    description: "Quarterly sales data",
    rows: [
      {
        cnfStyle: { val: "000000000010" },
        gridBefore: 1,
        gridAfter: 0,
        rowAlignment: "center" as const,
        cells: [
          {
            children: [{ paragraph: "Header A" }],
            noWrap: true,
            fitText: true,
          },
          {
            children: [{ paragraph: "Header B" }],
            noWrap: true,
            fitText: true,
          },
        ],
      },
      {
        widthBefore: { size: 100, type: WidthType.DXA },
        widthAfter: { size: 100, type: WidthType.DXA },
        cells: [
          {
            children: [{ paragraph: "Cell with long text that should not wrap" }],
            noWrap: true,
          },
          {
            children: [{ paragraph: "Normal cell" }],
            horizontalMerge: "continue",
          },
        ],
      },
      {
        cells: [
          {
            children: [{ paragraph: "Restart merge" }],
            horizontalMerge: "restart",
          },
          {
            children: [{ paragraph: "Continued" }],
            horizontalMerge: "continue",
          },
        ],
      },
    ],
  },
};

const buffer = await generateDocument({
  sections: [
    {
      children: [
        { paragraph: { text: "Table with skewed widths" } },
        table,
        { paragraph: { text: "Table with equal widths" } },
        table2,
        { paragraph: { text: "Table without setting widths" } },
        table3,
        {
          paragraph: {
            text: "Table with new properties (cnfStyle, gridBefore/After, noWrap, fitText, horizontalMerge, caption, description)",
          },
        },
        table4,
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
