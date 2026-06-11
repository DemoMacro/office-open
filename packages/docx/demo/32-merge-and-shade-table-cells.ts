import { writeFileSync } from "node:fs";
// Example of how you would merge cells together (Rows and Columns) and apply shading
// Also includes an example on how to center tables

import {
  AlignmentType,
  BorderStyle,
  HeadingLevel,
  ShadingType,
  WidthType,
  generateDocument,
} from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          table: {
            rows: [
              {
                cells: [
                  {
                    children: [{ paragraph: "Hello" }],
                    columnSpan: 2,
                  },
                ],
              },
              {
                cells: [{ children: [] }, { children: [] }],
              },
            ],
          },
        },
        {
          paragraph: {
            heading: HeadingLevel.HEADING_2,
            text: "Another table",
          },
        },
        {
          table: {
            alignment: AlignmentType.CENTER,
            columnWidths: ["0.69in", "0.69in", "0.69in"],
            rows: [
              {
                cells: [
                  {
                    children: [{ paragraph: "World" }],
                    columnSpan: 3,
                    margins: {
                      bottom: "0.69in",
                      left: "0.69in",
                      right: "0.69in",
                      top: "0.69in",
                    },
                  },
                ],
              },
              {
                cells: [{ children: [] }, { children: [] }, { children: [] }],
              },
            ],
            width: {
              size: 100,
              type: WidthType.AUTO,
            },
          },
        },
        {
          paragraph: {
            heading: HeadingLevel.HEADING_2,
            text: "Another table",
          },
        },
        {
          table: {
            alignment: AlignmentType.CENTER,
            margins: {
              bottom: "0.27in",
              left: "0.27in",
              right: "0.27in",
              top: "0.27in",
            },
            rows: [
              {
                cells: [
                  { children: [{ paragraph: "Foo" }] },
                  {
                    children: [{ paragraph: "v" }],
                    columnSpan: 3,
                  },
                ],
              },
              {
                cells: [
                  {
                    children: [{ paragraph: "Bar1" }],
                    shading: {
                      color: "auto",
                      fill: "b79c2f",
                      type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                    },
                  },
                  {
                    children: [{ paragraph: "Bar2" }],
                    shading: {
                      color: "auto",
                      fill: "42c5f4",
                      type: ShadingType.PERCENT_95,
                    },
                  },
                  {
                    children: [{ paragraph: "Bar3" }],
                    shading: {
                      color: "e2df0b",
                      fill: "880aa8",
                      type: ShadingType.PERCENT_10,
                    },
                  },
                  {
                    children: [{ paragraph: "Bar4" }],
                    shading: {
                      color: "auto",
                      fill: "FF0000",
                      type: ShadingType.CLEAR,
                    },
                  },
                ],
              },
            ],
            width: {
              size: "4.86in",
              type: WidthType.DXA,
            },
          },
        },
        { paragraph: "Merging columns 1" },
        {
          table: {
            rows: [
              {
                cells: [
                  {
                    children: [{ paragraph: "0,0" }],
                    columnSpan: 2,
                  },
                ],
              },
              {
                cells: [{ children: [{ paragraph: "1,0" }] }, { children: [{ paragraph: "1,1" }] }],
              },
              {
                cells: [
                  {
                    children: [{ paragraph: "2,0" }],
                    columnSpan: 2,
                  },
                ],
              },
            ],
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
          },
        },
        { paragraph: "Merging columns 2" },
        {
          table: {
            rows: [
              {
                cells: [
                  { children: [{ paragraph: "0,0" }] },
                  {
                    children: [{ paragraph: "0,1" }],
                    rowSpan: 2,
                  },
                  { children: [{ paragraph: "0,2" }] },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "1,0" }] },
                  {
                    children: [{ paragraph: "1,2" }],
                    rowSpan: 2,
                  },
                ],
              },
              {
                cells: [{ children: [{ paragraph: "2,0" }] }, { children: [{ paragraph: "2,1" }] }],
              },
            ],
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
          },
        },
        { paragraph: "Merging columns 3" },
        {
          table: {
            rows: (() => {
              const borders = {
                bottom: {
                  color: "FF0000",
                  size: 1,
                  style: BorderStyle.DASH_SMALL_GAP,
                },
                left: {
                  color: "FF0000",
                  size: 1,
                  style: BorderStyle.DASH_SMALL_GAP,
                },
                right: {
                  color: "FF0000",
                  size: 1,
                  style: BorderStyle.DASH_SMALL_GAP,
                },
                top: {
                  color: "FF0000",
                  size: 1,
                  style: BorderStyle.DASH_SMALL_GAP,
                },
              };
              return [
                {
                  cells: [
                    {
                      borders,
                      children: [{ paragraph: "0,0" }],
                      rowSpan: 2,
                    },
                    {
                      borders,
                      children: [{ paragraph: "0,1" }],
                    },
                  ],
                },
                {
                  cells: [
                    {
                      borders,
                      children: [{ paragraph: "1,1" }],
                      rowSpan: 2,
                    },
                  ],
                },
                {
                  cells: [
                    {
                      borders,
                      children: [{ paragraph: "2,0" }],
                    },
                  ],
                },
              ];
            })(),
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
          },
        },
        { paragraph: "Merging columns 4" },
        {
          table: {
            rows: [
              {
                cells: [
                  { children: [{ paragraph: "0,0" }] },
                  { children: [{ paragraph: "0,1" }] },
                  {
                    children: [{ paragraph: "0,2" }],
                    rowSpan: 2,
                  },
                  {
                    children: [{ paragraph: "0,3" }],
                    rowSpan: 3,
                  },
                ],
              },
              {
                cells: [
                  {
                    children: [{ paragraph: "1,0" }],
                    columnSpan: 2,
                  },
                ],
              },
              {
                cells: [
                  {
                    children: [{ paragraph: "2,0" }],
                    columnSpan: 2,
                  },
                  {
                    children: [{ paragraph: "2,2" }],
                    rowSpan: 2,
                  },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "3,0" }] },
                  { children: [{ paragraph: "3,1" }] },
                  { children: [{ paragraph: "3,3" }] },
                ],
              },
            ],
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
          },
        },
        { paragraph: "Merging columns 5" },
        {
          table: {
            rows: (() => {
              const borders = {
                bottom: {
                  color: "FF0000",
                  size: 1,
                  style: BorderStyle.DASH_SMALL_GAP,
                },
                left: {
                  color: "FF0000",
                  size: 1,
                  style: BorderStyle.DASH_SMALL_GAP,
                },
                right: {
                  color: "FF0000",
                  size: 1,
                  style: BorderStyle.DASH_SMALL_GAP,
                },
                top: {
                  color: "FF0000",
                  size: 1,
                  style: BorderStyle.DASH_SMALL_GAP,
                },
              };
              return [
                {
                  cells: [
                    { children: [{ paragraph: "1,1" }] },
                    { children: [{ paragraph: "1,2" }] },
                    { children: [{ paragraph: "1,3" }] },
                    { borders, children: [{ paragraph: "1,4" }], rowSpan: 4 },
                  ],
                },
                {
                  cells: [
                    { children: [{ paragraph: "2,1" }] },
                    { children: [{ paragraph: "2,2" }] },
                    { children: [{ paragraph: "2,3" }], rowSpan: 3 },
                  ],
                },
                {
                  cells: [
                    { children: [{ paragraph: "3,1" }] },
                    { children: [{ paragraph: "3,2" }], rowSpan: 2 },
                  ],
                },
                {
                  cells: [{ children: [{ paragraph: "4,1" }] }],
                },
              ];
            })(),
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
