// Example of how you would create a table with float positions

import { writeFileSync } from "node:fs";

import {
  OverlapType,
  RelativeHorizontalPosition,
  RelativeVerticalPosition,
  TableAnchorType,
  TableLayoutType,
  WidthType,
  generateDocument,
} from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          table: {
            float: {
              bottomFromText: 30,
              horizontalAnchor: TableAnchorType.MARGIN,
              leftFromText: 1000,
              overlap: OverlapType.NEVER,
              relativeHorizontalPosition: RelativeHorizontalPosition.RIGHT,
              relativeVerticalPosition: RelativeVerticalPosition.BOTTOM,
              rightFromText: 2000,
              topFromText: 1500,
              verticalAnchor: TableAnchorType.MARGIN,
            },
            layout: TableLayoutType.FIXED,
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
            width: {
              size: 4535,
              type: WidthType.DXA,
            },
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
