// Example of how you would create a table and add data to it

import { writeFileSync } from "node:fs";

import {
  HeadingLevel,
  TextDirection,
  VerticalAlignTable,
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
                    children: [{ paragraph: {} }, { paragraph: {} }],
                    verticalAlign: VerticalAlignTable.CENTER,
                  },
                  {
                    children: [{ paragraph: {} }, { paragraph: {} }],
                    verticalAlign: VerticalAlignTable.CENTER,
                  },
                  {
                    children: [{ paragraph: { text: "bottom to top" } }, { paragraph: {} }],
                    textDirection: TextDirection.BOTTOM_TO_TOP_LEFT_TO_RIGHT,
                  },
                  {
                    children: [{ paragraph: { text: "top to bottom" } }, { paragraph: {} }],
                    textDirection: TextDirection.TOP_TO_BOTTOM_RIGHT_TO_LEFT,
                  },
                ],
              },
              {
                cells: [
                  {
                    children: [
                      {
                        paragraph: {
                          heading: HeadingLevel.HEADING_1,
                          text: "Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah Blah",
                        },
                      },
                    ],
                  },
                  {
                    children: [
                      {
                        paragraph: {
                          text: "This text should be in the middle of the cell",
                        },
                      },
                    ],
                    verticalAlign: VerticalAlignTable.CENTER,
                  },
                  {
                    children: [
                      {
                        paragraph: {
                          text: "Text above should be vertical from bottom to top",
                        },
                      },
                    ],
                    verticalAlign: VerticalAlignTable.CENTER,
                  },
                  {
                    children: [
                      {
                        paragraph: {
                          text: "Text above should be vertical from top to bottom",
                        },
                      },
                    ],
                    verticalAlign: VerticalAlignTable.CENTER,
                  },
                ],
              },
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
