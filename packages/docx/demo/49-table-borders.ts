// Add custom borders and no-borders to the table itself

import { writeFileSync } from "node:fs";

import {
  BorderStyle,
  HeadingLevel,
  TABLE_BORDERS_NONE,
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
                    borders: {
                      bottom: {
                        color: "ff0000",
                        size: 1,
                        style: BorderStyle.DASH_SMALL_GAP,
                      },
                      left: {
                        color: "ff0000",
                        size: 1,
                        style: BorderStyle.DASH_SMALL_GAP,
                      },
                      right: {
                        color: "ff0000",
                        size: 1,
                        style: BorderStyle.DASH_SMALL_GAP,
                      },
                      top: {
                        color: "ff0000",
                        size: 1,
                        style: BorderStyle.DASH_SMALL_GAP,
                      },
                    },
                    children: [{ paragraph: "Hello" }],
                  },
                  { children: [] },
                ],
              },
              {
                cells: [{ children: [] }, { children: [{ paragraph: "World" }] }],
              },
            ],
          },
        },
        { paragraph: "Hello" },
        // Using the no-border convenience object. It is the same as writing this manually:
        // Const borders = {
        //     Top: {
        //         Style: BorderStyle.NONE,
        //         Size: 0,
        //         Color: "auto",
        //     },
        //     Bottom: {
        //         Style: BorderStyle.NONE,
        //         Size: 0,
        //         Color: "auto",
        //     },
        //     Left: {
        //         Style: BorderStyle.NONE,
        //         Size: 0,
        //         Color: "auto",
        //     },
        //     Right: {
        //         Style: BorderStyle.NONE,
        //         Size: 0,
        //         Color: "auto",
        //     },
        //     InsideHorizontal: {
        //         Style: BorderStyle.NONE,
        //         Size: 0,
        //         Color: "auto",
        //     },
        //     InsideVertical: {
        //         Style: BorderStyle.NONE,
        //         Size: 0,
        //         Color: "auto",
        //     },
        // };
        {
          table: {
            borders: TABLE_BORDERS_NONE,
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
