// Add custom borders to table cell

import { writeFileSync } from "node:fs";

import { BorderStyle, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          table: {
            rows: [
              {
                cells: [{ children: [] }, { children: [] }, { children: [] }, { children: [] }],
              },
              {
                cells: [
                  { children: [] },
                  {
                    borders: {
                      bottom: {
                        color: "0000FF",
                        size: 3,
                        style: BorderStyle.DOUBLE,
                      },
                      left: {
                        color: "00FF00",
                        size: 3,
                        style: BorderStyle.DASH_DOT_STROKED,
                      },
                      right: {
                        color: "#ff8000",
                        size: 3,
                        style: BorderStyle.DASH_DOT_STROKED,
                      },
                      top: {
                        color: "FF0000",
                        size: 3,
                        style: BorderStyle.DASH_DOT_STROKED,
                      },
                    },
                    children: [{ paragraph: "Hello" }],
                  },
                  { children: [] },
                  { children: [] },
                ],
              },
              {
                cells: [{ children: [] }, { children: [] }, { children: [] }, { children: [] }],
              },
              {
                cells: [{ children: [] }, { children: [] }, { children: [] }, { children: [] }],
              },
              // Row with diagonal cell borders
              {
                cells: [
                  {
                    borders: {
                      topLeftToBottomRight: {
                        color: "000000",
                        size: 4,
                        style: BorderStyle.SINGLE,
                      },
                    },
                    children: [{ paragraph: "tl2br" }],
                  },
                  {
                    borders: {
                      topRightToBottomLeft: {
                        color: "0000FF",
                        size: 4,
                        style: BorderStyle.SINGLE,
                      },
                    },
                    children: [{ paragraph: "tr2bl" }],
                  },
                  {
                    borders: {
                      topLeftToBottomRight: {
                        color: "FF0000",
                        size: 4,
                        style: BorderStyle.SINGLE,
                      },
                      topRightToBottomLeft: {
                        color: "0000FF",
                        size: 4,
                        style: BorderStyle.SINGLE,
                      },
                    },
                    children: [{ paragraph: "X" }],
                  },
                  { children: [] },
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
