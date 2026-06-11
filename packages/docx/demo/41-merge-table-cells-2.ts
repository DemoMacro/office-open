// Multiple cells merging in the same table - Rows and Columns

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          table: {
            rows: [
              {
                cells: [
                  { children: [{ paragraph: "0,0" }] },
                  {
                    children: [{ paragraph: "0,1" }],
                    columnSpan: 2,
                  },
                  { children: [{ paragraph: "0,3" }] },
                  {
                    children: [{ paragraph: "0,4" }],
                    columnSpan: 2,
                  },
                ],
              },
              {
                cells: [
                  {
                    children: [{ paragraph: "1,0" }],
                    columnSpan: 2,
                  },
                  {
                    children: [{ paragraph: "1,2" }],
                    columnSpan: 2,
                  },
                  {
                    children: [{ paragraph: "1,4" }],
                    columnSpan: 2,
                  },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "2,0" }] },
                  {
                    children: [{ paragraph: "2,1" }],
                    columnSpan: 2,
                  },
                  { children: [{ paragraph: "2,3" }] },
                  {
                    children: [{ paragraph: "2,4" }],
                    columnSpan: 2,
                  },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "3,0" }] },
                  { children: [{ paragraph: "3,1" }] },
                  { children: [{ paragraph: "3,2" }] },
                  { children: [{ paragraph: "3,3" }] },
                  { children: [{ paragraph: "3,4" }] },
                  { children: [{ paragraph: "3,5" }] },
                ],
              },
              {
                cells: [
                  {
                    children: [{ paragraph: "4,0" }],
                    columnSpan: 5,
                  },
                  { children: [{ paragraph: "4,5" }] },
                ],
              },
              {
                cells: [
                  { children: [] },
                  { children: [] },
                  { children: [] },
                  { children: [] },
                  { children: [] },
                  { children: [] },
                ],
              },
            ],
          },
        },
        { paragraph: "" },
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
                  { children: [{ paragraph: "0,3" }] },
                  { children: [{ paragraph: "0,4" }] },
                  { children: [{ paragraph: "0,5" }] },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "1,0" }] },
                  { children: [{ paragraph: "1,2" }] },
                  { children: [{ paragraph: "1,3" }] },
                  { children: [{ paragraph: "1,4" }] },
                  { children: [{ paragraph: "1,5" }] },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "2,0" }] },
                  { children: [{ paragraph: "2,1" }] },
                  { children: [{ paragraph: "2,2" }] },
                  { children: [{ paragraph: "2,3" }] },
                  { children: [{ paragraph: "2,4" }] },
                  { children: [{ paragraph: "2,5" }] },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "3,0" }] },
                  { children: [{ paragraph: "3,1" }] },
                  { children: [{ paragraph: "3,2" }] },
                  { children: [{ paragraph: "3,3" }] },
                  { children: [{ paragraph: "3,4" }] },
                  { children: [{ paragraph: "3,5" }] },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "4,0" }] },
                  { children: [{ paragraph: "4,1" }] },
                  { children: [{ paragraph: "4,2" }] },
                  { children: [{ paragraph: "4,3" }] },
                  { children: [{ paragraph: "4,4" }] },
                  { children: [{ paragraph: "4,5" }] },
                ],
              },
              {
                cells: [
                  { children: [] },
                  { children: [] },
                  { children: [] },
                  { children: [] },
                  { children: [] },
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
