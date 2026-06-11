// Page numbers

import { writeFileSync } from "node:fs";

import { AlignmentType, PageNumber, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: ["First Page", { pageBreak: true }],
          },
        },
        { paragraph: "Second Page" },
      ],
      footers: {
        default: [
          {
            paragraph: {
              alignment: AlignmentType.RIGHT,
              children: [
                "My Title ",
                {
                  children: ["Footer - Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES],
                },
              ],
            },
          },
        ],
        first: [
          {
            paragraph: {
              alignment: AlignmentType.RIGHT,
              children: [
                "First Page Footer ",
                {
                  children: ["Page ", PageNumber.CURRENT],
                },
              ],
            },
          },
        ],
      },
      headers: {
        default: [
          {
            paragraph: {
              alignment: AlignmentType.RIGHT,
              children: [
                "My Title ",
                {
                  children: ["Page ", PageNumber.CURRENT],
                },
              ],
            },
          },
        ],
        first: [
          {
            paragraph: {
              alignment: AlignmentType.RIGHT,
              children: [
                "First Page Header ",
                {
                  children: ["Page ", PageNumber.CURRENT],
                },
              ],
            },
          },
        ],
      },
      properties: {
        titlePage: true,
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
