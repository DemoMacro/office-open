// Example how to display page numbers

import { mkdirSync, writeFileSync } from "node:fs";

import { AlignmentType, NumberFormat, PageNumber, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: ["Hello World 1", { pageBreak: true }],
          },
        },
        {
          paragraph: {
            children: ["Hello World 2", { pageBreak: true }],
          },
        },
        {
          paragraph: {
            children: ["Hello World 3", { pageBreak: true }],
          },
        },
        {
          paragraph: {
            children: ["Hello World 4", { pageBreak: true }],
          },
        },
        {
          paragraph: {
            children: ["Hello World 5", { pageBreak: true }],
          },
        },
      ],
      footers: {
        default: [
          {
            paragraph: {
              alignment: AlignmentType.CENTER,
              children: [
                "Foo Bar corp. ",
                {
                  children: ["Page Number: ", PageNumber.CURRENT],
                },
                {
                  children: [" to ", PageNumber.TOTAL_PAGES],
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
              children: [
                "Foo Bar corp. ",
                {
                  children: ["Page Number ", PageNumber.CURRENT],
                },
                {
                  children: [" to ", PageNumber.TOTAL_PAGES],
                },
              ],
            },
          },
        ],
      },
      properties: {
        pageNumberType: {
          format: NumberFormat.DECIMAL,
          start: 1,
        },
      },
    },
  ],
});
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/39-page-numbers.docx", buffer);
