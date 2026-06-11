// Page numbers - Start from 0 on a new section

import { writeFileSync } from "node:fs";

import {
  AlignmentType,
  PageNumber,
  PageNumberSeparator,
  generateDocument,
} from "@office-open/docx";

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
    },
    {
      children: [
        {
          paragraph: {
            children: ["Third Page", { pageBreak: true }],
          },
        },
        { paragraph: "Fourth Page" },
      ],
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
                "First Page Header of Second section",
                {
                  children: ["Page ", PageNumber.CURRENT],
                },
              ],
            },
          },
        ],
      },
      properties: {
        page: {
          pageNumbers: {
            separator: PageNumberSeparator.EM_DASH,
            start: 1,
          },
        },
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
