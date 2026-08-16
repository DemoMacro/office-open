// Multiple sections and headers

import { mkdirSync, writeFileSync } from "node:fs";

import { NumberFormat, PageNumber, PageOrientation, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [{ paragraph: "Hello World" }],
    },
    {
      children: [{ paragraph: "hello" }],
      footers: {
        default: [{ paragraph: "Footer on another page" }],
      },
      headers: {
        default: [{ paragraph: "First Default Header on another page" }],
      },

      properties: {
        pageNumberType: {
          formatType: NumberFormat.DECIMAL,
          start: 1,
        },
      },
    },
    {
      children: [{ paragraph: "hello in landscape" }],
      footers: {
        default: [{ paragraph: "Footer on another page" }],
      },
      headers: {
        default: [{ paragraph: "Second Default Header on another page" }],
      },
      properties: {
        pageNumberType: {
          formatType: NumberFormat.DECIMAL,
          start: 1,
        },
        pageSize: {
          orientation: PageOrientation.LANDSCAPE,
        },
      },
    },
    {
      children: [
        {
          paragraph:
            "Page number in the header must be 2, because it continues from the previous section.",
        },
      ],
      headers: {
        default: [
          {
            paragraph: {
              children: [
                {
                  children: ["Page number: ", PageNumber.CURRENT],
                },
              ],
            },
          },
        ],
      },

      properties: {
        pageSize: {
          orientation: PageOrientation.PORTRAIT,
        },
      },
    },
    {
      children: [
        {
          paragraph:
            "Page number in the header must be III, because it continues from the previous section, but is defined as upper roman.",
        },
      ],
      headers: {
        default: [
          {
            paragraph: {
              children: [
                {
                  children: ["Page number: ", PageNumber.CURRENT],
                },
              ],
            },
          },
        ],
      },
      properties: {
        pageNumberType: {
          formatType: NumberFormat.UPPER_ROMAN,
        },
        pageSize: {
          orientation: PageOrientation.PORTRAIT,
        },
      },
    },
    {
      children: [
        {
          paragraph:
            "Page number in the header must be 25, because it is defined to start at 25 and to be decimal in this section.",
        },
      ],
      headers: {
        default: [
          {
            paragraph: {
              children: [
                {
                  children: ["Page number: ", PageNumber.CURRENT],
                },
              ],
            },
          },
        ],
      },
      properties: {
        pageNumberType: {
          formatType: NumberFormat.DECIMAL,
          start: 25,
        },
        pageSize: {
          orientation: PageOrientation.PORTRAIT,
        },
      },
    },
  ],
});
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/16-multiple-sections.docx", buffer);
