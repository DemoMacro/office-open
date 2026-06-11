// Usage of different Section Types

import { writeFileSync } from "node:fs";

import { SectionType, generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              "Hello World",
              {
                bold: true,
                text: "Foo Bar",
              },
            ],
          },
        },
      ],
      properties: {},
    },
    {
      children: [
        {
          paragraph: {
            children: [
              "Hello World",
              {
                bold: true,
                text: "Foo Bar",
              },
            ],
          },
        },
      ],
      properties: {
        type: SectionType.CONTINUOUS,
      },
    },
    {
      children: [
        {
          paragraph: {
            children: [
              "Hello World",
              {
                bold: true,
                text: "Foo Bar",
              },
            ],
          },
        },
      ],
      properties: {
        type: SectionType.ODD_PAGE,
      },
    },
    {
      children: [
        {
          paragraph: {
            children: [
              "Hello World",
              {
                bold: true,
                text: "Foo Bar",
              },
            ],
          },
        },
      ],
      properties: {
        type: SectionType.EVEN_PAGE,
      },
    },
    {
      children: [
        {
          paragraph: {
            children: [
              "Hello World",
              {
                bold: true,
                text: "Foo Bar",
              },
            ],
          },
        },
      ],
      properties: {
        type: SectionType.NEXT_PAGE,
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
