// Multiple sections with total number of pages in each section

import { writeFileSync } from "node:fs";

import { AlignmentType, NumberFormat, PageNumber, generateDocument } from "@office-open/docx";

const header = [
  {
    paragraph: {
      alignment: AlignmentType.CENTER,
      children: [
        "Header on another page",
        {
          children: ["Page number: ", PageNumber.CURRENT],
        },
        {
          children: [" to ", PageNumber.TOTAL_PAGES_IN_SECTION],
        },
      ],
    },
  },
];

const footer = [{ paragraph: "Foo Bar corp. " }];

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: ["Section 1", { pageBreak: true }, "Section 1", { pageBreak: true }],
          },
        },
      ],
      footers: {
        default: footer,
      },
      headers: {
        default: header,
      },
      properties: {
        page: {
          pageNumbers: {
            formatType: NumberFormat.DECIMAL,
            start: 1,
          },
        },
      },
    },
    {
      children: [
        {
          paragraph: {
            children: ["Section 2", { pageBreak: true }, "Section 2", { pageBreak: true }],
          },
        },
      ],
      footers: {
        default: footer,
      },
      headers: {
        default: header,
      },
      properties: {
        page: {
          pageNumbers: {
            formatType: NumberFormat.DECIMAL,
            start: 1,
          },
        },
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
