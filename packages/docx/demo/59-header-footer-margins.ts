// Move + offset header and footer

import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [{ paragraph: "Hello World" }],
      footers: {
        default: [
          {
            paragraph: {
              indent: {
                left: -400,
              },
              text: "Footer text",
            },
          },
        ],
      },
      headers: {
        default: [
          {
            paragraph: {
              indent: {
                left: -400,
              },
              text: "Header text",
            },
          },
          {
            paragraph: {
              indent: {
                left: -600,
              },
              text: "Some more header text",
            },
          },
        ],
      },
      properties: {
        page: {
          margin: {
            footer: "0.1cm",
            header: "0.2cm",
          },
        },
      },
    },
  ],
});
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/59-header-footer-margins.docx", buffer);
