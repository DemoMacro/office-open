// Section with 2 columns including a column break

import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              "This text will be in the first column.",
              { columnBreak: true },
              "This text will be in the second column.",
            ],
          },
        },
      ],
      properties: {
        columns: {
          count: 2,
          space: "1.2cm",
        },
      },
    },
  ],
});
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/67-column-break.docx", buffer);
