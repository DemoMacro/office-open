// Section with 2 columns including a column break

import { writeFileSync } from "node:fs";

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
        column: {
          count: 2,
          space: 708,
        },
      },
    },
  ],
});
writeFileSync("My Document.docx", buffer);
