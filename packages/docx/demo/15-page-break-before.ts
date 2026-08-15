// Page break before example

import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        { paragraph: "Hello World" },
        {
          paragraph: {
            pageBreakBefore: true,
            text: "Hello World on another page",
          },
        },
      ],
    },
  ],
});
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/15-page-break-before.docx", buffer);
