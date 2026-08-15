// Example of how to set the document to landscape

import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument, PageOrientation } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [{ paragraph: "Hello World" }],
      properties: {
        page: {
          size: {
            orientation: PageOrientation.LANDSCAPE,
          },
        },
      },
    },
  ],
});
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/7-landscape.docx", buffer);
