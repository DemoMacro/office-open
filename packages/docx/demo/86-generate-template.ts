// Generate a template document

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: ["{{template}}"],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
