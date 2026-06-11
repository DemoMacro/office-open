// Page break before example

import { writeFileSync } from "node:fs";

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
writeFileSync("My Document.docx", buffer);
