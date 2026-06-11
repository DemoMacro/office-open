// Break with clear attribute: controls how text reflows after a line break

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              "Line breaks in JSON API:",
              { text: "First line", break: 1 },
              "Second line (after break)",
              { text: "Third line", break: 2 },
              "Fourth line (after double break)",
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
