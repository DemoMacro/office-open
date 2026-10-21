// Patch canonical document text in an existing template.

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";

import { patchDocument } from "@office-open/docx";

const doc = await patchDocument({
  data: readFileSync("demo/assets/simple-template-3.docx"),
  commands: [
    {
      type: "replace",
      at: { field: "sections", text: { placeholder: "first-name" } },
      value: "John",
    },
    {
      type: "replace",
      at: { field: "sections", text: { placeholder: "salutation" } },
      value: "Mr.",
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/89-template-document.docx", doc);
