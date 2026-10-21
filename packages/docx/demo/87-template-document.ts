// Patch canonical document text in an existing template.

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";

import { patchDocument } from "@office-open/docx";

const doc = await patchDocument({
  data: readFileSync("demo/assets/simple-template-2.docx"),
  commands: [
    {
      type: "replace",
      at: { field: "sections", text: { placeholder: "name" } },
      value: "Max",
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/87-template-document.docx", doc);
