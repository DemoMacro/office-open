// Patch canonical document text in an existing template.

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";

import { patchDocument } from "@office-open/docx";

const replacements = {
  item_1: "Doe",
  name: "Mr",
  paragraph_replace: "Lorem ipsum paragraph",
  table_heading_1: "John",
};

const doc = await patchDocument({
  data: readFileSync("demo/assets/simple-template.docx"),
  commands: Object.entries(replacements).map(([placeholder, value]) => ({
    type: "replace",
    at: { field: "sections", text: { placeholder } },
    value,
  })),
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/88-template-document.docx", doc);
