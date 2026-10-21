// Patch canonical document text in an existing template.

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";

import { patchDocument } from "@office-open/docx";

const replacements = {
  footer_text: "replaced just as well",
  header_adjective: "Delightful Header",
  image_test: "Replaced image marker",
  item_1: "#657",
  name: "Sir. John Doe (The Conqueror)",
  paragraph_replace: "Lorem ipsum paragraph. Another paragraph.",
  table: "Replaced table",
  table_heading_1: "Heading wow!",
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
writeFileSync(".temp/85-template-document.docx", doc);
