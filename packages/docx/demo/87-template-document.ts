// Patch a document with patches

import { readFileSync, writeFileSync } from "node:fs";

import { patchDocument } from "@office-open/docx";

const doc = await patchDocument({
  data: readFileSync("demo/assets/simple-template-2.docx"),
  outputType: "nodebuffer",
  placeholders: {
    name: {
      children: ["Max"],
      type: "paragraph",
    },
  },
});
writeFileSync("My Document.docx", doc);
