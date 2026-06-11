// Patch a document with patches

import { readFileSync, writeFileSync } from "node:fs";

import { PatchType, patchDocument } from "@office-open/docx";

const doc = await patchDocument({
  data: readFileSync("demo/assets/simple-template-2.docx"),
  outputType: "nodebuffer",
  patches: {
    name: {
      children: ["Max"],
      type: PatchType.PARAGRAPH,
    },
  },
});
writeFileSync("My Document.docx", doc);
