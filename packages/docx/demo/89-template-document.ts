// Patch a document with patches

import { readFileSync, writeFileSync } from "node:fs";

import type { IPatch } from "@office-open/docx";
import { PatchType, patchDocument } from "@office-open/docx";

export const font = "Trebuchet MS";
export const getPatches = (fields: Record<string, string>) => {
  const patches: Record<string, IPatch> = {};

  for (const field in fields) {
    patches[field] = {
      children: [{ font, text: fields[field] }],
      type: PatchType.PARAGRAPH,
    };
  }

  return patches;
};

const patches = getPatches({
  "first-name": "John",
  salutation: "Mr.",
});

const doc = await patchDocument({
  data: readFileSync("demo/assets/simple-template-3.docx"),
  keepOriginalStyles: true,
  outputType: "nodebuffer",
  patches,
});
writeFileSync("My Document.docx", doc);
