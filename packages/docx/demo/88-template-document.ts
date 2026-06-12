// Patch a document with patches

import { readFileSync, writeFileSync } from "node:fs";

import type { Patch } from "@office-open/docx";
import { PatchType, patchDocument } from "@office-open/docx";

export const font = "Trebuchet MS";
export const getPatches = (fields: Record<string, string>) => {
  const patches: Record<string, Patch> = {};

  for (const field in fields) {
    patches[field] = {
      children: [{ font, text: fields[field] }],
      type: PatchType.PARAGRAPH,
    };
  }

  return patches;
};

const patches = getPatches({
  item_1: "Doe",
  name: "Mr",
  paragraph_replace: "Lorem ipsum paragraph",
  table_heading_1: "John",
});

const doc = await patchDocument({
  data: readFileSync("demo/assets/simple-template.docx"),
  outputType: "nodebuffer",
  patches,
});
writeFileSync("My Document.docx", doc);
