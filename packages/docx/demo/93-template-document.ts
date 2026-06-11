// Patch a document with patches

import { readFileSync, writeFileSync } from "node:fs";

import { PatchType, patchDocument } from "@office-open/docx";

const doc = await patchDocument({
  data: readFileSync("demo/assets/field-trip.docx"),
  outputType: "nodebuffer",
  patches: {
    address: {
      children: [{ text: "blah blah" }],
      type: PatchType.PARAGRAPH,
    },

    city: {
      children: [{ text: "test" }],
      type: PatchType.PARAGRAPH,
    },

    email_address: {
      children: [{ text: "test" }],
      type: PatchType.PARAGRAPH,
    },

    first_name: {
      children: [{ text: "test" }],
      type: PatchType.PARAGRAPH,
    },

    ft_dates: {
      children: [{ text: "test" }],
      type: PatchType.PARAGRAPH,
    },

    grade: {
      children: [{ text: "test" }],
      type: PatchType.PARAGRAPH,
    },

    last_name: {
      children: [{ text: "test" }],
      type: PatchType.PARAGRAPH,
    },

    phone: {
      children: [{ text: "test" }],
      type: PatchType.PARAGRAPH,
    },

    school_name: {
      children: [{ text: "test" }],
      type: PatchType.PARAGRAPH,
    },

    state: {
      children: [{ text: "test" }],
      type: PatchType.PARAGRAPH,
    },

    todays_date: {
      children: [{ text: new Date().toLocaleDateString() }],
      type: PatchType.PARAGRAPH,
    },

    zip: {
      children: [{ text: "test" }],
      type: PatchType.PARAGRAPH,
    },
  },
});
writeFileSync("My Document.docx", doc);
