// Patch a document with patches

import { readFileSync, writeFileSync } from "node:fs";

import { patchDocument } from "@office-open/docx";

const doc = await patchDocument({
  data: readFileSync("demo/assets/field-trip.docx"),
  outputType: "nodebuffer",
  placeholders: {
    address: {
      children: [{ text: "blah blah" }],
      type: "paragraph",
    },

    city: {
      children: [{ text: "test" }],
      type: "paragraph",
    },

    email_address: {
      children: [{ text: "test" }],
      type: "paragraph",
    },

    first_name: {
      children: [{ text: "test" }],
      type: "paragraph",
    },

    ft_dates: {
      children: [{ text: "test" }],
      type: "paragraph",
    },

    grade: {
      children: [{ text: "test" }],
      type: "paragraph",
    },

    last_name: {
      children: [{ text: "test" }],
      type: "paragraph",
    },

    phone: {
      children: [{ text: "test" }],
      type: "paragraph",
    },

    school_name: {
      children: [{ text: "test" }],
      type: "paragraph",
    },

    state: {
      children: [{ text: "test" }],
      type: "paragraph",
    },

    todays_date: {
      children: [{ text: new Date().toLocaleDateString() }],
      type: "paragraph",
    },

    zip: {
      children: [{ text: "test" }],
      type: "paragraph",
    },
  },
});
writeFileSync("My Document.docx", doc);
