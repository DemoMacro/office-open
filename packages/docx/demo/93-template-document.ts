// Patch canonical document text in an existing template.

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";

import { patchDocument } from "@office-open/docx";

const replacements: Record<string, string> = {
  address: "blah blah",
  city: "test",
  email_address: "test",
  first_name: "test",
  ft_dates: "test",
  grade: "test",
  last_name: "test",
  phone: "test",
  school_name: "test",
  state: "test",
  todays_date: new Date().toLocaleDateString(),
  zip: "test",
};

const doc = await patchDocument({
  data: readFileSync("demo/assets/field-trip.docx"),
  commands: Object.entries(replacements).map(([placeholder, value]) => ({
    type: "replace",
    at: { field: "sections", text: { placeholder } },
    value,
  })),
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/93-template-document.docx", doc);
