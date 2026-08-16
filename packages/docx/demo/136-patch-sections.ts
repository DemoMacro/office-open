// Patch a document: replace a section's content and append new sections.

import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument, patchDocument } from "@office-open/docx";

// Step 1: Generate a two-section template. The first section ends with a
//         section-break paragraph carrying its own page setup; the second
//         section is closed by the body-level sectPr.
const templateBuffer = await generateDocument({
  title: "Patch Sections Demo",
  sections: [
    {
      children: [{ paragraph: { children: ["Chapter 1 — placeholder content."] } }],
      properties: { pageSize: { width: 11906, height: 16838 } },
    },
    {
      children: [{ paragraph: { children: ["Chapter 2 — stays as-is."] } }],
    },
  ],
});

// Step 2: Patch — swap section 0's content (keeping its page setup since no
//         properties are given) and append a third section with its own
//         landscape page setup, inserted before the final body sectPr.
const patched = await patchDocument({
  outputType: "nodebuffer",
  data: templateBuffer,
  sections: {
    replace: [
      {
        index: 0,
        section: {
          children: [
            { paragraph: { heading: "Heading1", children: ["Chapter 1 — replaced"] } },
            { paragraph: { children: ["Content patched in over the original section."] } },
          ],
        },
      },
    ],
    append: [
      {
        children: [
          { paragraph: { heading: "Heading1", children: ["Appendix — appended section"] } },
          { paragraph: { children: [{ text: "Landscape page.", italic: true }] } },
        ],
        properties: { pageSize: { width: 16838, height: 11906 } },
      },
    ],
  },
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/136-patch-sections.docx", patched);
