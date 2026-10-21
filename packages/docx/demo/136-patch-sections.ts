// Replace and append canonical sections.

import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument, patchDocument } from "@office-open/docx";

const template = await generateDocument({
  title: "Patch Sections Demo",
  sections: [
    {
      children: [{ paragraph: { children: ["Chapter 1 — placeholder content."] } }],
      properties: { pageSize: { width: 11906, height: 16838 } },
    },
    { children: [{ paragraph: { children: ["Chapter 2 — stays as-is."] } }] },
  ],
});

const patched = await patchDocument({
  data: template,
  commands: [
    {
      type: "replace",
      at: { field: "sections", route: [{ index: 0 }] },
      value: {
        children: [
          { paragraph: { heading: "Heading1", children: ["Chapter 1 — replaced"] } },
          { paragraph: { children: ["Content patched in over the original section."] } },
        ],
      },
    },
    {
      type: "insert",
      at: { field: "sections" },
      position: "end",
      value: [
        {
          children: [
            { paragraph: { heading: "Heading1", children: ["Appendix — appended section"] } },
            { paragraph: { children: [{ text: "Landscape page.", italic: true }] } },
          ],
          properties: { pageSize: { width: 16838, height: 11906 } },
        },
      ],
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/136-patch-sections.docx", patched);
