// Replace canonical paragraphs to add inline comments.

import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument, patchDocument } from "@office-open/docx";

const buffer = await generateDocument({
  title: "Comment Patch Demo",
  sections: [
    {
      children: [
        { paragraph: { children: ["First paragraph of the document."] } },
        { paragraph: { children: ["Second paragraph with a {{name}} placeholder."] } },
      ],
    },
  ],
});

const patched = await patchDocument({
  data: buffer,
  commands: [
    {
      type: "replace",
      at: { field: "sections", route: [{ index: 0 }, { property: "children" }, { index: 0 }] },
      value: {
        paragraph: {
          children: [
            {
              comment: {
                author: "Alice",
                children: ["Comment anchored to the first paragraph."],
                wrap: ["First paragraph of the document."],
              },
            },
          ],
        },
      },
    },
    {
      type: "replace",
      at: { field: "sections", route: [{ index: 0 }, { property: "children" }, { index: 1 }] },
      value: {
        paragraph: {
          children: [
            {
              comment: {
                author: "Bob",
                children: ["Comment anchored to the placeholder run."],
                wrap: ["reviewer"],
              },
            },
          ],
        },
      },
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/133-patch-comments.docx", patched);
