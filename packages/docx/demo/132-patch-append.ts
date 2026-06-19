// Patch a document: append block-level content before the final section break.

import { writeFileSync } from "node:fs";

import { generateDocument, patchDocument } from "@office-open/docx";

// Step 1: Generate a template document. The body ends with a final section
//         break (<w:sectPr>); appended content is spliced in just before it.
const templateBuffer = await generateDocument({
  title: "Append Demo",
  sections: [
    {
      children: [{ paragraph: { children: ["This is the original body content."] } }],
    },
  ],
});

// Step 2: Patch — append paragraphs and a table. Appended content reuses the
//         SectionChild vocabulary (paragraph / table / ...), serialized through
//         the same compile path as generateDocument.
const patched = await patchDocument({
  outputType: "nodebuffer",
  data: templateBuffer,
  append: [
    { paragraph: { children: ["Appended paragraph one."] } },
    { paragraph: { children: [{ text: "Appended bold paragraph.", bold: true }] } },
    {
      table: {
        columnWidths: [3000, 3000],
        rows: [
          {
            cells: [
              { children: [{ paragraph: "Left cell" }] },
              { children: [{ paragraph: "Right cell" }] },
            ],
          },
        ],
      },
    },
  ],
});

writeFileSync("My Document.docx", patched);
