// Append canonical section children to an existing document.

import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument, patchDocument } from "@office-open/docx";

const template = await generateDocument({
  title: "Append Demo",
  sections: [
    {
      children: [{ paragraph: { children: ["This is the original body content."] } }],
    },
  ],
});

const patched = await patchDocument({
  data: template,
  commands: [
    {
      type: "insert",
      at: { field: "sections", route: [{ index: 0 }, { collection: "children" }] },
      position: "end",
      value: [
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
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/132-patch-append.docx", patched);
