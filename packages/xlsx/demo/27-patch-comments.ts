import { mkdirSync, readFileSync, writeFileSync } from "node:fs";

import { generateWorkbook, patchWorkbook } from "@office-open/xlsx";

// Patch cell comments (notes) onto an existing worksheet. Existing comments on
// a worksheet are merged — parsed, extended, and re-serialized.
const buffer = await generateWorkbook({
  title: "Comment Patch Demo",
  worksheets: [
    {
      name: "Report",
      rows: [
        { cells: [{ value: "Region" }, { value: "Sales" }] },
        { cells: [{ value: "North" }, { value: 1200 }] },
        { cells: [{ value: "South" }, { value: 980 }] },
      ],
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/27-patch-comments.xlsx", buffer);

const patched = await patchWorkbook({
  outputType: "uint8array",
  data: readFileSync(".temp/27-patch-comments.xlsx"),
  commands: [
    {
      type: "insert",
      at: { field: "worksheets", route: [{ name: "Report" }, { collection: "comments" }] },
      value: [
        { cell: "B2", author: "Alice", text: "Above target" },
        { cell: "B3", author: "Bob", text: "Needs follow-up" },
      ],
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/27-patch-comments.xlsx", patched);
