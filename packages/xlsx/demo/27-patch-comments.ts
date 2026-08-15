import { mkdirSync, readFileSync, writeFileSync } from "node:fs";

import { generateWorkbook, patchWorkbook } from "@office-open/xlsx";

// Patch cell comments (notes) onto an existing worksheet. comments keys are
// sheet names (as declared in workbook.xml); existing comments on a worksheet
// are merged — parsed, extended, and re-serialized (authors re-collected).
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
  comments: {
    Report: [
      { cell: "B2", author: "Alice", text: "Above target" },
      { cell: "B3", author: "Bob", text: "Needs follow-up" },
    ],
  },
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/27-patch-comments.xlsx", patched);
