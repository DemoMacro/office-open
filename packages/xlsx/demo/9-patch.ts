import { readFileSync, writeFileSync } from "node:fs";

import { generateWorkbook, patchWorkbook } from "@office-open/xlsx";

// Step 1: Create a template workbook with placeholders and two worksheets
const buffer = await generateWorkbook({
  title: "Template",
  worksheets: [
    {
      name: "Invoice",
      rows: [
        { cells: [{ value: "Invoice" }, { value: "{{number}}" }] },
        { cells: [{ value: "Customer" }, { value: "{{customer}}" }] },
        { cells: [{ value: "Amount" }, { value: "{{amount}}" }] },
        { cells: [{ value: "Date" }, { value: "{{date}}" }] },
      ],
    },
    {
      name: "Notes",
      rows: [{ cells: [{ value: "Original notes" }] }],
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);

// Step 2: Patch placeholders, replace the "Notes" sheet, and append a "Summary".
//          worksheets.replace keys are sheet names (as declared in workbook.xml);
//          worksheets.append adds sheets after the last existing one. The
//          shared-strings table is rebuilt so indexing continues correctly.
const patched = await patchWorkbook({
  outputType: "uint8array",
  data: readFileSync("My Workbook.xlsx"),
  placeholders: {
    number: "INV-2024-001",
    customer: "Acme Corp",
    amount: 1500, // number → typed numeric cell (t="n")
    date: new Date("2024-12-31"), // Date → numeric serial cell
  },
  worksheets: {
    replace: {
      Notes: { rows: [{ cells: [{ value: "Replaced notes" }] }] },
    },
    append: [{ name: "Summary", rows: [{ cells: [{ value: "Appended summary sheet" }] }] }],
  },
});

writeFileSync("My Workbook.xlsx", patched);

// Step 3: Append cell comments to the Invoice worksheet. Existing comments on a
//          worksheet are merged — parsed and re-serialized with the new notes.
//          comments keys are sheet names (as declared in workbook.xml).
const withComments = await patchWorkbook({
  outputType: "uint8array",
  data: readFileSync("My Workbook.xlsx"),
  comments: {
    Invoice: [
      { cell: "B1", author: "Alice", text: "Invoice number confirmed" },
      { cell: "B3", author: "Bob", text: "Amount needs approval" },
    ],
  },
});

writeFileSync("My Workbook.xlsx", withComments);
