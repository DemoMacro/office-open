import { writeFileSync, readFileSync } from "node:fs";

import { generateWorkbook, patchWorkbook } from "@office-open/xlsx";

// Step 1: Create a template workbook with placeholders
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
  ],
});

writeFileSync("My Workbook.xlsx", buffer);

// Step 2: Patch the placeholders
const patched = await patchWorkbook({
  outputType: "uint8array",
  data: readFileSync("My Workbook.xlsx"),
  placeholders: {
    number: "INV-2024-001",
    customer: "Acme Corp",
    amount: 1500, // number → typed numeric cell (t="n")
    date: new Date("2024-12-31"), // Date → numeric serial cell
  },
});

writeFileSync("My Workbook.xlsx", patched);
