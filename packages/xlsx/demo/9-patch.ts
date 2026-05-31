import { writeFileSync, readFileSync } from "node:fs";

import { Workbook, Packer, patchWorkbook } from "@office-open/xlsx";

// Step 1: Create a template workbook with placeholders
const template = new Workbook({
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

const buffer = await Packer.toBuffer(template);
writeFileSync("My Workbook.xlsx", buffer);

// Step 2: Patch the placeholders
const patched = await patchWorkbook({
  outputType: "uint8array",
  data: readFileSync("My Workbook.xlsx"),
  patches: {
    number: { value: "INV-2024-001" },
    customer: { value: "Acme Corp" },
    amount: { value: "$1,500.00" },
    date: { value: "2024-12-31" },
  },
});

writeFileSync("My Workbook.xlsx", patched);
