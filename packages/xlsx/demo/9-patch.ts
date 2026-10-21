import { mkdirSync, readFileSync, writeFileSync } from "node:fs";

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

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/9-patch.xlsx", buffer);

// Step 2: Patch placeholders, replace the "Notes" sheet, and append a "Summary".
//          Text targets keep operation order; worksheet replacement preserves
//          source wiring and the shared-strings table is rebuilt.
const patched = await patchWorkbook({
  outputType: "uint8array",
  data: readFileSync(".temp/9-patch.xlsx"),
  commands: [
    {
      type: "replace",
      at: {
        field: "worksheets",
        route: [
          { name: "Invoice" },
          { property: "rows" },
          { index: 0 },
          { property: "cells" },
          { index: 1 },
        ],
      },
      value: { value: "INV-2024-001" },
    },
    {
      type: "replace",
      at: {
        field: "worksheets",
        route: [
          { name: "Invoice" },
          { property: "rows" },
          { index: 1 },
          { property: "cells" },
          { index: 1 },
        ],
      },
      value: { value: "Acme Corp" },
    },
    {
      type: "replace",
      at: {
        field: "worksheets",
        route: [
          { name: "Invoice" },
          { property: "rows" },
          { index: 2 },
          { property: "cells" },
          { index: 1 },
        ],
      },
      value: { value: 1500 },
    },
    {
      type: "replace",
      at: {
        field: "worksheets",
        route: [
          { name: "Invoice" },
          { property: "rows" },
          { index: 3 },
          { property: "cells" },
          { index: 1 },
        ],
      },
      value: { value: new Date("2024-12-31") },
    },
    {
      type: "replace",
      at: { field: "worksheets", route: [{ name: "Notes" }] },
      value: { name: "Notes", rows: [{ cells: [{ value: "Replaced notes" }] }] },
    },
    {
      type: "insert",
      at: { field: "worksheets" },
      position: "end",
      value: [{ name: "Summary", rows: [{ cells: [{ value: "Appended summary sheet" }] }] }],
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/9-patch.xlsx", patched);
