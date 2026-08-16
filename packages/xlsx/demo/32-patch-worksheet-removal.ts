// Patch a workbook: remove worksheets from a template workbook.

import { mkdirSync, writeFileSync } from "node:fs";

import { generateWorkbook, patchWorkbook } from "@office-open/xlsx";

// Step 1: Create a three-sheet template — a data sheet plus two report
//         variants, not all of which every consumer needs. SalesTotal is
//         scoped to the final sheet (localSheetId 2).
const templateBuffer = await generateWorkbook({
  title: "Patch Worksheet Removal Demo",
  worksheets: [
    {
      name: "Data",
      rows: [
        {
          cells: [
            { reference: "A1", value: "Region" },
            { reference: "B1", value: "Sales" },
          ],
        },
        {
          cells: [
            { reference: "A2", value: "North" },
            { reference: "B2", value: 120 },
          ],
        },
      ],
    },
    { name: "Report Draft", rows: [{ cells: [{ reference: "A1", value: "Draft summary" }] }] },
    { name: "Report Final", rows: [{ cells: [{ reference: "A1", value: "Final summary" }] }] },
  ],
  definedNames: [{ name: "SalesTotal", value: "Data!B2", localSheetId: 2 }],
});

// Step 2: Patch — drop the draft sheet. definedName entries scoped to a
//         removed sheet are dropped; survivors are renumbered to the
//         post-removal sheet indices.
const patched = await patchWorkbook({
  outputType: "nodebuffer",
  data: templateBuffer,
  worksheets: { remove: ["Report Draft"] },
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/32-patch-worksheet-removal.xlsx", patched);
