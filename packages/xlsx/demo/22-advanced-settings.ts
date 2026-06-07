// Demo: Ignored errors, phonetic properties, sheet background, file sharing,
// file recovery, web publishing, custom workbook views, and function groups.

import { writeFileSync } from "node:fs";

import { generate } from "@office-open/xlsx";

const buffer = await generate({
  fileSharing: {
    readOnlyRecommended: true,
    userName: "Analyst",
  },
  fileRecoveryPr: {
    autoRecover: true,
    crashSave: true,
  },
  webPublishing: {
    css: true,
    allowPng: true,
    dpi: 150,
    targetScreenSize: "1024x768",
  },
  functionGroups: ["UDF1", "CustomFunctions"],
  customWorkbookViews: [
    {
      name: "Summary View",
      guid: "{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}",
      windowWidth: 28800,
      windowHeight: 15000,
      activeSheetId: 1,
      showFormulaBar: true,
      tabRatio: 500,
    },
  ],
  worksheets: [
    {
      name: "Data",
      rows: [
        {
          cells: [{ value: "Product" }, { value: "Price" }, { value: "Qty" }, { value: "Total" }],
        },
        { cells: [{ value: "Widget" }, { value: "19.99" }, { value: 100 }, { value: 1999 }] },
        { cells: [{ value: "Gadget" }, { value: "49.95" }, { value: 50 }, { value: 2497.5 }] },
        { cells: [{ value: "Doohickey" }, { value: "5.00" }, { value: 200 }, { value: 1000 }] },
      ],
      // Ignored errors: suppress "number stored as text" for the price column
      ignoredErrors: [
        { sqref: "B2:B4", numberStoredAsText: true },
        { sqref: "D2:D4", evalError: true },
      ],
      // Phonetic properties for CJK text (Japanese furigana)
      phoneticPr: {
        fontId: 0,
        type: "Hiragana",
        alignment: "left",
      },
      // Sheet background image
      backgroundImage: {
        data: new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), // minimal PNG header
        type: "png",
      },
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
