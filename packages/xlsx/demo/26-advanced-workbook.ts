// Advanced workbook options: custom views, file recovery, file sharing, function groups.

import { writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const buffer = await generateWorkbook({
  worksheets: [
    {
      name: "Data",
      rows: [
        { cells: [{ value: "Item" }, { value: "Price" }] },
        { cells: [{ value: "Widget" }, { value: 9.99 }] },
      ],
    },
  ],
  customWorkbookViews: [
    {
      name: "CustomView1",
      guid: "{12345678-1234-1234-1234-123456789ABC}",
      windowWidth: 20000,
      windowHeight: 12000,
      activeSheetId: 1,
      showFormulaBar: true,
      showStatusbar: true,
      showHorizontalScroll: true,
      showVerticalScroll: true,
      showSheetTabs: true,
    },
  ],
  fileRecoveryPr: {
    autoRecover: true,
    crashSave: true,
    dataExtractLoad: false,
    repairLoad: false,
  },
  functionGroups: ["MyCustomFunctions", "AnalysisToolPak"],
  fileSharing: {
    readOnlyRecommended: true,
    userName: "Alice",
    reservationPassword: "ABCD",
  },
  webPublishing: {
    css: true,
    thicket: true,
    longFileNames: true,
    vml: false,
    targetScreenSize: "800x600",
    dpi: 96,
    codePage: 65001,
  },
});

writeFileSync("My Workbook.xlsx", buffer);
