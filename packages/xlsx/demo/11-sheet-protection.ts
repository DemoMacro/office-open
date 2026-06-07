import { writeFileSync } from "node:fs";

import { generate } from "@office-open/xlsx";

const buffer = await generate({
  workbookProtection: {
    lockStructure: true,
    lockWindows: true,
    workbookPassword: "admin",
  },
  worksheets: [
    {
      name: "Protected",
      rows: [
        { cells: [{ value: "Quarter" }, { value: "2023" }, { value: "2024" }] },
        { cells: [{ value: "Q1" }, { value: 100 }, { value: 120 }] },
        { cells: [{ value: "Q2" }, { value: 150 }, { value: 170 }] },
        { cells: [{ value: "Q3" }, { value: 200 }, { value: 220 }] },
        { cells: [{ value: "Q4" }, { value: 180 }, { value: 200 }] },
      ],
      protection: {
        sheet: true,
        password: "secret",
        formatCells: false,
        insertRows: false,
        deleteColumns: false,
        sort: false,
        selectLockedCells: true,
        selectUnlockedCells: true,
      },
      pageSetup: {
        fitToWidth: 1,
        fitToHeight: 0,
        autoPageBreaks: true,
      },
      protectedRanges: [
        {
          name: "EditableRange",
          sqref: "B2:C5",
          password: "range1",
        },
      ],
    },
    {
      name: "ModernCrypto",
      rows: [
        { cells: [{ value: "Item" }, { value: "Value" }] },
        { cells: [{ value: "A" }, { value: 10 }] },
        { cells: [{ value: "B" }, { value: 20 }] },
      ],
      protection: {
        sheet: true,
        algorithmName: "SHA-512",
        hashValue: "dGVzdA==",
        saltValue: "c2FsdA==",
        spinCount: 100000,
        formatCells: false,
        formatColumns: false,
        formatRows: false,
      },
    },
    {
      name: "Unprotected",
      rows: [{ cells: [{ value: "This sheet is not protected" }] }],
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
