import { writeFileSync } from "node:fs";

import { Workbook, Packer } from "@office-open/xlsx";

const wb = new Workbook({
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
    },
    {
      name: "Unprotected",
      rows: [{ cells: [{ value: "This sheet is not protected" }] }],
    },
  ],
});

const buffer = await Packer.toBuffer(wb);
writeFileSync("My Workbook.xlsx", buffer);
