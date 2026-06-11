import { readFileSync, writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const dogPng = readFileSync("./demo/images/dog.png");
const catJpg = readFileSync("./demo/images/cat.jpg");

const buffer = await generateWorkbook({
  worksheets: [
    {
      name: "Images",
      rows: [
        { cells: [{ value: "Name" }, { value: "Type" }] },
        { cells: [{ value: "Dog" }, { value: "PNG" }] },
        { cells: [{ value: "Cat" }, { value: "JPEG" }] },
      ],
      images: [
        { data: dogPng, type: "png", col: 3, row: 1 },
        { data: catJpg, type: "jpeg", col: 3, row: 3 },
      ],
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
