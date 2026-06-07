import { writeFileSync } from "node:fs";

import { generate } from "@office-open/xlsx";

const buffer = await generate({
  worksheets: [
    {
      name: "Hyperlinks",
      rows: [
        { cells: [{ value: "External Link" }, { value: "Internal Link" }] },
        { cells: [{ value: "Google" }, { value: "Go to Data sheet" }] },
        { cells: [{ value: "GitHub" }] },
      ],
      hyperlinks: [
        {
          cell: "A2",
          target: { type: "external", url: "https://www.google.com" },
          tooltip: "Open Google",
        },
        {
          cell: "B2",
          target: { type: "internal", location: "Data!A1" },
          tooltip: "Jump to Data sheet",
        },
        {
          cell: "A3",
          target: { type: "external", url: "https://github.com" },
          display: "GitHub Repo",
        },
      ],
    },
    {
      name: "Data",
      rows: [
        { cells: [{ value: "Name" }, { value: "Value" }] },
        { cells: [{ value: "Item A" }, { value: 100 }] },
        { cells: [{ value: "Item B" }, { value: 200 }] },
      ],
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
