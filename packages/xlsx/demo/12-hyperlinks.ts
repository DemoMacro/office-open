import { mkdirSync, writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const buffer = await generateWorkbook({
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
          url: "https://www.google.com",
          tooltip: "Open Google",
        },
        {
          cell: "B2",
          location: "Data!A1",
          tooltip: "Jump to Data sheet",
        },
        {
          cell: "A3",
          url: "https://github.com",
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

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/12-hyperlinks.xlsx", buffer);
