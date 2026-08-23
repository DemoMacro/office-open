// Rich metadata part (xl/metadata.xml) — types, strings, and future metadata.
// cellMetadata/valueMetadata blocks are intentionally not exercised: Excel 365
// refuses the file outright when they appear fresh (with or without cell-level
// cm/vm references) — verified empirically; they stay round-trip only.

import { mkdirSync, writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const buffer = await generateWorkbook({
  worksheets: [
    {
      name: "Data",
      rows: [
        {
          cells: [
            { reference: "A1", value: "Widget" },
            { reference: "B1", value: 9.99 },
          ],
        },
      ],
    },
  ],
  metadata: {
    types: [
      {
        name: "XLDAPROPERTY",
        minSupportedVersion: 1,
        copy: true,
        delete: true,
        edit: true,
        pasteAll: true,
      },
    ],
    strings: [{ value: "Product catalog" }],
    futureMetadata: [{ name: "XLDAPROPERTY", blocks: [{}] }],
  },
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/30-metadata.xlsx", buffer);
