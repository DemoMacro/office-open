// Rich metadata part (xl/metadata.xml) with cell/value metadata references.

import { mkdirSync, writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const buffer = await generateWorkbook({
  worksheets: [
    {
      name: "Data",
      rows: [
        {
          cells: [
            { reference: "A1", value: "Widget", cellMetadataId: 1, valueMetadataId: 1 },
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
    cellMetadata: [{ records: [{ typeIndex: 0, valueIndex: 0 }] }],
    valueMetadata: [{ records: [{ typeIndex: 0, valueIndex: 0 }] }],
  },
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/30-metadata.xlsx", buffer);
